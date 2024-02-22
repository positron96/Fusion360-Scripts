# Author-Amanda Ghassaei
# Description-Turn your Fusion360 design history timeline into an animation

import traceback
import math
import os
import typing

import adsk
from adsk import core, fusion

FOps = fusion.FeatureOperations

app = core.Application.get()
if not app:
    raise RuntimeError('No Fusion application!')

ui = app.userInterface

# Global set of event handlers to keep them referenced for the duration of the command
handlers = []

# Keep the timelapse object in global namespace.
timelapse = None


class CommandExecuteHandler(core.CommandEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args: core.CommandEventArgs):
        try:
            unitsMgr = app.activeProduct.unitsManager
            command = args.firingEvent.sender
            inputs = typing.cast(core.CommandInputs, command.commandInputs)

            for input in inputs:
                iid = input.id
                if iid.endswith('range'):
                    if (not hasattr(timelapse, 'start')
                        or not hasattr(timelapse, 'end')
                    ):
                        raise TypeError('Unknown field: "{}"'.format(iid))
                    setattr(timelapse, 'start', input.valueOne)
                    setattr(timelapse, 'end', input.valueTwo)
                else:
                    if not hasattr(timelapse, iid):
                        raise TypeError('Unknown field: "{}"'.format(input.id))
                    setattr(timelapse, iid, input.value)

            timelapse.collectFrames()

            args.isValidResult = True
        except Exception:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


class CommandDestroyHandler(core.CommandEventHandler):

    def __init__(self):
        super().__init__()

    def notify(self, args):
        try:
            # When the command is done, terminate the script.
            # This will release all globals which will remove all event handlers.
            adsk.terminate()
        except Exception:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


class CommandCreatedHandler(core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args: core.CommandCreatedEventArgs):
        try:
            cmd = args.command
            cmd.isRepeatable = False
            onExecute = CommandExecuteHandler()
            cmd.execute.add(onExecute)
            onDestroy = CommandDestroyHandler()
            cmd.destroy.add(onDestroy)
            # Keep the handler referenced beyond this function.
            handlers.append(onExecute)
            handlers.append(onDestroy)

            max_int = 2147483647

            # Define the inputs.
            inputs = cmd.commandInputs
            # File params.
            inputs.addStringValueInput('foldername', 'Folder name', timelapse.foldername)
            inputs.addStringValueInput('outputPath', 'Output path', timelapse.outputPath)
            inputs.addBoolValueInput('saveObj', 'Save .obj files', True, '', timelapse.saveObj)
            inputs.addIntegerSpinnerCommandInput('width', 'Images width', 1, max_int, 1, timelapse.width)
            inputs.addIntegerSpinnerCommandInput('height', 'Images height', 1, max_int, 1, timelapse.height)
            # Animation params.
            inputs.addIntegerSliderCommandInput('range', 'Timeline range', 1, timelapse.timeline.count, True)
            inputs.itemById('range').valueOne = timelapse.start
            inputs.itemById('range').valueTwo = timelapse.end
            inputs.addIntegerSpinnerCommandInput(
                'interpolationFrames', 'Frames per operation', 1, max_int, 1, timelapse.interpolationFrames)
            inputs.addBoolValueInput('doFit', 'Fit design', True, '', timelapse.doFit)
            inputs.addBoolValueInput('doRotate', 'Rotate design', True, '', timelapse.doRotate)
            inputs.addIntegerSpinnerCommandInput(
                'framesPerRotation', 'Frames per rotation', 1, max_int, 1, timelapse.framesPerRotation)
            inputs.addIntegerSpinnerCommandInput(
                'finalFrames', 'Num final frames', 0, max_int, 1, timelapse.finalFrames)

        except Exception:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


class HistoryTimelapse:
    def __init__(self, design: fusion.Design):
        dataFile = app.activeDocument.dataFile

        # Set initial values.
        self.foldername = dataFile.name
        self.outputPath = os.path.expanduser('~')
        self.saveObj = False
        self._timeline = design.timeline
        self._design = design
        self.width = 2000
        self.height = 2000
        self.start = 1
        self.end = self._timeline.markerPosition
        self.interpolationFrames = 5
        self.doRotate = True
        self.framesPerRotation = 500
        self._finalFrames = 0

        self.doFit = False

    # Properties.
    @property
    def timeline(self):
        return self._timeline

    @property
    def finalFrames(self):
        return self._finalFrames

    @finalFrames.setter
    def finalFrames(self, value):
        if value < 0:
            value = 0
        self._finalFrames = value

    @property
    def design(self):
        return self._design

    def isNumericExtent(self, extent: fusion.ExtentDefinition):
        return isinstance(extent, (
            fusion.DistanceExtentDefinition,
            fusion.SymmetricExtentDefinition,
            fusion.AngleExtentDefinition))

    def collectFrames(self):
        start = self.start - 1 # Zero index the start value.
        end = self.end
        width = self.width
        height = self.height
        foldername = self.foldername
        outputPath = os.path.join(self.outputPath, foldername)
        saveObj = self.saveObj
        timeline = self.timeline
        documents = app.documents
        interpolationFrames = self.interpolationFrames
        framesPerRotation = self.framesPerRotation if self.doRotate else 0
        finalFrames = self.finalFrames
        doFit = self.doFit

        viewport = app.activeViewport
        if doFit:
            viewport.fit()
        camera = viewport.camera

        up = viewport.frontUpDirection
        rot = core.Matrix3D.create()
        rot.setToRotation(
            math.pi * 2.0 / framesPerRotation,
            up,
            core.Point3D.create(0, 0, 0))

        frame_num = 0
        startingAngle = None

        os.makedirs(outputPath, exist_ok=True)
        ff = open(os.path.join(outputPath, 'log.txt'), 'w')

        for timeline_pos in range(start, end):
            try:
                # Get feature at current timeline index.
                item = timeline.item(timeline_pos)

                # If item is suppressed, ignore.
                if item.isSuppressed:
                    continue
                # If item is group, ignore.
                if item.isGroup:
                    continue
                # # If item has an error, ignore.
                # TODO: Move is throwing error here.
                # if item.healthState == 2: # ErrorFeatureHealthState
                #     continue

                entity = item.entity
                # TODO: Move TimelineObject is not working properly here.
                if not entity:
                    continue
                classname = type(entity).__name__

                # Some operations should be ignored as they can't be easily animated.
                # TODO: feature handling list is not exhaustive, just what I frequently use.
                if classname == 'Occurrence':
                    continue
                    # # We only allow a fade in for Occurrence if this is the first item inserted into the design.
                    # if i > 0:
                    #     continue
                    # # And if insert operation not immediately folled by a joint.
                    # nextItem = type(timeline.item(i + 1).entity).__name__
                    # if nextItem == 'Joint':
                    #     continue
                # If item is sketch, ignore.
                elif classname == 'Sketch':
                    continue
                elif classname == 'ConstructionPlane':
                    continue
                elif classname == 'ConstructionAxis':
                    continue
                elif classname == 'ConstructionPoint':
                    continue
                elif classname == 'ThreadFeature':
                    continue
                elif classname == 'Combine':
                    continue
                elif classname == 'Canvas':
                    continue
            except Exception:
                continue

            # Set marker position.
            timeline.markerPosition = timeline_pos + 1

            # Get parameters to interpolate.
            interpolatedParameters = []
            stepSizes = []
            stepOffsets = []
            alphaComponents = []
            if isinstance(entity, fusion.ExtrudeFeature):
                numeric = False
                if self.isNumericExtent(entity.extentOne):
                    param = entity.extentOne.distance
                    interpolatedParameters.append(param)
                    stepSizes.append(param.value / interpolationFrames)
                    stepOffsets.append(0)
                    numeric = True
                if entity.hasTwoExtents and self.isNumericExtent(entity.extentTwo):
                    # Handle side 2.
                    param = entity.extentTwo.distance
                    interpolatedParameters.append(param)
                    stepSizes.append(param.value / interpolationFrames)
                    stepOffsets.append(0)
                    numeric = True
                # At the very least we can fade it in if it's a new body/component.
                if (
                    not numeric
                    and entity.operation in {
                        FOps.NewBodyFeatureOperation,
                        FOps.NewComponentFeatureOperation}
                ):
                    bodies = entity.bodies
                    for body in bodies:
                        alphaComponents.append(body)
            # if classname == 'OffsetFacesFeature': # TODO: unable to get extent parameter from this operation.
            if classname == 'Move':
                # TODO: implement this.
                try:
                    print(timeline_pos, 'move:', str(entity.transform.asArray()), file=ff)
                except Exception as e:
                    print(timeline_pos, 'move:', str(e), file=ff)
                continue
            elif classname == 'MirrorFeature':
                try:
                    print(timeline_pos, 'mirror: bodies', entity.bodies.count, file=ff)
                    print(timeline_pos, 'mirror: inputEntities', entity.inputEntities.count, file=ff)
                except Exception as e:
                    print(timeline_pos, 'mirror:', str(e), file=ff)
                continue
            elif isinstance(entity, fusion.RevolveFeature):
                if self.isNumericExtent(entity.extentDefinition):
                    param = entity.extentDefinition.angle
                    interpolatedParameters.append(param)
                    stepSizes.append(param.value / interpolationFrames)
                    # ui.messageBox(str(param.value / interpolationFrames))
                    stepOffsets.append(0)
            # if classname == 'FilletFeature' or classname == 'ChamferFeature': # TODO: unable to get extent parameter from this operation.
            elif classname == 'Joint':
                # This usually happens after occurrence, fade in the component.
                # Need to break link first so that opacity changes aren't
                # applied to all linked components - this is irreversible.
                # TODO: If the occurrence is not the top level component of the linked component, this may throw an error.
                if entity.occurrenceOne:
                    try:
                        if entity.occurrenceOne.isReferencedComponent:
                            entity.occurrenceOne.breakLink()
                    except Exception:
                        pass
                    alphaComponents.append(entity.occurrenceOne.component)
            elif classname == 'Occurrence':
                try:
                    if entity.isReferencedComponent:
                        entity.breakLink()
                except Exception:
                    pass
                alphaComponents.append(entity.component)
            elif isinstance(entity, fusion.RectangularPatternFeature):
                if entity.quantityOne:
                    param = entity.quantityOne
                    if param.value != 1:
                        interpolatedParameters.append(param)
                        stepSize = int(param.value / interpolationFrames)
                        if stepSize < 1:
                            stepSize = 1
                        stepSizes.append(stepSize)
                        stepOffsets.append(0)
                        if entity.distanceOne:
                            dist = entity.distanceOne
                            interpolatedParameters.append(dist)
                            distStepSize = dist.value / (param.value - 1)
                            # ui.messageBox(str(distStepSize))
                            stepSizes.append(distStepSize * stepSize)
                            stepOffsets.append(-distStepSize)
                if entity.quantityTwo:
                    param = entity.quantityTwo
                    if param.value != 1:
                        interpolatedParameters.append(param)
                        stepSize = int(param.value / interpolationFrames)
                        stepSizes.append(stepSize)
                        stepOffsets.append(0)
                        if entity.distanceTwo:
                            dist = entity.distanceTwo
                            interpolatedParameters.append(dist)
                            distStepSize = dist.value / (param.value - 1)
                            stepSizes.append(distStepSize * stepSize)
                            stepOffsets.append(-distStepSize)

            # Save original values and expressions.
            originalValues = [param.value for param in interpolatedParameters]
            originalExpressions = [param.expression for param in interpolatedParameters]
            originalAlphas = [comp.opacity for comp in alphaComponents]
            # Calc number of interpolation frames for this feature.

            _interpolationFrames = interpolationFrames if (timeline_pos < end - 1) else (interpolationFrames + finalFrames)
            for step in range(_interpolationFrames):
                # Interpolate parameters.
                for k in range(len(interpolatedParameters)):
                    try:
                        value = stepSizes[k] * (step + 1) + stepOffsets[k]
                        if abs(value) > abs(originalValues[k]):
                            value = originalValues[k]
                        # ui.messageBox(str(value))
                        interpolatedParameters[k].value = value
                        # Force a recompute (needed for symmetric Revolves for some reason?).
                        if isinstance(entity, fusion.RevolveFeature):
                            entity.extentDefinition.isSymmetric = entity.extentDefinition.isSymmetric
                    except RuntimeError:
                        # modifying extents may fail with
                        # e.g. "No body to cut" error
                        continue
                for k in range(len(alphaComponents)):
                    value = originalAlphas[k] * (step + 1) / interpolationFrames
                    if abs(value) > abs(originalAlphas[k]):
                        value = originalAlphas[k]
                    # ui.messageBox(str(value))
                    alphaComponents[k].opacity = value

                # Rotate camera around y axis.
                if framesPerRotation > 0:
                    camera = viewport.camera
                    # camera.isSmoothTransition = False
                    if doFit:
                        camera.isFitView = True
                    pt = camera.eye.copy()
                    pt.transformBy(rot)
                    camera.eye = pt
                    camera.upVector = up
                    # Set camera property to trigger update.
                    viewport.camera = camera

                # Save image.
                outputFilename = os.path.join(
                    outputPath,
                    'frame_{:05d}'.format(frame_num))
                success = app.activeViewport.saveAsImageFile(
                    outputFilename + '.png', width, height)
                if not success:
                    ui.messageBox('Failed saving viewport image.')
                    break

                # Save obj file if requested
                if saveObj:
                    success = self.saveObjFile(outputFilename + '.obj')
                    if not success:
                        ui.messageBox('Failed saving obj file.')
                        break

                frame_num += 1

            # Reset parameters.
            for k in range(len(interpolatedParameters)):
                interpolatedParameters[k].value = originalValues[k]
                interpolatedParameters[k].expression = originalExpressions[k]
            for k in range(len(alphaComponents)):
                alphaComponents[k].opacity = originalAlphas[k]

    def saveObjFile(self, file):
        '''Export an .obj file from the root component'''
        try:
            adsk.doEvents()
            bodies = []
            comp = self.design.rootComponent
            for body in comp.bRepBodies:
                bodies.append(body)
            for occurrence in comp.allOccurrences:
                for body in occurrence.bRepBodies:
                    bodies.append(body)

            meshes = []
            for body in bodies:
                mesher = body.meshManager.createMeshCalculator()
                mesher.setQuality(
                    fusion.TriangleMeshQualityOptions.NormalQualityTriangleMesh
                )
                mesh = mesher.calculate()
                meshes.append(mesh)

            triangle_count = 0
            vert_count = 0
            for mesh in meshes:
                triangle_count += mesh.triangleCount
                vert_count += mesh.nodeCount

            # Write the mesh to OBJ
            with open(file, 'w') as fh:
                fh.write('# WaveFront *.obj file\n')
                fh.write(f'# Vertices: {vert_count}\n')
                fh.write(f'# Triangles : {triangle_count}\n\n')

                for mesh in meshes:
                    verts = mesh.nodeCoordinates
                    for pt in verts:
                        fh.write(f'v {pt.x} {pt.y} {pt.z}\n')
                for mesh in meshes:
                    for vec in mesh.normalVectors:
                        fh.write(f'vn {vec.x} {vec.y} {vec.z}\n')

                index_offset = 0
                for mesh in meshes:
                    mesh_tri_count = mesh.triangleCount
                    indices = mesh.nodeIndices
                    for t in range(mesh_tri_count):
                        i0 = indices[t * 3] + 1 + index_offset
                        i1 = indices[t * 3 + 1] + 1 + index_offset
                        i2 = indices[t * 3 + 2] + 1 + index_offset
                        fh.write(f'f {i0}//{i0} {i1}//{i1} {i2}//{i2}\n')
                    index_offset += mesh.nodeCount

                fh.write(f'\n# End of file')
                return True

        except Exception as ex:
            return False


def run(context):
    global timelapse
    try:
        ui.messageBox(
            'WARNING: This script will make changes to your file'
            ' (e.g. break links to referenced components).'
            '  You may want to run this on a copy of your design.'
            '  You can quit Fusion now to stop the script if you'
            ' are unsure about continuing.')
        product = app.activeProduct
        design = fusion.Design.cast(product)
        if not design:
            ui.messageBox(
                'Script is not supported in current workspace,'
                ' please change to MODEL workspace and try again.')
            return
        # Init a timelapse object.
        if timelapse is None:
            timelapse = HistoryTimelapse(design)

        commandDefinitions = ui.commandDefinitions
        # Check the command exists or not.
        cmdDef = commandDefinitions.itemById('designhistoryanimation')
        if not cmdDef:
            cmdDef = commandDefinitions.addButtonDefinition(
                'designhistoryanimation',
                'Design History Animation',
                'Turn your Fusion360 design history into an animation')

        onCommandCreated = CommandCreatedHandler()
        cmdDef.commandCreated.add(onCommandCreated)
        # Keep the handler referenced beyond this function.
        handlers.append(onCommandCreated)
        inputs = core.NamedValues.create()
        cmdDef.execute(inputs)

        # Prevent this module from being terminated when the script returns,
        # because we are waiting for event handlers to fire.
        adsk.autoTerminate(False)

    except Exception:
        ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
