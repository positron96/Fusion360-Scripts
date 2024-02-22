# Author-Amanda Ghassaei
# Description-Turn your Fusion360 design history timeline into an animation

import traceback
import math
import os

import adsk
from adsk import core, fusion

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

    def notify(self, args: core.CommandCreatedEventArgs):
        try:
            unitsMgr = app.activeProduct.unitsManager
            command = args.firingEvent.sender
            inputs = command.commandInputs

            for input in inputs:
                if input.id == 'filename':
                    timelapse.filename = input.value
                elif input.id == 'outputPath':
                    timelapse.outputPath = input.value
                elif input.id == 'saveObj':
                    timelapse.saveObj = input.value
                elif input.id == 'width':
                    timelapse.width = input.value
                elif input.id == 'height':
                    timelapse.height = input.value
                elif input.id == 'range':
                    timelapse.start = input.valueOne
                    timelapse.end = input.valueTwo
                elif input.id == 'interpolationFrames':
                    timelapse.interpolationFrames = input.value
                elif input.id == 'rotate':
                    timelapse.rotate = input.value
                elif input.id == 'framesPerRotation':
                    timelapse.framesPerRotation = input.value
                elif input.id == 'finalFrames':
                    timelapse.finalFrames = input.value

            timelapse.collectFrames()

            args.isValidResult = True
        except Exception:
            if ui:
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
            if ui:
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
            inputs.addStringValueInput('foldername', 'Folder Name', timelapse.foldername)
            inputs.addStringValueInput('outputPath', 'Output Path', timelapse.outputPath)
            inputs.addBoolValueInput('saveObj', 'Save .obj Files', True, '', timelapse.saveObj)
            inputs.addIntegerSpinnerCommandInput('width', 'Image Width (px)', 1, max_int, 1, timelapse.width)
            inputs.addIntegerSpinnerCommandInput('height', 'Image Height (px)', 1, max_int, 1, timelapse.height)
            # Animation params.
            inputs.addIntegerSliderCommandInput('range', 'Timeline Range', 1, timelapse.timeline.count, True)
            inputs.itemById('range').valueOne = timelapse.start
            inputs.itemById('range').valueTwo = timelapse.end
            inputs.addIntegerSpinnerCommandInput('interpolationFrames', 'Frames per Operation', 1, max_int, 1, timelapse.interpolationFrames)
            inputs.addBoolValueInput('doFit', 'Fit Design', True, '', timelapse.doFit)
            inputs.addBoolValueInput('rotate', 'Rotate Design', True, '', timelapse.rotate)
            inputs.addIntegerSpinnerCommandInput('framesPerRotation', 'Frames per Rotation', 1, max_int, 1, timelapse.framesPerRotation)
            inputs.addIntegerSpinnerCommandInput('finalFrames', 'Num Final Frames', 0, max_int, 1, timelapse.finalFrames)

        except Exception:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


class HistoryTimelapse:
    def __init__(self, design):
        dataFile = app.activeDocument.dataFile

        # Set initial values.
        self._foldername = dataFile.name
        self._outputPath = os.path.expanduser("~/Desktop/")
        self._saveObj = False
        self._timeline = design.timeline
        self._width = 2000
        self._height = 2000
        self._start = 1
        self._end = self._timeline.markerPosition
        self._interpolationFrames = 5
        self._rotate = True
        self._framesPerRotation = 500
        self._finalFrames = 0
        self._design = design
        self.doFit = False

    # Properties.
    @property
    def foldername(self):
        return self._foldername

    @foldername.setter
    def foldername(self, value):
        self._foldername = value

    @property
    def outputPath(self):
        return self._outputPath

    @outputPath.setter
    def outputPath(self, value):
        self._outputPath = value

    @property
    def saveObj(self):
        return self._saveObj

    @saveObj.setter
    def saveObj(self, value):
        self._saveObj = value

    @property
    def timeline(self):
        return self._timeline

    @property
    def width(self):
        return self._width
    @width.setter
    def width(self, value):
        self._width = value

    @property
    def height(self):
        return self._height

    @height.setter
    def height(self, value):
        self._height = value

    @property
    def start(self):
        return self._start

    @start.setter
    def start(self, value):
        self._start = value

    @property
    def end(self):
        return self._end

    @end.setter
    def end(self, value):
        self._end = value

    @property
    def interpolationFrames(self):
        return self._interpolationFrames

    @interpolationFrames.setter
    def interpolationFrames(self, value):
        self._interpolationFrames = value

    @property
    def rotate(self):
        return self._rotate

    @rotate.setter
    def rotate(self, value):
        self._rotate = value

    @property
    def framesPerRotation(self):
        return self._framesPerRotation

    @framesPerRotation.setter
    def framesPerRotation(self, value):
        self._framesPerRotation = value

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

    def isNumericExtent(self, extent):
        classname = type(extent).__name__
        # ui.messageBox(classname)
        if classname == 'DistanceExtentDefinition':
            return True
        if classname == 'SymmetricExtentDefinition':
            return True
        if classname == 'AngleExtentDefinition':
            return True
        return False

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
        framesPerRotation = self.framesPerRotation if self.rotate else 0
        finalFrames = self.finalFrames
        doFit = self.doFit

        if ui:
            ui.messageBox('Foldername: {}'.format(foldername))

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

        num = 0
        startingAngle = None

        ff = open(os.path.join(outputPath, 'log.txt'), 'w')
        os.makedirs(outputPath, exist_ok=True)

        for i in range(start, end):
            try:
                # Get feature at current timeline index.
                item = timeline.item(i)

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
                # ui.messageBox(classname)
            except Exception:
                continue

            # Set marker position.
            timeline.markerPosition = i + 1

            # Get parameters to interpolate.
            interpolatedParameters = []
            stepSizes = []
            stepOffsets = []
            alphaComponents = []
            if classname == 'ExtrudeFeature':
                if self.isNumericExtent(entity.extentOne):
                    param = entity.extentOne.distance
                    interpolatedParameters.append(param)
                    stepSizes.append(param.value / interpolationFrames)
                    stepOffsets.append(0)
                # At the very least we can fade it in if it's a new body/component.
                elif entity.operation == 3 or entity.operation == 4: # NewBodyFeatureOperation or NewComponentFeatureOperation.
                    bodies = entity.bodies
                    for body in bodies:
                        alphaComponents.append(body)
                if entity.hasTwoExtents:
                    # Handle side 2.
                    if self.isNumericExtent(entity.extentTwo):
                        param = entity.extentTwo.distance
                        interpolatedParameters.append(param)
                        stepSizes.append(param.value / interpolationFrames)
                        stepOffsets.append(0)
                    # At the very least we can fade it in if it's a new body/component.
                    elif entity.operation == 3 or entity.operation == 4: # NewBodyFeatureOperation or NewComponentFeatureOperation.
                        bodies = entity.bodies
                        for body in bodies:
                            alphaComponents.append(body)
            # if classname == 'OffsetFacesFeature': # TODO: unable to get extent parameter from this operation.
            if classname == 'Move':
                # TODO: implement this.
                try:
                    print(i, 'move:', str(entity.transform.asArray()), file=ff)
                except Exception as e:
                    print(i, 'move:', str(e), file=ff)
                continue
            elif classname == 'MirrorFeature':
                try:
                    print(i, 'mirror: bodies', entity.bodies.count, file=ff)
                    print(i, 'mirror: inputEntities', entity.inputEntities.count, file=ff)
                except Exception as e:
                    print(i, 'mirror:', str(e), file=ff)
                continue
            elif classname == 'RevolveFeature':
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
            elif classname == 'RectangularPatternFeature':
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
            _interpolationFrames = interpolationFrames if (i < end - 1) else (interpolationFrames + finalFrames)
            for j in range(_interpolationFrames):
                # Interpolate parameters.
                for k in range(len(interpolatedParameters)):
                    try:
                        value = stepSizes[k] * (j + 1) + stepOffsets[k]
                        if abs(value) > abs(originalValues[k]):
                            value = originalValues[k]
                        # ui.messageBox(str(value))
                        interpolatedParameters[k].value = value
                        # Force a recompute (needed for symmetric Revolves for some reason?).
                        if classname == 'RevolveFeature':
                            entity.extentDefinition.isSymmetric = entity.extentDefinition.isSymmetric
                    except RuntimeError:
                        # modifying extents may fail with
                        # e.g. "No body to cut" error
                        continue
                for k in range(len(alphaComponents)):
                    value = originalAlphas[k] * (j + 1) / interpolationFrames
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
                    'frame_{:05d}'.format(num))
                success = app.activeViewport.saveAsImageFile(outputFilename + '.png', width, height)
                if not success:
                    ui.messageBox('Failed saving viewport image.')

                # Save obj file if requested
                if saveObj:
                    success = self.saveObjFile(outputFilename + '.obj')
                    if not success:
                        ui.messageBox('Failed saving obj file.')

                num += 1

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
                    adsk.fusion.TriangleMeshQualityOptions.NormalQualityTriangleMesh
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
        ui.messageBox('WARNING: This script will make changes to your file (e.g. break links to referenced components).  You may want to run this on a copy of your design.  You can quit Fusion now to stop the script if you are unsure about continuing.')
        product = app.activeProduct
        design = fusion.Design.cast(product)
        if not design:
            ui.messageBox('Script is not supported in current workspace, please change to MODEL workspace and try again.')
            return
        # Init a timelapse object.
        if timelapse is None:
            timelapse = HistoryTimelapse(design)

        commandDefinitions: core.CommandDefinitions = ui.commandDefinitions
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

        # Prevent this module from being terminated when the script returns, because we are waiting for event handlers to fire.
        adsk.autoTerminate(False)

    except Exception:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
