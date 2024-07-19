import uno, tempfile, sys
from PIL import Image
from io import BytesIO
import screen_io as ui


def CompessImages80():

    # Main Instances
    IMAGE = 'com.sun.star.drawing.GraphicObjectShape'
    SHAPE = 'com.sun.star.drawing.CustomShape'
    oDoc = XSCRIPTCONTEXT.getDocument() 
    oSheet = oDoc.CurrentController.ActiveSheet
    shapes_number = oSheet.DrawPage.Count

    # Delete lists
    shapes_list = []
    original_images = []

    for i in range(0, shapes_number):

        shape = oSheet.DrawPage[i]
        if shape.ShapeType == IMAGE:

            original_images.append(shape)

            # Get original image
            image_data = shape.Bitmap.DIB
            size = shape.getSize()
            position = shape.getPosition()
            im = Image.open(BytesIO(image_data.value))

            # Comress image to tempfile and paste it to sheet
            with tempfile.NamedTemporaryFile(suffix='.jpg') as jpg:
                im.save(jpg.name, format="JPEG", optimize=True, quality=80)

                img = oDoc.createInstance(IMAGE)
                img.GraphicURL = 'file://' + jpg.name
                img.setSize(size)
                img.setPosition(position)
                oSheet.DrawPage.add(img)

        # Append shapes to list
        elif shape.ShapeType == SHAPE:
            shapes_list.append(shape)

    # Delete shapes
    if shapes_list:
        for shape in shapes_list:
            shape.dispose()

    # Delete original images
    if original_images:
        for shape in original_images:
            shape.dispose()
    
    return


def CompessImages60():

    IMAGE = 'com.sun.star.drawing.GraphicObjectShape'
    SHAPE = 'com.sun.star.drawing.CustomShape'
    oDoc = XSCRIPTCONTEXT.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    shapes_number = oSheet.DrawPage.Count
    shapes_list = []
    original_images = []

    for i in range(0, shapes_number):

        shape = oSheet.DrawPage[i]
        if shape.ShapeType == IMAGE:

            original_images.append(shape)

            image_data = shape.Bitmap.DIB
            size = shape.getSize()
            position = shape.getPosition()

            im = Image.open(BytesIO(image_data.value))

            with tempfile.NamedTemporaryFile(suffix='.jpg') as jpg:
                im.save(jpg.name, format="JPEG", optimize=True, quality=60)

                img = oDoc.createInstance('com.sun.star.drawing.GraphicObjectShape')
                img.GraphicURL = 'file://' + jpg.name
                img.setSize(size)
                img.setPosition(position)
                oSheet.DrawPage.add(img)

        elif shape.ShapeType == SHAPE:
            shapes_list.append(shape)
    if shapes_list:
        for shape in shapes_list:
            shape.dispose()

    if original_images:
        for shape in original_images:
            shape.dispose()

    return


def GetInterpreterPath():
    # oDoc = XSCRIPTCONTEXT.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet

    ui.MsgBox(f'{sys.executable}', title="Interpreter path")

    return