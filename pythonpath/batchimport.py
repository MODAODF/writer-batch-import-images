

import easymacro as app

def image_batch_import():

    doc = app.active

    if (doc.type != "writer"):
        return

    cursor = doc.selection

    filters = (
        ('Images', '*.bmp;*.gif;*.jpg;*.jpeg;*.png;*.tif;*.tiff'),
    )
    images = app.paths.get('~/Pictures', filters, True)

    for image in images:
        cursor = cursor.insert_image(image)
        cursor = cursor.new_line(2)
        image = doc.draw_page[-1]
        image.original_size()

    return
