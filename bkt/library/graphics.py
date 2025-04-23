# -*- coding: utf-8 -*-
'''
Created on 04.06.2020

@author: fstallmann
'''



from bkt import dotnet
from bkt.helpers import memoize

Drawing = dotnet.import_drawing() #required for image resizing
Bitmap = Drawing.Bitmap


def make_thumbnail(path, width, height, save_path=None, background_color=None):
    '''
    Make thumbnail from image with given dimension.
    If no save_path is given, bitmap is returned.
    If no background_color is given, image and fill area is transparent.
    Based on https://stackoverflow.com/questions/1922040/how-to-resize-an-image-c-sharp
    '''
    with Drawing.Image.FromFile(path) as image:
        # init croped image
        bmp = Bitmap(width, height)
        graph = Drawing.Graphics.FromImage(bmp)

        if background_color is not None:
            graph.Clear(Drawing.ColorTranslator.FromOle(background_color))

        # compute scale
        scale = min(float(width) / image.Width, float(height) / image.Height)
        scaleWidth = int(image.Width * scale)
        scaleHeight = int(image.Height * scale)

        # set quality
        if background_color is None:
            graph.CompositingMode = Drawing.Drawing2D.CompositingMode.SourceCopy #determines whether pixels from a source image overwrite or are combined with background pixels, SourceCopy=preserve transparency
        graph.CompositingQuality  = Drawing.Drawing2D.CompositingQuality.HighQuality #determines the rendering quality level of layered images
        graph.InterpolationMode   = Drawing.Drawing2D.InterpolationMode.High #determines how intermediate values between two endpoints are calculated, better but slower: HighQualityBicubic
        graph.SmoothingMode       = Drawing.Drawing2D.SmoothingMode.AntiAlias #specifies whether lines, curves, and the edges of filled areas use smoothing (also called antialiasing)
        graph.PixelOffsetMode     = Drawing.Drawing2D.PixelOffsetMode.HighQuality #affects rendering quality when drawing the new image

        with Drawing.Imaging.ImageAttributes() as wrap_mode:
            wrap_mode.SetWrapMode(Drawing.Drawing2D.WrapMode.TileFlipXY)
            # redraw and save
            dest_rect = Drawing.Rectangle(int((width - scaleWidth)/2), int((height - scaleHeight)/2), scaleWidth, scaleHeight)
            graph.DrawImage(image, dest_rect, 0, 0, image.Width, image.Height, Drawing.GraphicsUnit.Pixel, wrap_mode)
        
        # dispose image
        image.Dispose()
        
        if save_path:
            bmp.Save(save_path)
            # close file
            bmp.Dispose()
        else:
            return bmp

def open_bitmap_nonblocking(path):
    with Bitmap(path) as img:
        bmp = Bitmap(img)
        img.Dispose()
    return bmp

@memoize
def empty_image(width, height):
    img = Bitmap(width, height)
    color = Drawing.ColorTranslator.FromHtml('#ffffff00')
    img.SetPixel(0, 0, color)
    return img