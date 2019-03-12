

import bkt



# ==============
# = SpinnerBox =
# ==============

width_changer = bkt.Callback(lambda shapes, value: [setattr(shp, 'width', value) for shp in shapes], bkt.CallbackTypes.on_change, shapes=True)
width_getter = bkt.Callback(lambda shapes: shapes[0].width, bkt.CallbackTypes.get_text, shapes=True)

spinner_group = bkt.ribbon.Group(
    label="Spinner-Control",
    children=[
        
        
        bkt.ribbon.SpinnerBox(
            label="simple spinner for shape width",
            size_string="##########",
            # get_text and on_change are callbacks for the editbox
            get_text = width_getter,
            on_change = width_changer,
            # increment/decrement-callbacks are passed to the buttons
            increment = bkt.Callback(lambda shapes: [setattr(shp, 'width', shapes[0].width+20) for shp in shapes], shapes=True),
            decrement = bkt.Callback(lambda shapes: [setattr(shp, 'width', shapes[0].width-20) for shp in shapes], shapes=True)
        ),
        
        
        bkt.ribbon.RoundingSpinnerBox(
            label="rounding spinner for width",
            supertip="the rounding-spinner rounds to multiples of 10\nUse Ctrl-key to increment in steps by 10",
            size_string="##########",
            get_text = width_getter,
            on_change = width_changer,
            
            # big_step specifies the default increment/decrement-step
            big_step=20,
            
            # small_step specifies the increment/decrement-step with pressed ctrl-key
            small_step=10,
            
            # rounding applies to increment/decrement-clicks
            # the attribute rounding_factor specifies the factor, such that values are multiples of this factor,
            # rounding_factor=25,
            
            # the round_at-attribute specifies the power x, such that rounding-factor is 10^(-x), e.g. round_at=2 rounds to multiples of 0.01
            round_at=-1
            
        ),
        
        bkt.ribbon.RoundingSpinnerBox(
            label="rounding with cm-conversion",
            supertip="the rounding-spinner converts to cm and rounds to multiples of .1\nUse Ctrl-key to increment in steps by 10",
            size_string="##########",
            get_text = width_getter,
            on_change = width_changer,
            
            # default settings for big_step, small_step, rounding-factor: round_cm / round_pt / round_int
            round_cm = True,
            
            convert = 'pt_to_cm'
            
        )
        
        
    ]
)




# =================
# = Color Gallery =
# =================

# define some methods for the color-gallery

def change_rgb(shapes, color):
    # change Fill- and Line-color to RGB-value
    for shp in shapes:
        shp.Fill.ForeColor.RGB = color
        shp.Line.ForeColor.RGB = color 

def change_theme(shapes, color_index, brightness):
    # change Fill- and Line-color to Theme-color-value
    for shp in shapes:
        shp.Fill.ForeColor.ObjectThemeColor = color_index
        shp.Line.ForeColor.ObjectThemeColor = color_index
        shp.Fill.ForeColor.Brightness = brightness
        shp.Line.ForeColor.Brightness = brightness

def selected_color(shapes):
    # return Theme- and RGB-color-value of first shape
    return [shapes[0].Fill.ForeColor.ObjectThemeColor, shapes[0].Fill.ForeColor.Brightness, shapes[0].Fill.ForeColor.RGB]

# define the color-gallery
color_gallery_group = bkt.ribbon.Group(
    label="Color-Gallery",
    children=[
        bkt.ribbon.ColorGallery(
            label="set background- and border-color",
            size="large",
            on_rgb_color_change = bkt.Callback(change_rgb, shapes=True),
            # on_theme_color_change is optional - but necessary for correct indication of selected color
            on_theme_color_change = bkt.Callback(change_theme, shapes=True),
            get_selected_color = bkt.Callback(selected_color, shapes=True)
        )
    ]
)




# ================
# = Dynamic Menu =
# ================

def get_content():
    x = bkt.ribbon.Menu(
        xmlns="http://schemas.microsoft.com/office/2009/07/customui",
        id=None,
        children=[
            bkt.ribbon.Button(label="Button1"),
            bkt.ribbon.Button(label="Button2"),
            bkt.ribbon.Button(label="Button3")
        ]
    )
    return x.xml_string()
    

dynamic_menu = bkt.ribbon.Group(
    label="Color-Gallery",
    children=[
        bkt.ribbon.DynamicMenu(
            label="Dynamic Menu",
            get_content = bkt.Callback(get_content)
        )
    ]
)




# ==================
# = Define the tab =
# ==================

bkt.powerpoint.add_tab(
    bkt.ribbon.Tab(
        label="Demo BKT-Controls",
        children = [ spinner_group, color_gallery_group, dynamic_menu] 
    )
)


