# -*- coding: utf-8 -*-
'''
Created on 26.01.2015

@author: rdebeerst
'''

import bkt
import bkt.ribbon
import bkt.callbacks
import bkt.factory as mod_factory
import bkt.annotation as mod_annotation


import clr
clr.AddReference("System.Drawing")
import System.Drawing.Bitmap as Bitmap


@bkt.control('button')
class SpecialButton(bkt.FeatureContainer):
    @bkt.callback()
    def on_action(self):
        print 'button action from container'

@bkt.control(bkt.ribbon.Button)
class SpecialButton2(bkt.FeatureContainer):
    @bkt.callback(context=True, shapes=True)
    def on_action(self, context, shapes):
        print 'button action from container with context'



@bkt.configure(label='Gruppe mit Button/EditBox')
#@bkt.uuid('e7a39709-94c1-494c-b9b4-9339b38b892f')
@bkt.group
class SomeGroup(bkt.FeatureContainer):

    @bkt.button
    #@bkt.uuid('3777a1e0-58cf-40b9-acba-10ced84340b3')
    @bkt.configure(label='Test', size='large', image_mso='HappyFace')
    def test_button(self):
        print 'hello world'

    paste = bkt.mso.control.PasteSpecialDialog
    
    @bkt.button
    @bkt.arg_context
    def button_with_context(self, context):
        print 'hello world'

    cut   = bkt.mso.control.Cut
    copy  = bkt.mso.control.CopySplitButton

    
    @test_button.get_enabled
    #@test_button.callback('get_enabled')
    def button_enabled(self):
        return True


    @bkt.large_button("I'm large")
    def a_large_button(self):
        print 'hello large button'
    
    bkt.use(SpecialButton)
    bkt.use(SpecialButton2)
    
    @bkt.edit_box
    def some_textbox(self, value):
        print 'text changed: ' + str(value)
    
    @some_textbox.get_text
    def some_textbox_get_text(self):
        return 'textbox text'
    
    # FIXME: parameter kann statt 'value' auch anders heissen wie hier 'new_text'
    # @bkt.edit_box
    # def some_textbox2(self, new_text):
    #     print 'text changed: ' + str(new_text)
    
    
    # toggle button example
    _pressed = True
    
    @bkt.configure(label='toggle me')
    @bkt.toggle_button
    def a_toggle_button(self, pressed):
        self._pressed = not self._pressed
    
    @a_toggle_button.get_pressed
    def a_toggle_button_get_pressed(self):
        return (self._pressed == True)
    
    


class SpinnerBox(bkt.ribbon.Box):
    default_callback = bkt.callbacks.Callbacks.on_change
    
    def __init__(self, **kwargs):
        #FIXME: kwargs ??
        super(SpinnerBox, self).__init__()
        self.txt_box = bkt.ribbon.EditBox(size_string='###')
        self.default_callback_control = self.txt_box

        self.inc_button = bkt.ribbon.Button(label="»")
        self.dec_button = bkt.ribbon.Button(label="«")

        self.children = [self.txt_box, self.dec_button, self.inc_button]

    def set_id(self, fallback_id=None, ribbon_short_id=None, id_tag=None):
        bkt.ribbon.Box.set_id(self,
                   fallback_id=fallback_id,
                   ribbon_short_id=ribbon_short_id,
                   id_tag=id_tag)
        box_id = self.args.id

        self.txt_box['id']    = box_id + '_text'
        self.inc_button['id'] = box_id + '_increment'
        self.dec_button['id'] = box_id + '_decrement'

    def add_callback(self, rc):
        if rc.callback.python_name == 'increment':
            self.inc_button.add_callback(rc, callback=bkt.callbacks.Callbacks.on_action)
        elif rc.callback.python_name == 'decrement':
            self.dec_button.add_callback(rc, callback=bkt.callbacks.Callbacks.on_action)
        else:
            self.txt_box.add_callback(rc)


 
    # def add_callbacks(self, callbacks):
    #     inc    = callbacks['increment'] or None
    #     dec    = callbacks['decrement'] or None
    #     get    = callbacks['get_text']  or None
    #     change = callbacks['on_change'] or None
    #
    #     if get:
    #         self.txt_box.get_text = get
    #     if change:
    #         self.txt_box.on_change = change
    #     if inc:
    #         self.inc_button.on_action = inc
    #     if dec:
    #         self.dec_button.on_action = dec



# @bkt.control(SpinnerBox)
# class TestSpinner(bkt.FeatureContainer):
#
#     @bkt.arg_context
#     def increment(self, context):
#         print('Yeah!', context)
#
#     def decrement(self):
#         pass
#
#     @bkt.callback('on_change')
#     def change_action(self, text):
#         pass
#
#     @bkt.callback('get_text')
#     def get_text(self):
#         pass


@bkt.configure(label='Test Spinner')
@bkt.control(bkt.ribbon.Group)
class TestSpinnerInGroup(bkt.FeatureContainer):
    @bkt.configure(label='a button')
    #FIXME: UI-Parameter erlauben, z.B. label
    @bkt.control(bkt.ribbon.Button, label='a button')
    def button(self):
        print 'a button'

    @bkt.configure(label='text', size_string='#######')
    @bkt.control(bkt.ribbon.SpinnerBox, label='text', size_string='#######')
    def change_method(self, value):
        print 'text changed: ' + value
    
    # @change_method.callback('get_text')
    @change_method.get_text
    def get_method(self):
        print 'return text'

    @change_method.increment
    def inc_method(self):
        print 'increment!'

    @change_method.decrement
    def dec_method(self):
        print 'decrement!'

    # @change_method.callback('increment')
    # def inc_method(self):
    #     print 'do increment'
    #
    # @change_method.callback('decrement')
    # def dec_method(self):
    #     print 'do decrement'





















@bkt.configure(label='Test Gallery')
@bkt.control(bkt.ribbon.Group)
class TestGalleryGroup(bkt.FeatureContainer):
    
    # @bkt.configure(label='a button' )
    # @bkt.control(bkt.ribbon.Button)
    # def the_action(self):
    #     print 'xx'
    #
    
    _selected = 0
    
    @bkt.configure(label='a gallery', columns='5')
    #FIXME: UI-Parameter erlauben, z.B. label
    @bkt.control(bkt.ribbon.Gallery)
    def the_action(self, selected_item, index):
        self._selected = index
        print 'action: selected=' + str(selected_item) + ' index=' + str(index)
        
    # mandatory
    @the_action.get_item_count
    def get_item_count(self):
        return 20
    
    @the_action.get_selected_item_index
    def get_selected_item_index(self):
        return self._selected
    
    @the_action.get_item_label
    def get_item_label(self, index):
        return 'item ' + str(index)
    
    # @the_action.get_item_image
    # def get_item_image(self, index):
    #     img = Bitmap(12, 12)
    #     return img

    @the_action.get_item_screentip
    def get_item_screentip(self, index):
        return 'screentip ' + str(index)
    
    
    @bkt.control(bkt.ribbon.ColorGallery)
    def the_color_action(self):
        print 'color action'

    @the_color_action.on_theme_color_change
    def the_color_action_theme(self, color_index, brightness):
        print 'theme-color action ' + str(color_index) + ' ' + str(brightness)
    
    




class foo(object):
    def bar(self):
        print 'foo bar'

aFoo = foo()

control_tree = bkt.ribbon.Group(
    label = 'a group',
    children = (
        bkt.ribbon.Button(
            label='button', 
            on_action=bkt.runtime.Callback(
                aFoo.bar,
                shapes=True,
                context=True
            ),
            get_visible=bkt.callbacks.Callback(xml_name='getVisible', dotnet_name='SpecialGetVisible'),
            get_image=bkt.runtime.Callback(
                aFoo.bar,
                #callback=bkt.callbacks.Callback(
                xml_name='getImageOtherXMLName',
                dotnet_name='SomeGetImageMethod'
            )
        ),
        bkt.ribbon.RibbonControl(
            'Button',
            label='another button',
            on_action=bkt.runtime.Callback(
                aFoo.bar
            )
        ),
    )
)

























@bkt.excel
@bkt.visio
@bkt.word
@bkt.powerpoint
@bkt.configure(label='Test')
@bkt.tab
class TabTest(bkt.FeatureContainer):
    grp1 = bkt.use(SomeGroup)
    grp2 = bkt.use(TestSpinnerInGroup)
    grp3 = bkt.use(TestGalleryGroup)
    # FIXME: 
    #grp3 = bkt.use(control_tree)
    






