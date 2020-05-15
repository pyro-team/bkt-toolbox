# -*- coding: utf-8 -*-


import bkt


@bkt.configure(label="Button Tests")
@bkt.group
class TestButton(bkt.FeatureContainer):
    
    @bkt.large_button("I'm large")
    def hello(self):
        bkt.helpers.log('large button')
    
    @bkt.configure(label="I'm small")
    @bkt.button
    def hello2(self):
        bkt.helpers.log('small button')
    
    @bkt.configure(label="toggle me")
    @bkt.toggle_button
    def hello3(self, pressed):
        bkt.helpers.log('now my toggle state is: ' + str(pressed))
    
    # @bkt.dialog_box_launcher()
    # def hello4(self):
    #     bkt.helpers.log('hello from dialog box laucher')


@bkt.configure(label="Test Edit Box")
@bkt.group
class TestEditBox(bkt.FeatureContainer):
    value = 5

    @bkt.configure(label='test')
    @bkt.edit_box
    def mytextbox(self, value):
        bkt.helpers.log('new value: ' + str(value))
        self.value = value

    @mytextbox.get_text
    def gettext(self):
        return self.value


    incdecvalue = 10
    
    # @bkt.configure(label='test incdec')
    # @bkt.spinner_box
    # def myindectextbox(self, value):
    #     self.incdecvalue = int(value)

    # @myindectextbox.get_text
    # def getincdectext(self):
    #     return self.incdecvalue

    # # overwrite default incrementor
    # @myindectextbox.incrementor
    # def incdectext_inc(self):
    #     self.incdecvalue = self.incdecvalue + 5


#
#
# ctxMenu = bkt.ribbon.ContextMenu(idMso='ContextMenuShape', children=[
# #    Button(idMso='FontDialog', visible='false'),
#     bkt.ribbon.Button(id='my-context-menu-button', label='a button on the shape context menu')
# ])
# bkt.powerpoint.add_context_menu(ctxMenu)
#
#
# @bkt.powerpoint
# class TestRibbon(bkt.ribbon.Tab):
#     label = u'Test Framework'
#
#     children = [ TestButton, TestEditBox ]
#
#     def __init__(self):
#         bkt.ribbon.Tab.__init__(self)
#
#         # Beispiel/Test fuer direkte Erzeugung von Ribbon-Objekten
#         self.children += [
#             bkt.ribbon.Group(
#                 label = 'TestRibbonControls',
#                 children = [
#                     # Attribute werden als Parameter uebergeben
#                     bkt.ribbon.Button(idMso='Copy'),
#                     bkt.ribbon.Button(idMso='Cut', label='myLabel'),
#                     # actions werden ueber BTKEventHandler definiert,
#                     # benoetigter context muss explizit angegeben werden (CallableContextInformation)
#                     bkt.ribbon.Button(id='test-button', label='test', on_action=bkt.EventHandler( self.helloWorld , context_info=None)),
#                     bkt.ribbon.EditBox(id='test-edit', label='value', on_change=None, get_text=None),
#                     # Laucher immer letztes Element der Gruppe, agiert wie ein Button
#                     bkt.ribbon.DialogBoxLauncher(idMso='PowerPointParagraphDialog')
#                 ]
#             )
#         ]
#
#
#     def helloWorld(self):
#         bkt.helpers.log('Hello World')
#
#
#
#
# def myCut():
#      bkt.helpers.log('do somthing before cut.\nReturn false to proceed with normal cut.')
#      return False
#
# bkt.powerpoint.add_repurposed_command(
#     bkt.ribbon.Command(idMso='Cut', id='Cut', on_action_repurposed=bkt.EventHandler(myCut, bkt.CallableContextInformation() ))
# )



@bkt.excel
@bkt.visio
@bkt.word
@bkt.powerpoint
@bkt.configure(label='Test Framework')
@bkt.tab
class TabTestFW(bkt.FeatureContainer):
    grp1 = bkt.use(TestButton)
    grp2 = bkt.use(TestEditBox)


