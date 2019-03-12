import bkt.console

@bkt.uuid('c3973689-0aec-4922-9846-80d1fdeed457')
@bkt.configure(label="BKT Dev Options")
@bkt.group
class DevGroup(bkt.FeatureContainer):
    #label = "BKT Dev Options"
    
    @bkt.button
    @bkt.uuid('ff170e04-e674-4ee7-9be2-60a6415900e8')
    @bkt.configure(label='Console', image_mso='WatchWindow', size='large')
    @bkt.no_transaction
    @bkt.arg_context
    def show_console(self, context):
        co = bkt.console
        co.console.Visible = True
        co.console.scroll_down()
        co.console.BringToFront()  # @UndefinedVariable
        co.console._globals['context'] = context
    
    @bkt.uuid('48dd8162-3948-4bc0-b50e-4f677f413692')
    @bkt.image_mso('Info')
    @bkt.arg_context
    @bkt.no_transaction
    @bkt.large_button('Show Config')
    def show_config(self, context):
        def _iter_lines():
            cfg = dict(context.config.items("BKT"))
            for k in sorted(cfg):
                yield k + ' = ' + str(getattr(context.config, k))
            yield ''
        
        bkt.console.show_message('\r\n'.join(_iter_lines()))

    @bkt.uuid('c2649109-cf04-4514-8a1e-75dc3a31776a')
    @bkt.image('xml')
    @bkt.arg_python_addin
    @bkt.arg_ribbon_id
    @bkt.no_transaction
    @bkt.large_button('Show Ribbon XML')
    def show_ribbon_xml(self, python_addin, ribbon_id):
        bkt.console.show_message(python_addin.get_custom_ui(ribbon_id))

    @bkt.uuid('043a3c86-6596-4e3d-9d92-727b870cfbf7')
    @bkt.image_mso('AccessRefreshAllLists')
    @bkt.arg_context
    @bkt.no_transaction
    @bkt.large_button('Reload BKT')
    def reload_bkt(self, context):
        try:
            addin = context.app.COMAddIns["BKT.AddIn"]
            addin.Connect = False
            addin.Connect = True
        except Exception, e:
            bkt.console.show_message(str(e))
   

# ===============================
# = ToggleButtons for TaskPanes =
# ===============================


# @bkt.configure(label="TaskPanes")
# @bkt.group
# class TaskPaneToggles(bkt.FeatureContainer):
#     # def __init__(self):
#     #     self.children = [
#     #         bkt.ribbon.ToggleButton(id='tptoggle-bkttaskpane', label='BKT Task Pane', tag='BKT Task Pane', get_pressed='GetPressed_TaskPaneToggler', on_action='OnAction_TaskPaneToggler', get_enabled='')
#     #         ]
#     #
#     #     # bkt.ribbon.Group.__init__(self, children = [
#     #     #     bkt.ribbon.ToggleButton(id='tptoggle-bkttaskpane', label='BKT Task Pane', tag='BKT Task Pane', get_pressed='GetPressed_TaskPaneToggler', on_action='OnAction_TaskPaneToggler', get_enabled='')
#     #     #     ])
#
#     # TODO:
#     #toggle_bkttaskpane = bkt.taskpane_toggler('BKT Task Pane', dict(label='xx', image_mso='xx'))
#
#     @bkt.configure(label='BKT Task Pane', on_action='OnAction_TaskPaneToggler', get_pressed='GetPressed_TaskPaneToggler')
#     @bkt.toggle_button
#     @bkt.callback_type(bkt.callbacks.CallbackTypes.get_enabled)
#     def toggle_bkttaskpane_enabled(self):
#         return True

if ((bkt.config.task_panes or False)):
    TaskPaneToggles = bkt.ribbon.Group(label="TaskPanes", children = [
        bkt.ribbon.ToggleButton(id='tptoggle-bkttaskpane', label='BKT Task Pane', image_mso='MenuToDoBar', tag='BKT Task Pane', get_pressed='GetPressed_TaskPaneToggler', on_action='OnAction_TaskPaneToggler')
       ])
else:
    TaskPaneToggles = None


@bkt.excel
@bkt.visio
@bkt.word
@bkt.powerpoint
@bkt.configure(id_mso='TabDeveloper')
@bkt.tab
class TabDeveloperOptions(bkt.FeatureContainer):
    id_mso = 'TabDeveloper'
    dev_groups = bkt.use(DevGroup)
    task_panes = TaskPaneToggles
