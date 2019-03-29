# -*- coding: utf-8 -*-

import bkt

class DevGroup(object):
    
    @staticmethod
    def show_console(context):
        import bkt.console as co
        co.console.Visible = True
        co.console.scroll_down()
        co.console.BringToFront()  # @UndefinedVariable
        co.console._globals['context'] = context
    
    @staticmethod
    def show_config(context):
        import bkt.console
        def _iter_lines():
            cfg = dict(context.config.items("BKT"))
            for k in sorted(cfg):
                yield k + ' = ' + str(getattr(context.config, k))
            yield ''
        
        bkt.console.show_message('\r\n'.join(_iter_lines()))
    
    @staticmethod
    def show_settings(context):
        import bkt.console
        def _iter_lines():
            for k in sorted(context.settings):
                yield k + ' = ' + repr(context.settings.get(k, "ERROR"))
            yield ''
        
        bkt.console.show_message('\r\n'.join(_iter_lines()))

    @staticmethod
    def show_ribbon_xml(python_addin, ribbon_id):
        import bkt.console
        bkt.console.show_message(python_addin.get_custom_ui(ribbon_id))

    @staticmethod
    def reload_bkt(context):
        import bkt.console
        try:
            addin = context.app.COMAddIns["BKT.AddIn"]
            addin.Connect = False
            addin.Connect = True
        except Exception, e:
            bkt.console.show_message(str(e))
   

dev_group = bkt.ribbon.Group(
    id="bkt_dev_group",
    label="BKT Dev Options",
    image_mso="AccessRefreshAllLists",
    children=[
        bkt.ribbon.Button(
            label="Console",
            size="large",
            image_mso="WatchWindow",
            on_action=bkt.Callback(DevGroup.show_console, context=True, transaction=False),
        ),
        bkt.ribbon.Button(
            label="Show Config",
            size="large",
            image_mso="Info",
            on_action=bkt.Callback(DevGroup.show_config, context=True, transaction=False),
        ),
        bkt.ribbon.Button(
            label="Show Settings",
            size="large",
            image_mso="Info",
            on_action=bkt.Callback(DevGroup.show_settings, context=True, transaction=False),
        ),
        bkt.ribbon.Button(
            label="Show Ribbon XML",
            size="large",
            image="xml",
            on_action=bkt.Callback(DevGroup.show_ribbon_xml, python_addin=True, ribbon_id=True, transaction=False),
        ),
        bkt.ribbon.Button(
            label="Reload BKT",
            size="large",
            image_mso="AccessRefreshAllLists",
            on_action=bkt.Callback(DevGroup.reload_bkt, context=True, transaction=False),
        ),
    ]
)

dev_tab = bkt.ribbon.Tab(
    idMso="TabDeveloper",
    children=[dev_group]
)


# ===============================
# = ToggleButtons for TaskPanes =
# ===============================

if ((bkt.config.task_panes or False)):
    dev_tab.children.append(
        bkt.ribbon.Group(label="TaskPanes", children = [
            bkt.ribbon.ToggleButton(id='tptoggle-bkttaskpane', size="large", label='BKT Task Pane', image_mso='MenuToDoBar', tag='BKT Task Pane', get_pressed='GetPressed_TaskPaneToggler', on_action='OnAction_TaskPaneToggler')
        ])
    )



bkt.powerpoint.add_tab(dev_tab)
bkt.word.add_tab(dev_tab)
bkt.excel.add_tab(dev_tab)
bkt.visio.add_tab(dev_tab)

# @bkt.excel
# @bkt.visio
# @bkt.word
# @bkt.powerpoint
# @bkt.configure(id_mso='TabDeveloper')
# @bkt.tab
# class TabDeveloperOptions(bkt.FeatureContainer):
#     id_mso = 'TabDeveloper'
#     dev_groups = dev_group
#     task_panes = TaskPaneToggles