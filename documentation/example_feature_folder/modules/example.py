# -*- coding: utf-8 -*-

import bkt

class Example(object):
    @staticmethod
    def show_message():
        bkt.helpers.message("Example button was pressed!")



bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    label=u'Example',
    children = [
        bkt.ribbon.Group(
            label = "Example",
            children = [
                bkt.ribbon.Button(
                    id = 'example-button',
                    image_mso = 'HappyFace',
                    size='large',
                    label='Example',
                    on_action=bkt.Callback(Example.show_message)
                ),
            ]
        )
    ]
))


