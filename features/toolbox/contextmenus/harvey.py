# -*- coding: utf-8 -*-
'''
Created on 01.08.2022

@author: fstallmann
'''



import bkt

from .. import harvey


class ContextHarveyShapes(object):
    @staticmethod
    def get_buttons(shapes):
        if not harvey.harvey_balls.change_harvey_enabled(shapes):
            return []
        return [
            bkt.ribbon.MenuSeparator(title="Harvey Balls"),
            ### Harvey
            harvey.harvey_size_gallery(
                # insert_before_mso='Cut',
                # id='ctx_harvey_ball_size_gallery',
                # get_visible=bkt.Callback(harvey.harvey_balls.change_harvey_enabled, shapes=True)
            ),
            harvey.harvey_color_gallery(
                # insert_before_mso='Cut',
                # id='ctx_harvey_ball_color_gallery',
                # get_visible=bkt.Callback(harvey.harvey_balls.change_harvey_enabled, shapes=True)
            ),
        ]