# -*- coding: utf-8 -*-
'''
Created on 2020-09-19
@author: Florian Stallmann
'''



import logging

import bkt.ui
notify_property = bkt.ui.notify_property

import bkt.library.powerpoint as pplib


class SlidesSync(object):
    def __init__(self, template, slides, format=True, position=True, text=True, shapes_add=True, shapes_remove=True, skip_placeholders=True):
        self._template = template
        self._slides = slides

        self._format = format
        self._position = position
        self._text = text
        self._shapes_add = shapes_add
        self._shapes_remove = shapes_remove
        self._skip_placeholders = skip_placeholders

    def sync_slides(self, worker):
        try:
            #FIXME: shape name is not unique
            template_shapes = {shape.name: shape for shape in self._template.shapes}
        except:
            logging.exception("error getting template shapes")
            return
        #go through all shapes of all other slides
        len_slides = len(self._slides)
        for i,slide in enumerate(self._slides):
            if slide == self._template:
                continue
            try:
                if worker.CancellationPending:
                    break
                worker.ReportProgress(round(i/len_slides*100))
                self.sync_slide(worker, template_shapes, slide)
            except:
                logging.exception("error with slide %s", slide.slideindex)
        
    def sync_slide(self, worker, template_shapes, slide):
        all_keys = set(template_shapes.keys())

        #go through all shapes
        for shp in list(iter(slide.shapes)):
            if worker.CancellationPending:
                return
            try:
                all_keys.remove(shp.name)
            except KeyError:
                #shape is not on template slide
                if self._shapes_remove and shp.visible: #FIXME: check for placeholder?
                    logging.info("slide sync removed %s", shp.name)
                    shp.delete()
            else:
                self.sync_shape(template_shapes[shp.name], shp)
        
        if self._shapes_add:
            for shp_id in all_keys:
                #add shape to slide
                old = template_shapes[shp_id]
                old.copy()
                # slide.shapes.paste()
                new = pplib.save_paste(slide.shapes)
                new.name = old.name #ensure that copied shape has the same name
                logging.info("slide sync added %s", old.name)
        
    def sync_shape(self, template_shape, shape):
        #skip invisible shapes and placeholders
        if not template_shape.visible or (self._skip_placeholders and template_shape.type == pplib.MsoShapeType['msoPlaceholder']):
            logging.info("slide sync skipped %s", template_shape.name)
            return
        
        logging.info("slide sync for template shape %s and shape %s", template_shape.name, shape.name)
        #apply stored formats
        if self._format:
            try:
                template_shape.pickup()
                shape.apply()
            except:
                logging.exception("slide sync error format")
        if self._position:
            try:
                shape.LockAspectRatio = 0
                shape.top, shape.left = template_shape.top, template_shape.left
                shape.width, shape.height = template_shape.width, template_shape.height
                shape.rotation = template_shape.rotation
                shape.LockAspectRatio = template_shape.LockAspectRatio
            except:
                logging.exception("slide sync error position")
        if self._text:
            try:
                if template_shape.hastextframe and template_shape.textframe2.hastext:
                    pplib.transfer_textrange(template_shape.textframe2.textrange, shape.textframe2.textrange)
            except:
                logging.exception("slide sync error text")


class ViewModel(bkt.ui.ViewModelSingleton):

    def __init__(self):
        super(ViewModel, self).__init__()

        self._sync_format = True
        self._sync_position = True
        self._sync_text = True
        self._sync_add = True
        self._sync_remove = True
        self._skip_placeholders = True

    @notify_property
    def sync_format(self):
        return self._sync_format
    @sync_format.setter
    def sync_format(self, value):
        self._sync_format = value

    @notify_property
    def sync_position(self):
        return self._sync_position
    @sync_position.setter
    def sync_position(self, value):
        self._sync_position = value

    @notify_property
    def sync_text(self):
        return self._sync_text
    @sync_text.setter
    def sync_text(self, value):
        self._sync_text = value

    @notify_property
    def sync_add(self):
        return self._sync_add
    @sync_add.setter
    def sync_add(self, value):
        self._sync_add = value

    @notify_property
    def sync_remove(self):
        return self._sync_remove
    @sync_remove.setter
    def sync_remove(self, value):
        self._sync_remove = value

    @notify_property
    def skip_placeholders(self):
        return self._skip_placeholders
    @skip_placeholders.setter
    def skip_placeholders(self, value):
        self._skip_placeholders = value



class SlideSyncWindow(bkt.ui.WpfWindowAbstract):
    _xamlname = 'slides_sync'
    _vm_class = ViewModel

    def __init__(self, context):
        self._template = min(context.slides, key=lambda s: s.slideindex)
        self._slides = context.slides
        super(SlideSyncWindow, self).__init__(context)
    
    def show_dialog(self, modal=True):
        if len(self._slides) < 2:
            return bkt.message.error("Es müssen mind. 2 Folien ausgewählt sein!")
        return super(SlideSyncWindow, self).show_dialog(modal)
    
    def sync(self, sender, event):
        vm = self._vm
        slidesync = SlidesSync(self._template, self._slides, vm.sync_format, vm.sync_position, vm.sync_text, vm.sync_add, vm.sync_remove, vm.skip_placeholders)
        bkt.ui.execute_with_progress_bar(slidesync.sync_slides, self._context)
        self.Close()
