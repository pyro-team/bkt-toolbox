# -*- coding: utf-8 -*-
'''
DEPRECATED: Factory of creating controls from annotation-syntax

Created on 23.11.2014
@author: cschmitt
'''



import logging
import sys, inspect

import bkt.annotation as mod_annotation
import bkt.ribbon as mod_ribbon
import bkt.callbacks as mod_callbacks


# ====================================
# = Access to Ribbon Control classes =
# ====================================
# NOTE: The following 2 lines used to be in bkt.ribbon, but it was moved here as it is only required for legacy annotation syntax

clsmembers = inspect.getmembers(sys.modules["bkt.ribbon"], inspect.isclass)
RIBBON_CONTROL_CLASSES = {member[1]._python_name:member[1] for member in clsmembers if issubclass(member[1], mod_ribbon.RibbonControl) and hasattr(member[1], '_python_name')}


class ControlFactory(object):
    def __init__(self, annotation, id_tag=None, ribbon_info=None):
        if mod_annotation.is_feature_container_class(annotation):
            annotation = annotation._annotation
        self.annotation = annotation
        self.id_tag = id_tag
        self.ribbon_info = ribbon_info
        
        self.container_instance = None
        
    @property
    def ribbon_short_id(self):
        if self.ribbon_info is None:
            return None
        return self.ribbon_info.short_id
    
    def create_control(self):
        logging.debug("create_control for ControlFactory: %s" % self)
        logging.debug(self.annotation)
        annotation = self.annotation
        
        # container-Instanz erzeugen
        if annotation.is_root:
            self.container_instance = self.annotation.target()
            logging.debug("annotation-root dump: %s" % self.container_instance._dump())
            
            
        # callback annotationen auflösen, Instanz-Methoden dranhängen --> Callback
        if self.container_instance is not None:
            # FIXME: flacher Durchlauf durch _annotation_members ?
            for child in self.container_instance._annotation_members.values():
                ci = child.callback_info or mod_annotation.CallbackInformation()
                callback = mod_callbacks.Callback(self.container_instance, child.target.__name__, ci.callback_type, ci.invocation_context)
                # if not child.callback_info is None:
                #     # FIXME: make better Callback-Constructor
                #     rc.callback_type = child.callback_info.callback_type
                # if rc.callback_type is None:
                #     if mod_callbacks.CallbackTypes.has_callback_type(child.target.__name__):
                #         # method_name is a known callback
                #         rc.callback_type = mod_callbacks.CallbackTypes.get_callback_type(child.target.__name__)
                child.callback = callback
        
         
        ### Annotations-Baum auflösen
        
        # Root-Control erstellen
        if isinstance(annotation.ui_info.node_type, str):
            instance = RIBBON_CONTROL_CLASSES[annotation.ui_info.node_type](id_tag=self.id_tag, **annotation.ui_info.control_args)
        else:
            instance = annotation.ui_info.node_type(id_tag=self.id_tag, **annotation.ui_info.control_args)
        instance.set_id(annotation.target_name, self.ribbon_short_id, self.id_tag)
        self.control = instance
        
        
        ## Children einsammeln
        
        logging.debug("annotation children: %s" % self.annotation.children)
        mixed_children = list(self.annotation.children)
        if hasattr(self.annotation.target, '_mso_controls'):
             mixed_children += self.annotation.target._mso_controls
             logging.debug("mso_controls: %s" % self.annotation.target._mso_controls)
        if hasattr(self.annotation.target, '_usage'):
             mixed_children += self.annotation.target._usage
             logging.debug("usage: %s" % self.annotation.target._usage)
        mixed_children.sort(key=lambda elem : elem.target_order)
        
        ## Annotationen verarbeiten
        callbacks = {}
        children = {}
        sorted_chidren = []
        
        for child in mixed_children:
            
            if isinstance(child, mod_annotation.ContainerUsage):
                if child.container._annotation is None:
                    raise mod_annotation.DecorationError('object %s is not annotated, but referenced with @bkt.use(...)' % child.container)
                #FIXME: use =+ ???
                id_tag = child.id_tag
                child = child.container._annotation
                factory = ControlFactory(child, id_tag, ribbon_info=self.ribbon_info)
                #children.append( factory.create_control() )
                children[child.attribute] = factory.create_control()
                sorted_chidren.append(children[child.attribute])
            
            elif isinstance(child, mod_annotation.Annotation):
                if hasattr(child, 'callback'):
                    callback = child.callback
                    callbacks[child.target_name] = callback
                else:
                    callback = mod_callbacks.Callback(None, None, None)
                
                if child.ui_info is not None: #and child.ui_info.node_type in mod_ribbon.RIBBON_CONTROL_CLASSES:
                    # Annotation mit Control
                    factory = ControlFactory(child, self.id_tag, ribbon_info=self.ribbon_info)
                    control = factory.create_control()
                
                    callback.control = control #FIXME: wo wird das benoetigt?
                
                    # Control hat Callback
                    if callback.callback_type is None:
                        # control kann default callback haben
                        callback.callback_type = control.default_callback
                    
                    control.add_callback(callback)
                    #children.append(control)
                    children[child.target_name] = control
                    sorted_chidren.append(control)
                    
                elif child.callback_info is not None:
                    # Annotation mit Callback ohne eigenem Control
                    callback = child.callback
                    callback.control = self.control
                    self.control.add_callback(callback)
                    
                else:
                    raise ValueError
                    
            elif isinstance(child, mod_ribbon.RibbonControl):
                children[child.attribute] = child
                sorted_chidren.append(child)
                
            else:
                raise ValueError("Expected ContainerUsage, Annotation or RibbonControl. Got %s" % child)
                
        
        self.control.children += sorted_chidren
        
        
        if self.container_instance:
            if hasattr(self.container_instance, '_create_control'):
                self.control = self.container_instance._create_control(self.control, children, callbacks)
                if not (isinstance(self.control, mod_ribbon.RibbonControl) or mod_annotation.is_feature_container_class(self.control)):
                     raise TypeError('Unexpeced return-type in _create_control. Got %s, expected RibbonControl/FeatureContainer, in class %s.' % (type(self.control), type(self.container_instance)))
                
                # Rueckverweise korrigieren
                self.repair_control_references(self.control)
                
        return self.control

    
    def repair_control_references(self, control):
        for cb in control._callbacks.values():
            if cb.control == None:
                cb.control = control
        
        for child in control.children:
            self.repair_control_references(child)
