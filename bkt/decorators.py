# -*- coding: utf-8 -*-
'''
DEPRECATED: Decorators for old annotation-syntax

Created on 11.11.2014
@author: cschmitt
'''



from bkt.annotation import (require_ui,
                         ensure_callback_nfo,
                         ensure_invocation_context,
                         AnnotationCommand, #MagicLazyAnnotator,
                         AnnotatedMethod,
                         #LazyAnnotation,
                         UIControlAnnotationCommand,
                         PRIO_CONFIGURATION,
                         PRIO_LOWEST,
                         PRIO_NODE_DEFINITION,
                         PRIO_CALLBACK_DEFINITION,
                         DecorationError,
                         ContainerUsage
                         )
 
from bkt.callbacks import CallbackTypes




# ============================
# = callback-type decorators =
# ============================

def create_callback_command(callback_type, lazy=False):
    @ensure_callback_nfo
    def annotate_callback_lazy(annotation):
        if annotation.callback_info.callback_type is None:
            annotation.callback_info.callback_type = callback_type
        
    @ensure_callback_nfo
    def annotate_callback(annotation):
        if annotation.callback_info.callback_type is not None:
            raise DecorationError('callback for target %s already set to %s' % (annotation.target, annotation.callback_information.callback_type))
        annotation.callback_info.callback_type = callback_type
    
    return AnnotationCommand(annotate_callback_lazy, prio=PRIO_CALLBACK_DEFINITION+1) if lazy else AnnotationCommand(annotate_callback, prio=PRIO_CALLBACK_DEFINITION)

regular_callback_decorators = {}
for callback in CallbackTypes.callback_list():
    regular_callback_decorators[callback.python_name] = create_callback_command(callback, lazy=False)

increment = regular_callback_decorators['increment']
decrement = regular_callback_decorators['decrement']
on_change = regular_callback_decorators['on_change']
get_text  = regular_callback_decorators['get_text']



# FIXME
#LazyAnnotation._sub_annotators.update(child_decorators)
#LazyAnnotation._sub_annotators.update(regular_callback_decorators)

# @ensure_callback_nfo
# def callback_annotator(annotation, callback_type):
#     annotation.callback_info.callback = callback_type
#
# callback_type = AnnotationCommand(callback_annotator, PRIO_CALLBACK_DEFINITION)

def callback_type(callback_type):
    return create_callback_command(callback_type, lazy=False)

# def callback_type_str(callback_type_str):
#     callback_type = getattr(CallbackTypes, callback_type_str)
#     return callback_type(callback_type)
#
#
# AnnotatedMethod._default_sub_annotation_command = callback_type_str

@ensure_callback_nfo
def _set_callback_type_by_string(annotation, callback_type_str):
    callback_type = getattr(CallbackTypes, callback_type_str)
    # if annotation.callback_info.callback_type is not None:
    #     raise DecorationError('callback for target %s already set to %s' % (annotation.target, annotation.callback_information.callback_type))
    annotation.callback_info.callback_type = callback_type
    
AnnotatedMethod._default_sub_annotation_command = AnnotationCommand(_set_callback_type_by_string, PRIO_CALLBACK_DEFINITION)


# =================
# = UI Decorators =
# =================

#@AnnotationCommand(PRIO_CONFIGURATION)
@require_ui
def _configure(annotation, *args, **kwargs):
    # args erlauben, aus kwargs['posargs'] lesen, welche Parameter damit gesetzt werden
    ui_nfo = annotation.ui_info
    if len(args) > 0:
        if 'pos_args' in kwargs and len(kwargs['pos_args']) >= len(args):
            for i in range(0,len(args)):
                kwargs[kwargs['pos_args'][i]] = args[i]
            del kwargs['pos_args']
        else:
            print('args=%s, kwargs=%s', (args, kwargs))
            raise ValueError('Too many non-keyword-arguments for configure-annotation, args=%s, kwargs=%s' % (args, kwargs))
    
    if 'pos_args' in kwargs:
        del kwargs['pos_args']
    
    for k, v in kwargs.items():
        #if not hasattr(ui_nfo, k):
        #    raise ValueError('unkown attribute %s' % k)
        ui_nfo.control_args[k] = v

@require_ui
def _uuid(annotation, uuid):
    annotation.ui_info.control_args['uuid'] = uuid

@require_ui
def _image_mso(annotation, image_mso):
    annotation.ui_info.control_args['image_mso'] = image_mso

@require_ui
def _image(annotation, image):
    annotation.ui_info.control_args['image'] = image

        
def node_type_decorator(node_type):
    #return AnnotationCommand(NodeTypeSetter(node_type), PRIO_NODE_DEFINITION)
    return NodeTypeSetter(node_type)

configure = AnnotationCommand(_configure, PRIO_CONFIGURATION)
uuid      = AnnotationCommand(_uuid, PRIO_CONFIGURATION)
image_mso = AnnotationCommand(_image_mso, PRIO_CONFIGURATION)
image     = AnnotationCommand(_image, PRIO_CONFIGURATION)


def uicontrol(node_type, **kwargs):
    #return configure(**kwargs)(node_type_decorator(node_type))
    return UIControlAnnotationCommand(node_type) * configure(**kwargs)

#control = UIControlAnnotationCommand(node_type)(configure)

#ribbon        = UIControlAnnotationCommand('ribbon')
#tabs          = UIControlAnnotationCommand('tabs')
config_with_label = configure(pos_args=['label'])
tab           = config_with_label * UIControlAnnotationCommand('tab')
group         = config_with_label * UIControlAnnotationCommand('group')
menu          = config_with_label * UIControlAnnotationCommand('menu')
box           = config_with_label * UIControlAnnotationCommand('box')

button        = config_with_label * create_callback_command(CallbackTypes.on_action) * UIControlAnnotationCommand('button') 
toggle_button = config_with_label * create_callback_command(CallbackTypes.on_toggle_action, lazy=True) * UIControlAnnotationCommand('toggle_button')
edit_box      = config_with_label * create_callback_command(CallbackTypes.on_change) * UIControlAnnotationCommand('edit_box')
spinner_box   = config_with_label * create_callback_command(CallbackTypes.on_change) * UIControlAnnotationCommand('spinner_box')
combo_box     = config_with_label * create_callback_command(CallbackTypes.on_change) * UIControlAnnotationCommand('combo_box')

gallery       = config_with_label * create_callback_command(CallbackTypes.on_action_indexed) * UIControlAnnotationCommand('gallery')
item          = config_with_label * UIControlAnnotationCommand('item')


large_button = configure(size='large', pos_args=['label']) * button

#button       = AnnotationCommand(create_callback_command(CallbackTypes.on_action, lazy=True), prio=100)( UIControlAnnotationCommand('button') )

# def large_button(label):
#     def deco(target):
#         t = button(target)
#         t = configure(size='large', label=label)(t)
#         return t
#     return deco


child_decorators = dict(group=group,
                        gallery=gallery,
                        button=button,
                        item=item,
                        edit_box=edit_box,
                        tab=tab)



###############################################



use = ContainerUsage



###############################################




def invocation_annotation(anno_func):
    
    @ensure_invocation_context
    def annotator(annotation):
        anno_func(annotation.callback_info.invocation_context)
    
    annotator.__name__ = anno_func.__name__ + '|' + annotator.__name__
    return AnnotationCommand(annotator)

@invocation_annotation
def require_text(ictx):
    ictx.require_text = True

@invocation_annotation
def arg_page_shapes(ictx):
    ictx.page_shapes = True

@invocation_annotation
def arg_python_addin(ictx):
    ictx.python_addin = True

@invocation_annotation
def arg_presentation(ictx):
    ictx.presentation = True

@invocation_annotation
def arg_context(ictx):
    ictx.context = True

@invocation_annotation
def arg_ribbon_id(ictx):
    ictx.ribbon_id = True

@invocation_annotation
def arg_shapes(ictx):
    ictx.shapes = True

def arg_shapes_limited(shapes_min=None, shapes_max=None):
    @invocation_annotation
    def annotate(ictx):
        ictx.shapes = True
        ictx.shapes_max = shapes_max
        ictx.shapes_min = shapes_min
    return annotate

@invocation_annotation
def arg_shape(ictx):
    ictx.shapes = True
    ictx.shape = True
    ictx.shapes_max = 1
    ictx.shapes_min = 1

@invocation_annotation
def arg_slides(ictx):
    ictx.slides = True

def arg_slides_limited(shapes_min=None, shapes_max=None):
    @invocation_annotation
    def annotate(ictx):
        ictx.slides = True
        ictx.slides_max = shapes_max
        ictx.slides_min = shapes_min
    return annotate

@invocation_annotation
def arg_slide(ictx):
    ictx.slides = True
    ictx.slide = True
    ictx.slides_max = 1
    ictx.slides_min = 1

@invocation_annotation
def no_transaction(ictx):
    ictx.transaction = False

@invocation_annotation
def transaction(ictx):
    ictx.transaction = True
    