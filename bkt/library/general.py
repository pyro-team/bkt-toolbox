
import bkt.library.system


def apply_delta_on_ALT_key(setter_method, getter_method, shapes, value, **kwargs):
    '''
    If the ALT key is pressed, the setter-method is called shifting all values by the same delta for every shape,
    i.e. setter_method([shape], old_value + delta, **kwargs) is called for every shape

    The delta-value is obtained by comparing getter_method([shapes[0]]) and value.
    For every shape, old_value is obtained using getter_method([shape]).
    
    If the ALT key is not pressed, setter_method(shapes, value, **kwargs) is called
    '''
    
    alt_state = bkt.library.system.get_key_state(bkt.library.system.key_code.ALT)
    
    if not alt_state:
        for shape in shapes:
            setter_method(shape=shape, value=value, **kwargs)
        
    else:
        delta = value - getter_method(shape=shapes[0], **kwargs)
        for shape in shapes:
            old_value = getter_method(shape=shape, **kwargs)
            setter_method(shape=shape, value=old_value + delta, **kwargs)

    return None

    
