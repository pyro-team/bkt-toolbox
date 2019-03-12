# -*- coding: utf-8 -*-
'''
Created on 25.02.2019

@author: fstallmann
'''

import os.path

# install bkt-framework with additional config
features_path = os.path.normpath(os.path.join(os.path.dirname(__file__),'..', 'features'))

default_config=dict(
        feature_folders=[
            os.path.join(features_path, 'toolbox'),
            os.path.join(features_path, 'ppt_thumbnails'),
            os.path.join(features_path, 'ppt_notes'),
            os.path.join(features_path, 'ppt_shapetables'),
            os.path.join(features_path, 'ppt_circlify'),
        ],
        shape_library_folders = '',
        shape_libraries =  '',
        chart_library_folders = '',
        chart_libraries =  ''
    )