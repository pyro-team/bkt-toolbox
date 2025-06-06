# -*- coding: utf-8 -*-
'''
Created on 05.01.2023

@author: fstallmann
'''

class MockCallback(object):
    @staticmethod
    def on_action():
        return 'onaction'


def do_something():
    return 'dosomething'