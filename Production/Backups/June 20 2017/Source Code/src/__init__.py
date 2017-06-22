#! python3
# -*- coding: utf-8 -*-
"""Core package initialization
"""
from . import core
from . import utils
from . import excel_control
from . import config

import sys

__project__     = 'sheet_converter'
__description__ = 'Scrapes data from an input file and populates a template.'
__version__     = '0.3'
__all__         = [ "core", "utils", "excel_control" ]

args  = None
index = None

if sys.version_info < (3, 0):
    sys.exit('Python 3.0+ is required.')

def init(_args):
    global args, index
    args = _args
    index = config.index.Index()
