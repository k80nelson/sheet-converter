# -*- coding: utf-8 -*-
"""Core package initialization
"""
from . import utils
from . import excel_control
from . import config

import sys

# all this can be ignored. Just metadata (that I never use... whoops)
__project__     = 'sheet_converter'
__description__ = 'Scrapes data from an input file and populates a template.'
__version__     = '3.0'
__all__         = [ "core", "utils", "excel_control" ]

args  = None
index = None


if sys.version_info < (3, 0):
    sys.exit('Python 3.0+ is required.') # ;)

# the global arguments
def init(_args):
    global args
    args = _args
