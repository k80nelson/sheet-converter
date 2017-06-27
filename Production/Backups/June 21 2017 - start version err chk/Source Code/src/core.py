#! python3
# -*- coding: utf-8 -*-

"""Core functionality is abstracted into excel_control.py
"""
import src.excel_control as exc

def start():
    control = exc.Excel_Control()
    control.populate()
    control.save()
