#!/usr/bin/env python3
# -*- coding: utf-8 -*-
'''Convert.py -- Converts GEMS-style excel reports into a more readible format

This file is the main entry point of the program. All that is really here is the
args parser, and then it passes over control to ./src/excel_control.py

All of the comments and documentation here assumes a limited knowledge of
python--If you're already knowledgable on python you can ignore most of it.
That said, this is one of my first times programming in python, so if you see
an opportunity for optimizations to be put in place, by all means do so.

__author__ = 'Katie Nelson'
__email__  = ('KatherineE.Nelson@Scotiabank.com', 'ketnscripts@gmail.com')

Please dont message that gmail email unless you're real stuck on some bit of code
and I'm no longer working at Scotiabank. Thanks! :)

'''
import argparse
import os
import src as convert


def main():
    """Main entry point of program. Parses args and initializes data
    """
    default_output_dir    = os.path.join(os.getcwd(), 'Runtime', 'Output')
    default_input_dir     = os.path.join(os.getcwd(), 'Runtime', 'Input')
    default_backup_dir    = os.path.join(os.getcwd(), 'Runtime', 'Backups')
    default_template_path = os.path.join(os.getcwd(), 'Runtime',
                                         'Templates', 'Template.xlsx')

    # argument initialization. Specific usage is exmpanded on in the README
    p = argparse.ArgumentParser(description=convert.__description__)
    p.add_argument('--output-dir',
                   default=default_output_dir,
                   help='Specifies output directory')
    p.add_argument('--input-dir',
                   default=default_input_dir,
                   help='Specifies input directory')
    p.add_argument('--backup-dir',
                   default=default_backup_dir,
                   help='Specifies backup directory')
    p.add_argument('--template',
                   default=default_template_path,
                   help='Path to template file. WARNING: CHANGE CONFIG FILE IF \
                   TEMPLATE HAS CHANGED')
    p.add_argument('-b', '--backup',
                   action='store_true',
                   help='Option for creating backup files')
    p.add_argument('--output-name',
                   default='INTLStatusReport',
                   help='Output file name')
    p.add_argument('-q', '--quick',
                   action='store_true',
                   help='Doesn\'t open the generated excel file upon completion')

    # args parsing
    args = p.parse_args()


    # this call is to .src/__init__.py -- it stores the args and initializes some
    # other data
    convert.init(args)

    # call to ./src/excel_control.py
    convert.excel_control.start()

if __name__ == '__main__':
    main()
