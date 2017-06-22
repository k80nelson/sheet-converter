#! python3
# -*- coding: utf-8 -*-

import argparse
import os
import src as convert
import src.config as config

# main entry point of the program. Lots to configure --use python converter.py -h for deets
def main():
    """Main entry point of program. Parses args and initializes data
    """
    default_output_dir    = os.path.join(os.getcwd(), 'Output')
    default_input_dir     = os.path.join(os.getcwd(), 'Input')
    default_backup_dir    = os.path.join(os.getcwd(), 'Backups')
    default_template_path = os.path.join(os.getcwd(), 'Templates', 'Template.xlsx')

    parser = argparse.ArgumentParser(description=convert.__description__)
    parser.add_argument('--output-dir', default=default_output_dir,
                         help='Specifies output directory')
    parser.add_argument('--input-dir', default=default_input_dir,
                         help='Specifies intput directory')
    parser.add_argument('--backup-dir', default=default_backup_dir,
                         help='Specifies backup directory')
    parser.add_argument('--template', default=default_template_path,
                         help='Path to template file. WARNING: CHANGE CONFIG FILE IF TEMPLATE HAS CHANGED')
    parser.add_argument('-b', '--backup', action='store_true',
                         help='Option for creating backup files')
    parser.add_argument('--output-name', default='INTLStatusReport',
                         help='Output file name')

    args = parser.parse_args()
    convert.init(args)
    convert.core.start()

if __name__ == '__main__':
    main()
