================
Sheet Converter
================

  Author: Katie Nelson
  Email:  KatherineE.Nelson@Scotiabank.com

================
  Installation
================

  First, make sure you have Python 3.0+ installed on your computer.
  This is absolutely necessary for fairly obvious reasons.

  after that, run
    pip install -r requirements.txt

  in any shell or cmd environment while in the root directory. This
  should install all the required modules. After that, you're pretty
  much good to go.

================
  Configuration
================

  If you'd like to repurpose this program into one that works for any
  input structure, the way it's currently configured may make that
  difficult. Basically everything besides directory structure is hard
  coded. I've commented where things would need to be changed, but it
  may be easier to reverse-engineer the program yourself for your
  own purposes than try to add functionality to this one.

  If you would really like to, under src/config, there is a dictionary
  of rows in the template excel file--each country's row for the
  specific task that's being documented. The template file is a
  pre-made excel sheet with all the data areas blank. The only other
  things you would need to change is the for loop in the excel_control
  file under src/. That loops from the second to the 90th row in the
  input file and extracts data from each column.

  Basic configuration like Input file destination and output file name
  can all be configured at run time--see usage for deets.

================
  Usage
================

  in any shell/cmd environment run:

    python config.py

  to use all the default options. Configure the options as below:

  usage: convert.py [-h] [--output-dir OUTPUT_DIR] [--input-dir INPUT_DIR]
   [--backup-dir BACKUP_DIR] [--template TEMPLATE] [-b]
   [--output-name OUTPUT_NAME]

   Scrapes data from an input file and populates a template.

   optional arguments:

     -h, --help                 show this help message and exit

     --output-dir OUTPUT_DIR    Specifies output directory
                                Default: ./Output

     --input-dir INPUT_DIR      Specifies intput directory
                                Default: ./Input

     --backup-dir BACKUP_DIR    Specifies backup directory
                                Default: ./Backups

     --template TEMPLATE        Path to template file.
                                WARNING: CHANGE CONFIG FILE IF
                                TEMPLATE HAS CHANGED
                                Default: ./Template/Template.xlsx

     -b, --backup               Creates backups of the input files under
                                the backup-dir directory
                                Default: no backup

     --output-name OUTPUT_NAME  Output file name
                                Default: INTLStatusReport


================
  Source Code
================

  ./
    convert.py
    requirements.txt
  ./src
    __init__.py
    core.py
    excel_control.py
    utils.py
    ./Config
      __init__.py
      index.py
