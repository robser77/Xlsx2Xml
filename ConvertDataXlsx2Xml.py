import logging
import argparse
import lxml.etree as etree
import sys
from openpyxl import load_workbook
from openpyxl import utils

def parse_workbook(workbook):
    root = etree.Element("workbook")
    for i, sheetname in enumerate(workbook.sheetnames):
        sheet = workbook[sheetname]
        sheet_tag = etree.SubElement(root, 'sheet' + str(i + 1))
        for j, row in enumerate(sheet.iter_rows(values_only=True)):
            row_tag = etree.SubElement(sheet_tag, 'row' + str(j + 1))
            for k, column in enumerate(row):
                column_tag = etree.SubElement(row_tag, 'column' + str(k + 1))
                column_tag.text = str(column) if column != None else ""
    tree = etree.ElementTree(root)
    return tree

def main():
    """used when executed from command line"""

    h_verbose = 'Increase output verbosity'
    h_input = 'Path and name to the input file that you want to transform.'
    h_output = 'Path and name to the output xml file that you want to create or "-" for stdout.'
    h_pretty = 'Pretty print output.'
    h_split = 'Create an output file for each worksheet.'

    parser = argparse.ArgumentParser()
    parser.add_argument("-v", "--verbose", help=h_verbose, action="store_true")
    parser.add_argument("-i", "--input", help=h_input, nargs=1, required=True)
    parser.add_argument("-o", "--output", help=h_output, default='-', nargs=1)
    parser.add_argument("-fo", "--pretty", help=h_pretty, action="store_true")
    parser.add_argument("-sp", "--split", help=h_split, action="store_true")

    args = parser.parse_args()
    if args.verbose:
        loglevel = 'INFO'
    else:
        loglevel = 'WARNING'

    if args.input:
        input_xlsx_name = args.input[0]
    if args.output:
        output_file_name = args.output[0]

    pretty_print = True if args.pretty else False

    numeric_level = getattr(logging, loglevel.upper())
    logging.basicConfig(level=numeric_level)

    try:
        workbook = load_workbook(filename=input_xlsx_name)

    except FileNotFoundError as error:
        print('Error: {}'.format(error))
        sys.exit(3)
    except utils.exceptions.InvalidFileException as error:
        print('Error: {}'.format(error))
        sys.exit(2)
    except:
        print('Error {}'.format(sys.exc_info()))
        sys.exit(1)

    tree = parse_workbook(workbook)

    if not args.output or output_file_name == '-' :
        sys.stdout.write(etree.tostring(tree, pretty_print=pretty_print, encoding='unicode'))
    elif len(output_file_name) > 0:
        try:
            tree.write(output_file_name, pretty_print=pretty_print, encoding='UTF-8')
        except PermissionError as error:
            print('Failed to write to file: {}'.format(error))
            quit()

if __name__ == "__main__":
    main()
