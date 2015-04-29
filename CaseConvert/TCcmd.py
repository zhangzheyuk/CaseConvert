__author__ = 'zzy'
# !/usr/bin/env python
#-*- coding: utf-8 -*-
import argparse
import textwrap
import sys
import xlrd
from lxml import etree

# parser = argparse.ArgumentParser(
#     formatter_class=argparse.RawTextHelpFormatter,
#     description=textwrap.dedent(u'''\
#         -----------------------------------------------------------------------
#
#             -*-test_link-*-
#         ***********************************************************************
#         Function: Convert excel files to xml files which testlink know.
#         Usage:  use './test_link.py -h' to see help
#         Example: ./test_link.py  test_link.xls
#
#         '''))
#
# parser.add_argument('input', action="store", help="Input file name")
# parser.add_argument('-o', dest='output',
#                     help=u'output file name, default is output.xml', default='output.xml')
# parser.add_argument('--version', action='version',
#                     version=u'%test_link 1.0')
# options = parser.parse_args()


f_in = xlrd.open_workbook("113.xls")
sheet = f_in.sheet_by_index(0)

# create XML
testcases = etree.Element('testcases')

print ("row = %d"%(sheet.nrows))


def format_str(rawstr):
    rawstr = "<![CDATA[<p>" + rawstr + "</p>"
    return rawstr


for seq in range(1, sheet.nrows):
    # print(sheet.row_values(seq))
    try:
        name_ = sheet.row_values(seq)[0]
        summary_ = sheet.row_values(seq)[1]
        pre_ = sheet.row_values(seq)[2]
        importance_ = sheet.row_values(seq)[3]
        step_ = sheet.row_values(seq)[4]
        expect_ = sheet.row_values(seq)[5]
        exe_type_ = sheet.row_values(seq)[6]

        test_case = etree.SubElement(testcases, 'testcase', name=name_)
        summary = etree.SubElement(test_case, 'summary')
        summary.text = format_str(summary_)
        print(summary.text)

        preconditions = etree.SubElement(test_case, 'preconditions')
        print(pre_)
        # preconditions.text = u'"{0}"'.format(pre_)
        pre_ = pre_.replace("\n", "</br>")
        print(pre_)
        preconditions.text = format_str(pre_)
        print(preconditions.text)

        # <importance><![CDATA[2]]></importance>
        importance_level = etree.SubElement(test_case, 'importance')
        importance_level.text = str(int(importance_))
        print(importance_level.text)

        steps = etree.SubElement(test_case, 'steps')
        step = etree.SubElement(steps, 'step')
        step_number = etree.SubElement(step, 'step_number')
        step_number.text = str(1)

        # Transform the steps
        actions = etree.SubElement(step, 'actions')
        # actions.text = u'"{0}"'.format(step_)
        step_ = step_.replace("\n", "</br>")
        print(step_)
        actions.text = format_str(step_)
        print(actions.text)
        expectedresults = etree.SubElement(step, 'expectedresults')

        # expectedresults.text = u'"{0}"'.format(expect_)
        expect_ = expect_.replace("\n", "</br>")
        print(expect_)
        expectedresults.text = format_str(expect_)
        print(expectedresults.text)

        execution_type = etree.SubElement(step, 'execution_type')
        execution_type.text = str(int(exe_type_))
        print(execution_type.text)

    except Exception as e:
        print("line:", seq)
        print(str(e))
        for item in sys.exc_info():
            print item

s = etree.tostring(testcases, pretty_print=True)
f_out = open("my.xml", 'w')
f_out.write(s)