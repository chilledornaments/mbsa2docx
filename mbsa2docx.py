#!/usr/bin/python3
import argparse
import os
import glob
import shutil
import xml.etree.ElementTree as ET
import sys
import string
from docx import Document
from docx.shared import Inches,Pt,RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import re
import collections

#TODO: make lowest score an argument

#Set up Argparser
parser = argparse.ArgumentParser(description="A script to create a Flexential report from an MBSA report")
parser.add_argument("MBSAfile",help="The name of the .mbsa file")
args=parser.parse_args()

#Globals
MBSAfile = args.MBSAfile
docxname = args.MBSAfile.replace(".mbsa","_report.docx")
xmlroot = ET.parse(MBSAfile).getroot()
domain = xmlroot.get('Domain')
hostname  = xmlroot.get('Machine')
ipaddr = xmlroot.get('IP')
grade = xmlroot.get('Grade')
summarizedInfo = (domain, "\\", hostname, "", ipaddr, "", "Overall Grade: ", grade)
document = Document()
#create doc
def docx():
    document.add_heading('MBSA Report', 1)
    paragraph_format = document.styles['Normal'].paragraph_format
    paragraph_format.space_before = Pt(6)
    paragraph_format.space_after = Pt(6)
    document.save(docxname)
    readMBSA()
#do some useful stuff
def readMBSA():
        document.add_heading(summarizedInfo, 1)
        paragraph_format = document.styles['Normal'].paragraph_format
        paragraph_format.space_before = Pt(6)
        paragraph_format.space_after = Pt(6)
        for check in xmlroot.iter('Check'):
            seccheck = check.get('Name')
            seccheck = seccheck.upper()
            paragraph = document.add_paragraph('',style='Heading 1')
            runner=paragraph.add_run('{}'.format(seccheck))
            runner.bold=True
            document.save(docxname)
            for sectionGrade in check.get('Grade'):
                formattedSecGrade = ("Section Grade: " + str(sectionGrade))
                paragraph = document.add_paragraph('',style='Body Text')
                runner = paragraph.add_run('{}'.format(formattedSecGrade))
                runner.bold=True
                table = document.add_table(rows=1, cols=1)
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = seccheck
                table.style = 'TableGrid'
                document.save(docxname)
            for advice in check.iter('Advice'):
                easyadvice=(advice.text)
                if easyadvice is "None":
                    easyadvice = "No further information"
                    paragraph = document.add_paragraph('',style='Body Text')
                    runner = paragraph.add_run('{}'.format(easyadvice))
                    runner.bold=False
                    document.save(docxname)
                else:
                    paragraph = document.add_paragraph('',style='Body Text')
                    runner = paragraph.add_run('{}'.format(easyadvice))
                    runner.bold=False
                    document.save(docxname)
            for details in check.iter('Detail'):
                #this part isn't working
                deets = details.get('text')
                if not deets:
                    deets = ""
                    deets = details.get('text')
                    paragraph = document.add_paragraph('',style='Body Text')
                    runner = paragraph.add_run('{}'.format(deets))
                    runner.bold=False
                    document.save(docxname)
                elif deets is None:
                    deets = ""
                    deets = details.get('text')
                    paragraph = document.add_paragraph('',style='Body Text')
                    runner = paragraph.add_run('{}'.format(deets))
                    runner.bold=False
                    document.save(docxname)
                else:
                    paragraph = document.add_paragraph('',style='Body Text')
                    runner = paragraph.add_run('{}'.format(deets))
                    runner.bold=False
                    document.save(docxname)
                for rowG in details.iter('Row'):
                    grade = rowG.get('Grade')
                    if grade == 6:
                        grade = "TEST NOT APPLICABLE"
                    for col in rowG.iter('Col'):
                        colText=(col.text)
                        extColText = ("Row Grade: " + grade + "  " + "--" +  "  " + colText)
                        cell = table.add_row().cells
                        cell[0].text = extColText
                        document.save(docxname)
docx()
