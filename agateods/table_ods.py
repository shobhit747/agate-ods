"""
This module contains the ODS extension to :class:`Table <agate.table.Table>`.
"""

import agate

import zipfile
import lxml.etree as etree
import xml.etree.ElementTree as ET
import io
import logging

ODS_CONTENT_FILE = 'content.xml'

class UnsupportedFileExtensionError(Exception):
    """Raised when the file has an unsupported extension."""
    pass

def get_namespaces_lxml(xml_file):
    """
    Extracts all the namespaces present in XML file
    """
    xml_file.seek(0)
    tree = etree.parse(xml_file)
    return tree.getroot().nsmap

def read_ods_content_file(ods_file_path : str):
    """
    Opens the .ods file and returns 
    """
    extension = ods_file_path.split('.')[-1]
    if not zipfile.is_zipfile(ods_file_path) and extension != 'ods':
        raise UnsupportedFileExtensionError(f'Failed to read {ods_file_path} as an ODS file.')

    ods_file_zip = zipfile.ZipFile(ods_file_path,mode='r')
    if ODS_CONTENT_FILE in ods_file_zip.namelist():
        content_file = ods_file_zip.open(ODS_CONTENT_FILE)
        return content_file

def from_ods(cls,file_path, sheet=None, skip_lines=0, header=True, row_limit=None, **kwargs):
    if not isinstance(skip_lines, int):
        raise ValueError('skip_lines argument must be an int')

    content_file = read_ods_content_file(file_path)
    tree = ET.parse(content_file)
    root = tree.getroot()
    ns = get_namespaces_lxml(content_file) #xml file namespaces 

    table_tag = '{%s}table' % ns['table']
    table_row_tag = '{%s}table-row' % ns['table']
    cell_tag = '{%s}table-cell' % ns['table']
    p_tag = '{%s}p' % ns['text']

    sheetnames = list()
    for table in root.iter(table_tag):
        sheetnames.append(table.attrib['{%s}name' % ns['table']])
    sheets = [sheets for sheets in root.iter(table_tag)]

    if sheet is not None:
        if isinstance(sheet, int):
            sheet_to_operate_on = sheets[sheet]
        elif isinstance(sheet,str):
            if sheet in sheetnames:
                sheet_to_operate_on = sheets[sheetnames.index(sheet)]
            else:
                raise ValueError(f"No sheet with name '{sheet}'")
        else:
            raise ValueError("sheet argument must be an int or string")
    else:
        sheet_to_operate_on = sheets[0]
    
    rows = list()
    first_row = True
    for table_row in sheet_to_operate_on.iter(table_row_tag):
        row = list()

        for data_cell in table_row.iter(cell_tag):
            text = data_cell.findtext(p_tag) if data_cell.findtext(p_tag) is not None else ''
            row.append(text)

        if row[-1] == '':
            row.pop(-1)     #remove row padding
        if len(row) == 0:
            continue        #remove empty row

        if (skip_lines is not None and skip_lines > 0):
            if header and first_row:
               first_row = False
            else: 
                skip_lines = skip_lines - 1
                continue
            
        rows.append(row)

    column_types = list()
    if 'column_types' in kwargs.keys():
        column_types = kwargs.get('column_types')

    if header is True:
        columns = rows[0]   #creating a column row for agate
    else:
        if 'column_names' in kwargs.keys():
            columns = kwargs.get('column_names')
        else:
            raise ValueError('column_names argument must be provided if header is set to be False')
    
    rows.pop(0)         #removing extra column row from data
    table = agate.Table(rows=rows,column_names=columns)
    return table

agate.Table.from_ods = classmethod(from_ods)