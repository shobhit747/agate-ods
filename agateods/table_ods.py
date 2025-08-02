"""
This module contains the ODS extension to :class:`Table <agate.table.Table>`.
"""

import agate

import zipfile
import lxml.etree as etree
import xml.etree.ElementTree as ET
import io
import re
from collections import OrderedDict

ODS_CONTENT_FILE = 'content.xml'

ODS_TYPE_TO_AGATE = {
    'string':agate.Text(),
    'float':agate.Number(),
    'currency':agate.Number(),
    'date':agate.DateTime(),
    'time':agate.DateTime(),
    'boolean':agate.Boolean(),
    'percentage':agate.Number()
}

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

def resolve_data_value(data_cell,ns):
    """
    Convert an ODS worksheet cell value to its appropriate Agate data type.
    """

    p_tag = '{%s}p' % ns['text']
    value_type_attr = '{%s}value-type' % ns['office']
    office_namespace = ns['office']
    resolve_value_attr = {
        'float': '{%s}value' % office_namespace,
        'percentage': '{%s}value' % office_namespace,
        'currency': '{%s}currency' % office_namespace,
        'date': '{%s}date-value' % office_namespace,
        'time': '{%s}time-value' % office_namespace,
        'boolean': '{%s}boolean-value' % office_namespace
    }

    def resolve_string_to_number(value):
        return re.sub(r"[^\d.]+", "", value)

    cell_data_type = data_cell.attrib.get(value_type_attr)
    if  cell_data_type == 'string':
        data_value = data_cell.findtext(p_tag) if data_cell.findtext(p_tag) is not None else ''
    elif cell_data_type == 'currency':
        data_value = resolve_string_to_number(
            data_cell.findtext(p_tag) if data_cell.findtext(p_tag) is not None else ''
            )
    else:
        data_value = data_cell.attrib.get(resolve_value_attr[cell_data_type])

    return data_value

def from_ods(cls,file_path, sheet=None, skip_lines=0, header=True, row_limit=None, **kwargs):
    """
    Parse an ODS file.

    :param path:
        Path to an ODS file to load.
    :param sheet:
        The name or integer indice of the worksheet to load. If not specified
        then the first sheet will be used.
    :param skip_lines:
        The number of rows to skip from the top of the sheet.
    :param header:
        If :code:`True`, the first row is assumed to contain column names.
    :param row_limit:
        Limit how many rows of data will be read.
    """
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
    
    value_type_attr = '{%s}value-type' % ns['office']
    value_attr = '{%s}value' % ns['office']
    padding_attr = '{%s}number-columns-repeated' % ns['table']

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
    column_types = list()
    calculated_column_types = dict()
    for table_row in sheet_to_operate_on.iter(table_row_tag):   #iterate through rows of the table
        row = list()
        
        column_number = 0
        for data_cell in table_row.iter(cell_tag):
            cell_data_type = data_cell.attrib.get(value_type_attr)
            if padding_attr not in data_cell.attrib.keys(): #remove row padding
                data_value = resolve_data_value(data_cell,ns)
                row.append(data_value)

                if header and first_row:
                    continue
                if column_number not in calculated_column_types.keys():
                    calculated_column_types[column_number] = cell_data_type
                else:
                    if calculated_column_types[column_number] != cell_data_type:
                        raise TypeError(f"Type mismatch at row {len(rows)} column {column_number}")
                column_number = column_number + 1
                
        if len(row) == 0:
            continue        #remove empty row
        
        if row_limit is not None and skip_lines <= 0:
            if row_limit > 0:
                if not first_row or not header :
                    row_limit = row_limit - 1
            else:
                break

        if skip_lines is not None and skip_lines > 0:
            if header and not first_row:
                skip_lines = skip_lines - 1
                continue
            elif not header:
                skip_lines = skip_lines - 1
                continue

        if first_row:
            first_row = False

        rows.append(row)

    column_types = list()
    if 'column_types' in kwargs.keys():
        column_types = kwargs.get('column_types')
    else:
        calculated_column_types = OrderedDict(sorted(calculated_column_types.items()))
        data_types = calculated_column_types.values()
        column_types = [ODS_TYPE_TO_AGATE[data_type] for data_type in data_types]

    if header is True:
        columns = rows[0]   #creating a column row for agate
        rows.pop(0)         #removing extra column row from data
    else:
        if 'column_names' in kwargs.keys():
            columns = kwargs.get('column_names')
        else:
            raise ValueError('column_names argument must be provided if header is set to be False')
    
    table = agate.Table(rows=rows,column_names=columns)
    return table

agate.Table.from_ods = classmethod(from_ods)