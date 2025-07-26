"""
This module contains the ODS extension to :class:`Table <agate.table.Table>`.
"""

import agate

import zipfile
import lxml.etree as etree
import xml.etree.ElementTree as ET
import pathlib
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


