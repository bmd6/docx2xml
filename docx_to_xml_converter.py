#!/usr/bin/env python3
"""
docx_to_xml_converter.py

A Python module to convert .docx files to XML format.
Handles section headers with different levels, paragraphs (including step numbers and letters),
and tables. Omits review comments and tracked changes for simplicity.

Date: 2024-10-20
"""

import os
import sys
import logging
from typing import Any, Dict, List, Optional
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
import xml.etree.ElementTree as ET

# Configure logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)  # Set to DEBUG for detailed logging

# Create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)  # Change to DEBUG to see detailed logs on console

# Create file handler which logs even debug messages
fh = logging.FileHandler('docx_to_xml_converter.log')
fh.setLevel(logging.DEBUG)

# Create formatter and add it to the handlers
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
ch.setFormatter(formatter)
fh.setFormatter(formatter)

# Add the handlers to the logger
logger.addHandler(ch)
logger.addHandler(fh)


class DocxParser:
    """
    A class to parse .docx files and extract their contents, including
    paragraphs, tables, and section headers.
    """

    def __init__(self, filepath: str):
        """
        Initializes the DocxParser with the path to the .docx file.

        :param filepath: Path to the .docx file to be parsed.
        """
        self.filepath = filepath
        self.document = None
        logger.debug(f"DocxParser initialized with file: {self.filepath}")

    def load_document(self) -> None:
        """
        Loads the .docx document.

        :raises FileNotFoundError: If the file does not exist.
        :raises Exception: If the file cannot be opened as a .docx.
        """
        logger.debug("Attempting to load the .docx document.")
        if not os.path.isfile(self.filepath):
            logger.error(f"File not found: {self.filepath}")
            raise FileNotFoundError(f"File not found: {self.filepath}")
        try:
            self.document = Document(self.filepath)
            logger.info(f"Successfully loaded document: {self.filepath}")
        except Exception as e:
            logger.error(f"Error loading document: {e}")
            raise

    def parse_document(self) -> Dict[str, Any]:
        """
        Parses the loaded document and extracts its content, including
        elements and tables.

        :return: A dictionary containing parsed elements and tables.
        :raises Exception: If the document is not loaded or cannot be parsed.
        """
        logger.debug("Starting to parse the document.")
        if self.document is None:
            logger.error("Document not loaded. Call load_document() first.")
            raise Exception("Document not loaded. Call load_document() first.")

        parsed_content = {
            'elements': self.parse_elements(),
            'tables': self.parse_tables()
        }
        logger.info("Document parsing completed successfully.")
        return parsed_content

    def parse_elements(self) -> List[Dict[str, Any]]:
        """
        Parses all elements in the document in order, identifying whether
        each is a header or a paragraph.

        :return: A list of dictionaries containing parsed elements.
        """
        logger.debug("Parsing all elements in the document.")
        parsed_elements = []
        total_paragraphs = len(self.document.paragraphs)
        logger.info(f"Total paragraphs to parse: {total_paragraphs}")

        for idx, para in enumerate(self.document.paragraphs, start=1):
            para_text = para.text.strip()
            if not para_text:
                logger.debug(f"Encountered an empty paragraph at index {idx}. Skipping.")
                continue  # Skip empty paragraphs

            # Check if the paragraph is a header
            style = para.style
            if style.type == WD_STYLE_TYPE.PARAGRAPH and style.name.startswith('Heading'):
                # Extract header level
                try:
                    level = int(style.name.split(' ')[1])
                except (IndexError, ValueError):
                    level = 1  # Default to level 1 if parsing fails
                    logger.warning(f"Unable to determine header level for style: {style.name}. Defaulting to level 1.")
                parsed_elements.append({
                    'type': 'header',
                    'level': level,
                    'text': para_text
                })
                logger.debug(f"Parsed Header level {level}: {para_text}")
            else:
                # Regular paragraph
                parsed_elements.append({
                    'type': 'paragraph',
                    'id': str(idx),
                    'text': para_text
                })
                logger.debug(f"Parsed Paragraph {idx}: {para_text}")

            # Provide progress feedback every 10 paragraphs
            if idx % 10 == 0 or idx == total_paragraphs:
                logger.info(f"Parsed {idx}/{total_paragraphs} paragraphs.")

        return parsed_elements

    def parse_tables(self) -> List[Dict[str, Any]]:
        """
        Parses the tables in the document.

        :return: A list of dictionaries containing parsed table data.
        """
        logger.debug("Parsing tables.")
        parsed_tables = []
        total_tables = len(self.document.tables)
        logger.info(f"Total tables to parse: {total_tables}")

        for tbl_idx, table in enumerate(self.document.tables, start=1):
            table_dict = {
                'id': str(tbl_idx),
                'rows': []
            }
            for row_idx, row in enumerate(table.rows, start=1):
                row_data = []
                for cell_idx, cell in enumerate(row.cells, start=1):
                    # Concatenate all paragraph texts within the cell, separated by newlines
                    cell_text = '\n'.join([para.text.strip() for para in cell.paragraphs if para.text.strip()])
                    row_data.append({
                        'id': str(cell_idx),
                        'text': cell_text
                    })
                table_dict['rows'].append({
                    'id': str(row_idx),
                    'cells': row_data
                })
            parsed_tables.append(table_dict)
            logger.debug(f"Parsed Table {tbl_idx}.")
            
            # Provide progress feedback every 5 tables
            if tbl_idx % 5 == 0 or tbl_idx == total_tables:
                logger.info(f"Parsed {tbl_idx}/{total_tables} tables.")

        return parsed_tables


class XMLConverter:
    """
    A class to convert parsed document data into an XML structure,
    including headers and paragraphs, and tables.
    """

    def __init__(self, data: Dict[str, Any]):
        """
        Initializes the XMLConverter with parsed data.

        :param data: Parsed data from the .docx file.
        """
        self.data = data
        logger.debug("XMLConverter initialized with parsed data.")

    def build_xml(self) -> ET.Element:
        """
        Builds the XML ElementTree from the parsed data.

        :return: The root XML element.
        """
        logger.debug("Starting to build XML structure.")
        root = ET.Element('Document')

        # Add Elements (Headers and Paragraphs)
        elements_element = ET.SubElement(root, 'Elements')
        for element in self.data.get('elements', []):
            if element['type'] == 'header':
                header_element = ET.SubElement(elements_element, 'Header', level=str(element['level']))
                header_element.text = element['text']
                logger.debug(f"Added Header level {element['level']}: {element['text']}")
            elif element['type'] == 'paragraph':
                para_element = ET.SubElement(elements_element, 'Paragraph', id=element['id'])
                para_element.text = element['text']
                logger.debug(f"Added Paragraph {element['id']}: {element['text']}")
            else:
                logger.warning(f"Unknown element type: {element['type']}")

        # Add Tables
        tables_element = ET.SubElement(root, 'Tables')
        for table in self.data.get('tables', []):
            table_element = ET.SubElement(tables_element, 'Table', id=table['id'])
            for row in table['rows']:
                row_element = ET.SubElement(table_element, 'Row', id=row['id'])
                for cell in row['cells']:
                    cell_element = ET.SubElement(row_element, 'Cell', id=cell['id'])
                    cell_element.text = cell['text']
            logger.debug(f"Added Table {table['id']}.")

        logger.info("XML structure built successfully.")
        return root

    def save_xml(self, root: ET.Element, output_path: str) -> None:
        """
        Saves the XML ElementTree to a file.

        :param root: The root XML element.
        :param output_path: Path where the XML file will be saved.
        :raises Exception: If the file cannot be written.
        """
        logger.debug(f"Attempting to save XML to: {output_path}")
        try:
            tree = ET.ElementTree(root)
            tree.write(output_path, encoding='utf-8', xml_declaration=True)
            logger.info(f"XML file saved successfully at: {output_path}")
        except Exception as e:
            logger.error(f"Error saving XML file: {e}")
            raise


class Converter:
    """
    A facade class to convert a .docx file to an XML file, handling parsing and conversion.
    """

    def __init__(self, input_path: str, output_path: Optional[str] = None):
        """
        Initializes the Converter with input and output paths.

        :param input_path: Path to the input .docx file.
        :param output_path: Path to the output XML file. If None, replaces .docx with .xml.
        """
        self.input_path = input_path
        if output_path:
            self.output_path = output_path
        else:
            base, _ = os.path.splitext(input_path)
            self.output_path = f"{base}.xml"
        logger.debug(f"Converter initialized with input: {self.input_path}, output: {self.output_path}")

    def convert(self) -> None:
        """
        Performs the conversion from .docx to XML.

        :raises Exception: If any step in the conversion fails.
        """
        logger.info(f"Starting conversion: {self.input_path} -> {self.output_path}")
        parser = DocxParser(self.input_path)
        parser.load_document()
        parsed_data = parser.parse_document()

        converter = XMLConverter(parsed_data)
        xml_root = converter.build_xml()
        converter.save_xml(xml_root, self.output_path)
        logger.info("Conversion completed successfully.")


def convert_docx_to_xml(input_path: str, output_path: Optional[str] = None) -> None:
    """
    Converts a .docx file to an XML file.

    :param input_path: Path to the input .docx file.
    :param output_path: Path to the output XML file. If None, replaces .docx with .xml.
    """
    try:
        converter = Converter(input_path, output_path)
        converter.convert()
    except Exception as e:
        logger.exception(f"Failed to convert {input_path} to XML. Error: {e}")
        raise


def main():
    """
    The main function to execute when running the script directly.
    Parses command-line arguments and performs the conversion.
    """
    import argparse

    parser = argparse.ArgumentParser(description='Convert a .docx file to XML format.')
    parser.add_argument('input', help='Path to the input .docx file.')
    parser.add_argument('-o', '--output', help='Path to the output XML file.')

    args = parser.parse_args()

    try:
        convert_docx_to_xml(args.input, args.output)
    except Exception as e:
        logger.error(f"An error occurred during conversion: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
