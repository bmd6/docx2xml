import os
import sys
import logging
import xml.etree.ElementTree as ET
from xml.dom import minidom
from typing import Optional, Dict, List
import pythoncom
import win32com.client
import re
from dataclasses import dataclass, field

# Configure logging
def setup_logging(verbose: bool = False):
    """
    Sets up logging to file and console.

    :param verbose: If True, sets console logging to INFO level. Else, WARNING.
    """
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    # File handler for detailed logs
    file_handler = logging.FileHandler('docx_to_xml.log', mode='a', encoding='utf-8')
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)

    # Console handler for status updates
    console_handler = logging.StreamHandler(sys.stdout)
    if verbose:
        console_level = logging.INFO
    else:
        console_level = logging.WARNING
    console_formatter = logging.Formatter('%(levelname)s: %(message)s')
    console_handler.setFormatter(console_formatter)
    console_handler.setLevel(console_level)
    logger.addHandler(console_handler)

@dataclass
class ListItem:
    number: str
    text: str
    level: int
    children: List['ListItem'] = field(default_factory=list)

class DocxToXmlConverter:
    """
    A class to convert a DOCX file to an XML file, extracting multi-level lists and section headers.
    """

    def __init__(self, input_path: str, output_path: Optional[str] = None):
        """
        Initializes the converter with input and output paths.

        :param input_path: Path to the input DOCX file.
        :param output_path: Path to the output XML file. If None, replaces .docx with .xml.
        """
        self.input_path = input_path
        self.output_path = output_path or self._generate_output_path()
        self.word_app = None

    def _generate_output_path(self) -> str:
        """
        Generates the output XML file path based on the input DOCX path.

        :return: Output XML file path.
        """
        base, _ = os.path.splitext(self.input_path)
        return f"{base}.xml"

    def _initialize_word(self):
        """
        Initializes the Word application object.

        Raises:
            Exception: If Word application cannot be initialized.
        """
        try:
            logging.info("Initializing Word application...")
            pythoncom.CoInitialize()
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = True
            self.word_app.DisplayAlerts = 0  # wdAlertsNone
            logging.info("Word application initialized successfully.")
        except Exception as e:
            logging.error(f"Failed to initialize Word application: {e}")
            raise

    def _cleanup_word(self):
        """
        Closes the Word application.
        """
        try:
            if self.word_app:
                logging.info("Closing Word application...")
                self.word_app.Quit()
                logging.info("Word application closed.")
        except Exception as e:
            logging.warning(f"Error while closing Word application: {e}")

    def convert(self):
        """
        Performs the conversion from DOCX to XML.
        """
        try:
            self._initialize_word()
            logging.info(f"Opening DOCX file: {self.input_path}")
            try:
                doc = self.word_app.Documents.Open(
                    FileName=self.input_path,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False
                )
                logging.info("DOCX file opened successfully.")
            except Exception as e:
                logging.error(f"Failed to open DOCX file: {e}")
                raise

            content = self._extract_content(doc)
            logging.info("Content extraction completed.")

            root = self._build_xml(content)
            logging.info("XML structure built successfully.")

            xml_str = self._prettify_xml(root)
            logging.info("XML string prepared for writing.")

            logging.info(f"Writing XML to file: {self.output_path}")
            try:
                with open(self.output_path, 'w', encoding='utf-8') as f:
                    f.write(xml_str)
                logging.info("XML file written successfully.")
            except Exception as e:
                logging.error(f"Failed to write XML file: {e}")
                raise

            doc.Close(False)
            logging.info("DOCX document closed.")

        except Exception as e:
            logging.error(f"An error occurred during conversion: {e}", exc_info=True)
            raise
        finally:
            self._cleanup_word()

    def _extract_content(self, doc) -> Dict[str, Dict[str, any]]:
        """
        Extracts headers and multi-level lists from the Word document.

        :param doc: The opened Word document.
        :return: Dictionary with headers as keys and their details (level and items) as values.
        """
        content = {}
        current_header = None
        current_header_level = 1
        current_items = []
        list_stack = []  # Stack to manage list hierarchy
        total_paragraphs = doc.Paragraphs.Count
        processed_paragraphs = 0

        logging.info(f"Total paragraphs to process: {total_paragraphs}")

        for para in doc.Paragraphs:
            processed_paragraphs += 1
            if processed_paragraphs % 50 == 0 or processed_paragraphs == total_paragraphs:
                logging.info(f"Processed {processed_paragraphs}/{total_paragraphs} paragraphs.")

            # Ignore if the paragraph is a revision or comment
            if self._is_revision_or_comment(para):
                continue

            style = para.Style.NameLocal
            text = para.Range.Text.strip()
            logging.debug(f"Processing paragraph {processed_paragraphs}: Style='{style}', Text='{text}'")

            if self._is_heading(style):
                if current_header:
                    # Add the previous header and its items (even if items are empty)
                    content[current_header] = {
                        'level': current_header_level,
                        'items': current_items
                    }
                    logging.info(f"Added header: '{current_header}' with level {current_header_level} and {len(current_items)} list items.")
                current_header = text
                current_header_level = self._get_heading_level_from_style(style)
                current_items = []
                list_stack = []
                logging.info(f"Detected header: '{current_header}' with level {current_header_level}")
                continue

            if not text:
                continue

            if para.Range.ListFormat.ListType != 0:  # 0 means no list
                list_item = self._create_list_item(para)
                if list_item:
                    self._add_list_item_to_content(list_item, current_items, list_stack)
            else:
                # Handle non-list paragraphs if needed
                continue

        # Add the last header and its items (even if items are empty)
        if current_header:
            content[current_header] = {
                'level': current_header_level,
                'items': current_items
            }
            logging.info(f"Added header: '{current_header}' with level {current_header_level} and {len(current_items)} list items.")

        logging.info("Completed content extraction.")
        return content

    def _is_revision_or_comment(self, para) -> bool:
        """
        Checks if the paragraph is a revision or a comment.

        :param para: The paragraph object.
        :return: True if it's a revision or comment, False otherwise.
        """
        try:
            if para.Range.Revisions.Count > 0:
                logging.debug("Paragraph is a revision; ignoring.")
                return True
            if para.Range.Comments.Count > 0:
                logging.debug("Paragraph has comments; ignoring.")
                return True
        except Exception as e:
            logging.warning(f"Failed to check revisions/comments: {e}")
        return False

    def _is_heading(self, style: str) -> bool:
        """
        Determines if the paragraph style is a heading.

        :param style: The style name.
        :return: True if it's a heading style, False otherwise.
        """
        return style.startswith('Heading')

    def _get_heading_level_from_style(self, style: str) -> int:
        """
        Extracts the heading level from the style name.

        :param style: The style name (e.g., 'Heading 1', 'Heading 2').
        :return: Heading level as integer.
        """
        match = re.search(r'Heading\s+(\d+)', style, re.IGNORECASE)
        if match:
            level = int(match.group(1))
            logging.debug(f"Extracted heading level: {level} from style '{style}'")
            return level
        logging.debug(f"Failed to extract heading level from style '{style}'. Defaulting to 1.")
        return 1

    def _create_list_item(self, para) -> Optional[ListItem]:
        """
        Creates a ListItem object from a list paragraph.

        :param para: The paragraph object.
        :return: ListItem object or None.
        """
        try:
            text = para.Range.Text.strip()
            level = para.Range.ListFormat.ListLevelNumber
            list_string = para.Range.ListFormat.ListString

            # Clean the text by removing the list marker
            cleaned_text = re.sub(r'^\s*(?:\d+|[a-zA-Z])[).]\s*', '', text)

            list_item = ListItem(
                number=f"{list_string}",
                text=cleaned_text,
                level=level - 1  # Convert to 0-based indexing
            )
            logging.debug(f"Created ListItem: {list_item}")
            return list_item
        except Exception as e:
            logging.warning(f"Failed to create ListItem: {e}")
            return None

    def _add_list_item_to_content(self, list_item: ListItem, current_items: List[ListItem], list_stack: List[ListItem]):
        """
        Adds a ListItem to the current_items list, handling hierarchy based on level.

        :param list_item: The ListItem to add.
        :param current_items: The current list of ListItems under the current header.
        :param list_stack: Stack to manage current hierarchy levels.
        """
        try:
            # Adjust the stack to the current level
            while len(list_stack) > list_item.level:
                popped = list_stack.pop()
                logging.debug(f"Popped from stack: {popped.number}")

            if list_item.level == 0:
                current_items.append(list_item)
                list_stack.append(list_item)
                logging.info(f"Added ListItem to current_items: {list_item.number} {list_item.text}")
            else:
                if list_stack:
                    parent = list_stack[-1]
                    parent.children.append(list_item)
                    list_stack.append(list_item)
                    logging.info(f"Added ListItem as child to '{parent.number}': {list_item.number} {list_item.text}")
                else:
                    # If stack is empty, treat it as a top-level item
                    current_items.append(list_item)
                    list_stack.append(list_item)
                    logging.info(f"Added ListItem to current_items (no parent): {list_item.number} {list_item.text}")
        except Exception as e:
            logging.warning(f"Failed to add ListItem to content: {e}")

    def _build_xml(self, content: Dict[str, Dict[str, any]]) -> ET.Element:
        """
        Builds an XML Element from the extracted content.

        :param content: Dictionary with headers as keys and their details as values.
        :return: Root XML Element.
        """
        root = ET.Element('Document')

        for header, details in content.items():
            header_element = ET.SubElement(root, 'Header', attrib={'level': str(details['level'])})
            header_element.text = header
            logging.info(f"Added Header to XML: '{header}' with level {details['level']}")

            for item in details['items']:
                self._add_list_item_to_xml(header_element, item)

        logging.info("XML structure built successfully.")
        return root

    def _add_list_item_to_xml(self, parent_xml: ET.Element, list_item: ListItem):
        """
        Recursively adds ListItem elements to the XML.

        :param parent_xml: The parent XML element.
        :param list_item: The ListItem to add.
        """
        try:
            list_element = ET.SubElement(parent_xml, 'ListItem', attrib={
                'level': str(list_item.level),
                'marker': list_item.number
            })
            list_element.text = list_item.text
            logging.info(f"Added ListItem to XML: {list_item.number} {list_item.text}")

            for child in list_item.children:
                self._add_list_item_to_xml(list_element, child)

        except Exception as e:
            logging.warning(f"Failed to add ListItem to XML: {e}")

    def _prettify_xml(self, elem: ET.Element) -> str:
        """
        Returns a pretty-printed XML string for the Element.

        :param elem: The root XML element.
        :return: Pretty-printed XML string.
        """
        rough_string = ET.tostring(elem, 'utf-8')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ")

def extract_requirements(text: str) -> List[str]:
    """Extract requirement IDs from text."""
    return re.findall(r'\[([^\]]+)\]', text)

def main():
    """
    Main function to execute the conversion. Parses command-line arguments for input and output paths.
    """
    import argparse

    parser = argparse.ArgumentParser(description='Convert DOCX to XML, extracting lists and headers.')
    parser.add_argument('input', help='Path to the input DOCX file.')
    parser.add_argument('-o', '--output', help='Path to the output XML file. If not provided, replaces .docx with .xml.')
    parser.add_argument('--verbose', action='store_true', help='Enable verbose output to the console.')

    args = parser.parse_args()

    # Setup logging with optional verbosity
    setup_logging(verbose=args.verbose)

    if not os.path.isfile(args.input):
        logging.error(f"Input file does not exist: {args.input}")
        print(f"Error: Input file does not exist: {args.input}")
        sys.exit(1)

    converter = DocxToXmlConverter(input_path=args.input, output_path=args.output)
    try:
        logging.info("Starting conversion process...")
        converter.convert()
        print(f"Conversion successful. XML saved to: {converter.output_path}")
    except Exception as e:
        print(f"Conversion failed: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()
