# docx2xml
Convert docx to XML for easier parsing and analysis.

1. Prerequisites:<br>
`pip install python-docx`

2. Running the script directly:<br>
`python docx_to_xml_converter.py path/to/input.docx`

3. Running the script directly with an output path:<br>
`python docx_to_xml_converter.py path/to/input.docx -o path/to/output.xml`

4. Running the script with verbose logging: <br>
`python docx_to_xml_converter.py path/to/input.docx -o path/to/output.xml --verbose`

5. Importing as a module:<br>
```python
from docx_to_xml_converter import convert_docx_to_xml
input_file = 'path/to/input.docx'
output_file = 'path/to/output.xml'
convert_docx_to_xml(input_file, output_file)\
```

6. Exmaple output: <br>
```xml
<?xml version='1.0' encoding='utf-8'?>
<Document>
    <Paragraphs>
        <Paragraph id="1">
            <Style>Normal</Style>
            <TextElements>
                <Run bold="False" italic="False" underline="False">This is an example</Run>
                <NonBreakingSpace id="2"/>
                <Run bold="False" italic="False" underline="False">with non-breaking spaces.</Run>
            </TextElements>
            <Breaks>
                <Page_break>True</Page_break>
            </Breaks>
        </Paragraph>
        <Paragraph id="2">
            <Style>Heading 1</Style>
            <TextElements>
                <Run bold="True" italic="False" underline="False">Section 1</Run>
            </TextElements>
            <Breaks>
                <Section_break>True</Section_break>
            </Breaks>
        </Paragraph>
    </Paragraphs>
    <Tables>
        <Table id="1">
            <Row id="1">
                <Cell id="1">Header 1</Cell>
                <Cell id="2">Header 2</Cell>
            </Row>
            <Row id="2">
                <Cell id="1">Data 1</Cell>
                <Cell id="2">Data 2</Cell>
            </Row>
        </Table>
    </Tables>
</Document>

```
