import streamlit as st
import pandas as pd
# from docxtpl import DocxTemplate
from io import BytesIO
import zipfile
# import xml.etree.ElementTree as ET


def create_word_doc(df):
    """Create Word document without lxml dependency"""
    # Create minimal DOCX structure in memory
    buffer = BytesIO()

    with zipfile.ZipFile(buffer, 'w') as zf:
        # Create minimal document.xml
        document_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                <w:p><w:r><w:t>Data from Excel</w:t></w:r></w:p>
                <w:tbl>
                    <w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>
                    <w:tblGrid>
                        {col_grid}
                    </w:tblGrid>
                    {rows}
                </w:tbl>
            </w:body>
        </w:document>"""

        # Generate column grid
        col_grid = '\n'.join(['<w:gridCol w:w="2000"/>'] * len(df.columns))

        # Generate header row
        rows = '<w:tr>' + ''.join(
            f'<w:tc><w:p><w:r><w:t>{col}</w:t></w:r></w:p></w:tc>' for col in df.columns) + '</w:tr>'

        # Generate data rows
        for _, row in df.iterrows():
            rows += '<w:tr>' + ''.join(
                f'<w:tc><w:p><w:r><w:t>{val}</w:t></w:r></w:p></w:tc>' for val in row) + '</w:tr>'

        # Complete the document
        document_xml = document_xml.format(col_grid=col_grid, rows=rows)

        # Add required files to the ZIP
        zf.writestr('word/document.xml', document_xml)
        zf.writestr('[Content_Types].xml', """<?xml version="1.0" encoding="UTF-8"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>""")
        zf.writestr('_rels/.rels', """<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>""")

    buffer.seek(0)
    return buffer


def main():
    st.title("Excel to Word Converter (No lxml)")

    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx", "xls"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.dataframe(df.head())

        if st.button("Generate Word Document"):
            doc_buffer = create_word_doc(df)
            st.download_button(
                label="Download DOCX",
                data=doc_buffer,
                file_name="export.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )


if __name__ == "__main__":
    main()