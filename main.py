# Developed with pleasure in PyCharm IDE

from typing import List
from util import print_colored
from cli import CLI, Settings
from aggregator import Aggregator, Report
from reader import Table, PDFReader
from writer import XslxWriter


def main():
    settings: Settings = CLI().get_settings()
    pdf_reader: PDFReader = PDFReader(settings)

    tables: List[Table] = pdf_reader.extract_data_from_pdfs()
    reports: List[Report] = Aggregator().generate_reports(tables)

    writer: XslxWriter = XslxWriter(settings)
    writer.generate_xlsx(reports)
    print_colored("Done!", "green")

if __name__ == "__main__":
    main()

