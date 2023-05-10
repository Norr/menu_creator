from openpyxl import load_workbook
import os

import jinja2
import pdfkit
from pathlib import Path

config = pdfkit.configuration(wkhtmltopdf=r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe')

def menu_to_dict(filename:str) -> dict:
    """Function returns dict of excel file contains menu data.

    :param filename: Excel file contains menu data.
    :type filename: str
    :return: Dict of data.
    :rtype: dict
    """

    workbook = load_workbook(filename=filename)
    sections = workbook.sheetnames
    menu_data = {}
    for section in sections:
        menu_data[section] = tuple(element for element in workbook[section].values)
    return menu_data

def _html2pdf(html_file, output_file, **options):
    """
    Function that create pdf file from rendered html template.
    :return: PDF file
    """
    with open(file=html_file) as f:
        pdfkit.from_file(f, output_file, options=options, configuration=config)
#test

def render_pdf(format:str='A4'):
        """
        Method that creating html page with email contents like Microsoft Outlook print style.
        :return: file rendered.html
        """
        template_loader = jinja2.FileSystemLoader(searchpath=os.path.join(os.getcwd()))
        template_env = jinja2.Environment(loader=template_loader)
        template_file = 'template.html'
        rendered_file = 'rendered.html'
        template = template_env.get_template(template_file)
        output_data = template.render(menu_to_dict(filename='menu.xlsx'))
        with open(rendered_file, mode='w', encoding='utf8') as rendered:
            rendered.write(output_data)
        _html2pdf(rendered_file, f"menu_{format}.pdf")


