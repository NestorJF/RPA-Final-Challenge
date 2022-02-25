from cmath import log
from turtle import position

from libraries.common import act_on_element, capture_page_screenshot, log_message, check_file_download_complete, file_system, pdf, files
from config import OUTPUT_FOLDER
from datetime import datetime

class GOBPE():

    def __init__(self, rpa_selenium_instance, credentials: dict):
        self.browser = rpa_selenium_instance
        self.GOBPE_url = credentials["url"]
        self.files_to_download_dict_list = []
        self.downloaded_files_dict_list = []
        self.result_text = ""

    def access_gobpe(self):
        """
        Access GOBPE from the browser.
        """
        log_message("Start - Access GOBPE")
        self.browser.go_to(self.GOBPE_url)
        log_message("End - Access GOBPE")

    def go_to_onpe_reports(self):
        """
        Go to the ONPE reports section
        """
        act_on_element('//ul[@class="list-footer"]/li/a[@href="/estado"]', 'click_element')
        act_on_element('//a[@class="power-card" and @href="/estado/organismos-autonomos"]', 'click_element')
        act_on_element('//article[@id="onpe"]//a', 'click_element')
        self.switch_to_specified_section()
        act_on_element('//a[@data-origin="onpe-publicaciones-ver-más-link"]', 'click_element')
        

    def switch_to_specified_section(self):
        """
        Select the menu specified in the text file
        """
        category = file_system.read_file("Category.txt", encoding = "utf-8")
        act_on_element('//li[@class="navigator__item"]/a[text() = "{}"]'.format(category), 'click_element')

    def search_onpe_reports(self):
        """
        Search the ONPE reports with a defined time range
        """
        since_date = "28-10-2021"
        until_date = datetime.today().strftime("%d-%m-%Y")
        url = "https://www.gob.pe/busquedas?contenido[]=publicaciones&desde={}&hasta={}&institucion[]=onpe&sheet=1&sort_by=recent".format(since_date, until_date)
        self.browser.go_to(url)

    def download_Files(self):
        """
        Donwloads files from GOBPE.
        """
        report_elements = act_on_element('//ul[@class="results"]/li', "find_elements")
        for report_element in report_elements:
            title_report = report_element.find_element_by_xpath('.//h3').text
            download_required = next((report_dict['Download Required'] for report_dict in self.excel_data_dict_list if title_report == report_dict['Name']), "")
            if download_required.lower() == "yes":
                download_button = report_element.find_element_by_xpath('.//a[text() = "Descargar"]')
                act_on_element(download_button, 'click_element')
                check_file_download_complete("pdf", 20)


    def read_pdf_reports(self):
        """
        Read PDF Reports and extract files names, amount of pages.
        """
        files_downloaded = file_system.find_files("{}/*.{}".format(OUTPUT_FOLDER, "pdf"))
        for file_downloaded in files_downloaded:
            file_name = file_system.get_file_name(file_downloaded)
            text_dict = pdf.get_text_from_pdf(file_downloaded)
            pages_amount = len(text_dict)

            report_dict = {
                "File Name": file_name,
                "Amount of pages": pages_amount
            }
            self.downloaded_files_dict_list.append(report_dict)

            if pages_amount > 50:
                self.result_text = self.result_text + "File Name: " + file_name + "\n"
                self.result_text = self.result_text + "----------------------" + "\n"


    def write_data_excel(self):
        """
        Writes the results to an excel file.
        """
        files.create_workbook(path = "{}/Results.xlsx".format(OUTPUT_FOLDER))
        files.rename_worksheet("Sheet", "Results")
        files.append_rows_to_worksheet(self.downloaded_files_dict_list, name = "Results", header = True, start = None)
        files.save_workbook(path = None)
        files.close_workbook()

    def write_result_txt(self):
        """
        Writes in a txt the file name of reports with the previous matched condition (more than 50 pages)
        """
        file_system.create_file("{}/Results.txt".format(OUTPUT_FOLDER), content = self.result_text, encoding = 'utf-8', overwrite = True)

    def read_files_to_download_excel(self):
            files.open_workbook("Files_To_Download.xlsx")
            self.excel_data_dict_list = files.read_worksheet(name = "Sheet1", header = True)
            files.close_workbook()