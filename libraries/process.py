import time
from libraries.common import log_message, capture_page_screenshot, browser
from libraries.gobpe.gobpe import GOBPE

from config import OUTPUT_FOLDER, tabs_dict


class Process:
    def __init__(self, credentials: dict):
        log_message("Initialization")
        prefs = {
            "profile.default_content_setting_values.notifications": 2,
            "profile.default_content_settings.popups": 0,
            "directory_upgrade": True,
            "download.default_directory": OUTPUT_FOLDER,
            "plugins.always_open_pdf_externally": True,
            "download.prompt_for_download": False
        }
        browser.open_available_browser(preferences = prefs)
        browser.set_window_size(1920, 1080)
        browser.maximize_browser_window()
    
        gobpe = GOBPE(browser, {"url": "https://www.gob.pe/"})
        gobpe.access_gobpe()
        self.gobpe = gobpe 

    def start(self):

        log_message("Start - GO TO ONPE REPORTS")
        self.gobpe.go_to_onpe_reports()
        log_message("End - GO TO ONPE REPORTS")
        log_message("Start - SEARCH ONPE REPORTS")
        self.gobpe.search_onpe_reports()
        log_message("End - SEARCH ONPE REPORTS")
        log_message("Start - Read Files to Download Excel")
        self.gobpe.read_files_to_download_excel()
        log_message("End - Read Files to Download Excel")
        log_message("Start - Download files")
        self.gobpe.download_Files()
        log_message("End - Download files")
        log_message("Start - Read PDF Reports")
        self.gobpe.read_pdf_reports()
        log_message("End - Read PDF Reports")
        log_message("Start - Write data to Excel")
        self.gobpe.write_data_excel()
        log_message("End - Write data to Excel")
        log_message("Start - Write Result txt")
        self.gobpe.write_result_txt()
        log_message("End - Write Result txt")

    def finish(self):
        log_message("DW Process Finished")
        browser.close_browser()