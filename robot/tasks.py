from RPA.Robocorp.WorkItems import WorkItems
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from datetime import datetime
from dateutil.relativedelta import relativedelta
import re
import os
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import random

items = WorkItems()
browser_lib = Selenium()
browser_lib.auto_close = True

def random_timer():
    return 50 + float((random.randint(0, 50)/2))



class Step1:
    def __init__(self):
        items.get_input_work_item()
        self.input_items = items.get_work_item_variables()

    def setting_scraping_inputs(self):
        sections = random.choice(self.input_items["sections"])
        search_phrase = random.choice(self.input_items["search_phrase"])
        n_months = random.randint(0, 3)

        scraping_data_dict = {
            "sections": sections,
            "search_phrase": search_phrase,
            "n_months": n_months
        }
        return scraping_data_dict

class Step2:

    def __init__(self, previous_step_data):

        self.url = 'https://www.reuters.com/'
        self.input_items = previous_step_data
        self.sections = self.input_items["sections"]
        self.search_phrase = self.input_items["search_phrase"]
        self.n_months = self.input_items["n_months"]

    def open_website(self):

        try:
            browser_lib.open_available_browser(
                self.url, browser_selection="chrome", maximized=True)
        except Exception as e:
            print(e)
            browser_lib.close_all_browsers()

    def enter_search_phrase(self):
        
        try:
            browser_lib.set_browser_implicit_wait(60)
            browser_lib.click_button_when_visible(
                "xpath:/html/body/div[1]/header/div/div/div/div/div[3]/div[2]/button")
            browser_lib.set_browser_implicit_wait(70)
            browser_lib.input_text_when_element_is_visible(
                "xpath:/html/body/div[1]/header/div/div/div/div/div[3]/div[2]/div/div/input", self.search_phrase)
            browser_lib.set_browser_implicit_wait(80)
            browser_lib.click_button_when_visible(
                "xpath:/html/body/div[1]/header/div/div/div/div/div[3]/div[2]/div/button[2]")
        except Exception as e:
            print(e)
            browser_lib.close_all_browsers()
            
    def section_selection(self):
        browser_lib.set_browser_implicit_wait(65)
        browser_lib.click_button_when_visible(
            "xpath:/html/body/div[1]/div[2]/div[1]/div/div/div/div/form/div/div[1]/button"
        )
        list_of_sections = browser_lib.find_elements("xpath:/html/body/div[1]/div[2]/div[1]/div/div/div/div/form/div/div[1]/div[2]/ul")
        for li in list_of_sections:
            if li.text == self.sections:
                browser_lib.set_browser_implicit_wait(50)
                browser_lib.click_element(li)
            

class Step3():
    def __init__(self, previous_step_data):
        self.input_items = previous_step_data
        self.months = self.input_items["n_months"]
        self.search_phrase = self.input_items["search_phrase"]
        self.workbook = Workbook()


    def is_date_within_interval(self, text):
        current_time = datetime.now()
        oldest_time = 0
        if self.months == 0:
            oldest_time = 1
        else:
            oldest_time = self.months
        correct_oldest_month = current_time - relativedelta(months=oldest_time)
        date_pattern = r"(\w+\s\d{1,2},\s\d{4})"
        match = re.search(date_pattern, text)
        if match:
            matched_date_str = match.group(1)
            parsed_date = datetime.strptime(matched_date_str, "%B %d, %Y")
            return (bool(correct_oldest_month <= parsed_date <= current_time),
                    parsed_date)
        else:
            return (False, [])

    def count_search_phrase_occurrences(self, search_phrase, text):

        pattern = re.compile(r'\b{}\b'.format(re.escape(search_phrase)))
        matches = re.findall(pattern, text)

        return len(matches)

    def contains_money(self, title, description):

        money_pattern = r'\$[\d,.]+|\d+\s*(?:dollars|USD)'

        return bool(re.search(money_pattern, title) or
                    re.search(money_pattern, description))

    def insert_images_to_excel(self):

        current_folder = os.getcwd()
        image_folder = os.path.join(current_folder, 'output')
        worksheet = self.workbook.active
        column = 'H'
        counter = 2

        for image_file in (os.listdir(image_folder)):
            if image_file.endswith('.png'):
                image_path = os.path.join(image_folder, image_file)
                img = Image(image_path)
                img.anchor = f'{column}{counter}'
                worksheet.add_image(img)
                counter += counter
        self.workbook.save("output_excel_file.xlsx")

    def iterate_through_news(self):
        browser_lib.set_browser_implicit_wait(random_timer())
        news_items = browser_lib.find_elements("xpath:/html/body/div[1]/div[2]/div[2]/div/div[2]/div[2]/ul")
        counter = 0
        titles = []
        dates = []
        descriptions = []
        picture_filename = []
        count_of_search_phrase_in_title = []
        count_of_search_phrase_in_description = []
        contains_money_reference = []
        read_page = 1
        while (read_page):

            for news in news_items:
                texts = browser_lib.get_text(news)
                image = news.find_element("tag:img")
                text_splitted = texts.split('\n')
                browser_lib.set_browser_implicit_wait(random_timer())
                browser_lib.click_button_when_visible("xpath:/html/body/div[1]/div[2]/div[1]/div/div/div/div/form/div/div[2]/button")
                if(self.months == 0 or self.months == 1):
                    browser_lib.set_browser_implicit_wait(random_timer())
                    browser_lib.click_element_when_clickable("xpath/html/body/div[1]/div[2]/div[1]/div/div/div/div/form/div/div[2]/div[2]/ul/li[4]")
                else:
                    browser_lib.set_browser_implicit_wait(random_timer())
                    browser_lib.click_element_when_clickable("xpath:/html/body/div[1]/div[2]/div[1]/div/div/div/div/form/div/div[2]/div[2]/ul/li[5]")
                        
                date_of_interest = (self.is_date_within_interval
                                    (text_splitted[1])[0],
                                    self.is_date_within_interval
                                    (text_splitted[1])[1])

                if (date_of_interest[0]):
                    titles.append(text_splitted[0])
                    dates.append(date_of_interest[1].strftime("%B %d, %Y"))
                    descriptions.append(text_splitted[2])
                    browser_lib.set_browser_implicit_wait(random_timer())
                    picture_filename.append(
                        browser_lib.get_element_attribute(image, "src"))
                    browser_lib.set_browser_implicit_wait(random_timer())
                    browser_lib.capture_element_screenshot(
                        image, "output/screenshot{input}.png"
                        .format(input=counter))
                    count_of_search_phrase_in_title.append(
                        self.count_search_phrase_occurrences
                        (self.search_phrase, text_splitted[0]))
                    count_of_search_phrase_in_description.append(
                        self.count_search_phrase_occurrences
                        (self.search_phrase, text_splitted[2]))
                    contains_money_reference.append(
                        self.contains_money
                        (text_splitted[0], text_splitted[2]))
                    counter += 1
            if len(titles) % 20 == 0 and len(titles) != 0:
                browser_lib.set_browser_implicit_wait(random_timer())
                browser_lib.click_link(
                    "xpath:/html/body/div[1]/div[2]/div[2]/div/div[2]/div[3]/button[2]")
            else:
                read_page = 0

            full_list = [titles, dates, descriptions, picture_filename,
                         count_of_search_phrase_in_title,
                         count_of_search_phrase_in_description,
                         contains_money_reference]

            full_list_transposed = list(zip(*full_list))

            worksheet = self.workbook.active
            headers = ['Title', 'Date', 'Description', 'Picture Source',
                       'Count of Search Phrase in Title',
                                'Count of Search Phrase in Description',
                                'Contains Money Reference', 'Image']

            worksheet.append(headers)
            for row_data in full_list_transposed:
                worksheet.append(row_data)
            self.insert_images_to_excel()



if __name__ == "__main__":

    step1 = Step1()
    data_for_step2 = step1.setting_scraping_inputs()
    step2 = Step2(data_for_step2)
    step2.open_website()
    step2.enter_search_phrase()
    step2.section_selection()
    step3 = Step3(data_for_step2)
    step3.iterate_through_news()
    items.create_output_work_item(files="output_excel_file.xlsx", save=True)
