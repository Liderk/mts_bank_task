from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from PIL import Image
import io
import pytesseract
import xlsxwriter


class FsspParser(object):
    def __init__(self, driver):
        self.driver = driver

    def parse(self, initials):
        initials_keys = ['surname', 'name', 'patronymic', 'birth']
        initials_value = initials.split()
        initials = dict(zip(initials_keys, initials_value))
        self._input_man_initials(**initials)
        header = self.driver.find_elements_by_tag_name('th')
        data = self.driver.find_elements_by_tag_name('td')
        self._write_xmls(header, data, **initials)

        # self.close()

    def _write_xmls(self, header, data, **initials):
        workbook = xlsxwriter.Workbook('hello.xlsx')
        worksheet = workbook.add_worksheet()
        row, col = 0, 0
        for item in header:
            worksheet.write(row, col, item.text)
            col += 1
        surname = initials['surname'].upper()
        for item in data[1:]:
            temp = item.text
            if surname in temp:
                row += 1
                col = 0
            worksheet.write(row, col, item.text)
            col += 1
        workbook.close()

    def _input_man_initials(self, **initials):
        self.driver.get('https://fssp.gov.ru/')
        WebDriverWait(self.driver, 1000000).until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@type='button']"))).click()
        self.driver.find_element_by_xpath(
            '//*[@id="app"]/main/section[1]/div/div/div/div/div[2]/div/div/div/form/div[8]/div[1]/div/div[1]/a').click()

        input_last_name = self.driver.find_element_by_xpath(
            "//input[@name='is[last_name]']")
        input_last_name.click()
        input_last_name.clear()
        input_last_name.send_keys(initials['surname'])

        input_first_name = self.driver.find_element_by_xpath(
            "//input[@name='is[first_name]']")
        input_first_name.click()
        input_first_name.clear()
        input_first_name.send_keys(initials['name'])

        input_patronymic_name = self.driver.find_element_by_xpath(
            "//input[@name='is[patronymic]']")
        input_patronymic_name.click()
        input_patronymic_name.clear()
        input_patronymic_name.send_keys(initials['patronymic'])

        input_birth_date = self.driver.find_element_by_xpath(
            "//input[@name='is[date]']")
        input_birth_date.click()
        input_birth_date.clear()
        input_birth_date.send_keys(initials['birth'])
        input_birth_date.submit()

        pass_capcha = self._pass_capcha()
        if pass_capcha:
            return True

    def _pass_capcha(self):

        pass_capcha = True

        while pass_capcha:
            self.driver.implicitly_wait(10)

            capcha_text = self.driver.find_element_by_xpath(
                '//input[@id="captcha-popup-code"]')
            capcha_text.click()
            capcha_text.clear()
            image = self.driver.find_element_by_id(
                "capchaVisual").screenshot_as_png
            imageStream = io.BytesIO(image)
            im = Image.open(imageStream)
            im.save('capcha.png')
            text = pytesseract.image_to_string('capcha.png', lang='rus')
            text = ''.join(
                [val for val in text if val.isalpha() or val.isnumeric()])
            capcha_text.send_keys(text)
            capcha_text.submit()

            try:
                self.driver.find_element_by_class_name('results-frame')

            except NoSuchElementException:

                self.driver.refresh()
            else:
                pass_capcha = False
        return True

    def close(self):
        self.driver.close()
        quit()


def main():
    obj = 'Кондратьев Сергей Сергеевич 03.11.1990'
    driver = Firefox()
    parser = FsspParser(driver)
    parser.parse(obj)


if __name__ == '__main__':
    main()
