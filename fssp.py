import io

import pytesseract
import xlsxwriter
from PIL import Image
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import Firefox
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class FsspParser(object):
    """Класс парсинга данных"""
    def __init__(self, driver):
        self.driver = driver

    def parse_person_data(self, initials):
        '''
        Принимает на вход данные в виде имени, фамилии, отчества, даты
        рождения. Возвращает данные о судебных предписаниях по указанным данным
        '''
        initials_keys = ['surname', 'name', 'patronymic', 'birth']
        initials_value = initials.split()
        person_data = dict(zip(initials_keys, initials_value))
        self._input_man_data(**person_data)
        header = self.driver.find_elements_by_tag_name('th')
        person_data['header'] = header
        data = self.driver.find_elements_by_tag_name('td')
        person_data['data'] = data[1:]

        return person_data

    def _input_man_data(self, **initials):
        """Вводит данные в поисковую систему"""

        self.driver.get('https://fssp.gov.ru/')

        WebDriverWait(self.driver, 1000000).until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@type='button']"))).click()

        self.driver.find_element_by_link_text('Расширенный поиск').click()

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

        self.driver.implicitly_wait(5)
        try:
            self.driver.find_element_by_xpath(
                '//input[@id="captcha-popup-code"]')
        except NoSuchElementException:
            return True
        pass_capcha = self._pass_capcha()

        if pass_capcha:
            return True

    def _pass_capcha(self):
        """
        Прохождение капчи. Работает не всегда с первогораза, но 100%
        проходит капчу. В дальнейшем возможно включение режима обучения, для
        более правильной работы.
        """
        try_pass_capcha = True

        while try_pass_capcha:

            capcha_text_imput = self.driver.find_element_by_xpath(
                '//input[@id="captcha-popup-code"]')
            capcha_text_imput.click()
            capcha_text_imput.clear()
            # распазнование капчи
            # получение элемента содержащего картинку с текстом капчи
            image = self.driver.find_element_by_id(
                "capchaVisual").screenshot_as_png
            # получение потока байтов нужной области
            imageStream = io.BytesIO(image)
            # сохранение потока в картинку с помощью Pillow
            im = Image.open(imageStream)
            im.save('capcha.png')
            # распазнование текста на картинке
            text = pytesseract.image_to_string('capcha.png', lang='rus')
            # фильтрация лишниц символов в распознаном текесте
            text = ''.join(
                [val for val in text if val.isalpha() or val.isnumeric()])
            capcha_text_imput.send_keys(text)
            capcha_text_imput.submit()

            self.driver.implicitly_wait(10)

            try:
                self.driver.find_element_by_class_name('results-frame')

            except NoSuchElementException:
                self.driver.refresh()
                print('Не прошел капчу. Пробую еще.')
            else:
                try_pass_capcha = False

        return True

    def close(self):
        self.driver.close()
        quit()


class XmlsWriter(object):
    """Осуществляет запись спарсенных данных в файл dossier.xlsx"""
    def __init__(self):
        self.workbook = xlsxwriter.Workbook('dossier.xlsx')

    def write_xlsx(self, data):
        surname = data['surname'].upper()
        judgment = data.get('data', None)
        if not judgment:
            return print(
                f'На {data["surname"]} {data["name"]} нет открытых дел')

        worksheet = self.workbook.add_worksheet()
        row, col = 0, 0
        for item in data['header']:
            worksheet.write(row, col, item.text)
            col += 1

        for item in data['data']:
            _ = item.text
            if surname in _:
                row += 1
                col = 0
            worksheet.write(row, col, item.text)
            col += 1
        print(f'Данные по {data["surname"]} {data["name"]} записанны в'
              f' файл dossier.xlsx')
        return

    def close_xlsx(self):
        self.workbook.close()
        print('Все записанно')


def main():
    example = [
        'Кондратьев Сергей Сергеевич 03.11.1990',
        'Христолюбов Алексей Сергеевич 03.11.1990'
    ]
    options = Options()
    options.headless = True
    # executable_path = указать путь до geckodriver в вашей системе
    # driver = Firefox(options=options, executable_path=executable_path)
    driver = Firefox(options=options)
    parser = FsspParser(driver)
    get_xlsx = XmlsWriter()
    for person in example:
        data = parser.parse_person_data(person)
        get_xlsx.write_xlsx(data)
    get_xlsx.close_xlsx()
    parser.close()


if __name__ == '__main__':
    main()
