import xlrd
import xlsxwriter
from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import Select


class SudRfParser(object):
    def __init__(self, driver):
        self.driver = driver

    def parse_person_data(self, initials):
        person_data = {'initials': initials}
        self.driver.get('https://sudrf.ru/index.php?id=300#sp')
        self._input_man_data(initials)
        self.driver.implicitly_wait(10)
        header = self.driver.find_elements_by_xpath(
            '//*[@id="resultTable"]/table/thead/tr/td')

        if not header:
            return person_data

        person_data['header'] = header

        data = self.driver.find_elements_by_xpath(
            '//*[@id="resultTable"]/table/tbody/tr/td')
        person_data['data'] = data

        return person_data

    def _input_man_data(self, initials):
        select = self.driver.find_elements_by_xpath("//*[@id='court_subj']")[
            -1]
        select = Select(select)
        select.select_by_visible_text('Город Москва')

        input_person = self.driver.find_element_by_id('f_name')
        input_person.click()
        input_person.clear()
        input_person.send_keys(initials)
        input_person.submit()

    def close(self):
        self.driver.close()
        quit()


class XmlsWriter(object):
    def __init__(self):
        self.workbook = xlsxwriter.Workbook('sudrf.xlsx')

    def write_xlsx(self, data):
        judgment = data.get('data', None)
        if not judgment:
            return print(f'На {data["initials"]} нет открытых дел')

        worksheet = self.workbook.add_worksheet()
        row, col = 0, 0
        for item in data['header']:
            worksheet.write(row, col, item.text)
            col += 1

        for item in data['data']:
            if 'суд' in item.text:
                row += 1
                col = 0
            worksheet.write(row, col, item.text)
            col += 1
        return

    def close_xlsx(self):
        self.workbook.close()
        print('Все записанно')


def main():
    options = Options()
    options.headless = True
    driver = Firefox()
    get_xlsx = XmlsWriter()
    parser = SudRfParser(driver)
    rb = xlrd.open_workbook('mans_name.xlsx')
    sheet = rb.sheet_by_index(0)
    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        person = ' '.join(row)
        data = parser.parse_person_data(person)
        get_xlsx.write_xlsx(data)
    get_xlsx.close_xlsx()
    parser.close()


if __name__ == '__main__':
    main()
