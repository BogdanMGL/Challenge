"""Template robot with Python."""
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.HTTP import HTTP
from RPA.FileSystem import FileSystem


import os


url = "https://itdashboard.gov"
browser = Selenium()
excel = Files()
file = FileSystem()

view_agencies = 'css:#node-23 a'
table_agencies = 'css:#agency-tiles-container'
agencia_name = 'css:#agency-tiles-widget span.h4.w200'
agencia_spending = 'css:#agency-tiles-widget span.h1.w900'
first_agency = 'css:#agency-tiles-widget a'
table_individual_investment = "css:div.dataTables_scroll"
select_number_of_row = 'css:select[name ="investments-table-object_length"]'
row_table_individual_investment = "css:#investments-table-object > tbody > tr"
download_link = "css:#business-case-pdf > a"


def open_website():
    browser.set_download_directory(os.path.abspath(os.curdir) + "/output")
    browser.open_available_browser(url)


def get_agencies():

    browser.click_element_if_visible(view_agencies)
    browser.wait_until_element_is_visible(table_agencies)
    names_list = browser.get_webelements(agencia_name)
    list_spending = browser.get_webelements(agencia_spending)
    list_agencies = []
    for name, spending in zip(names_list, list_spending):
        list_agencies.append({"name": name.text, "spending": spending.text})
    return list_agencies


def create_excel_table():
    excel.create_workbook("output/challenge.xlsx")
    excel.rename_worksheet("Sheet", "Agencies")


def agencies_to_the_table(agencies):
    excel.set_worksheet_value("1", "1", "Agency name")
    excel.set_worksheet_value("1", "2", "Amount of expenses")
    excel.append_rows_to_worksheet(agencies)
    excel.save_workbook()


def get_individual_investments():
    browser.click_element_if_visible(first_agency)
    browser.wait_until_element_is_visible(
        table_individual_investment, timeout="50")
    browser.select_from_list_by_value(select_number_of_row, "-1")
    browser.wait_until_page_contains_element(
        row_table_individual_investment, limit=209, timeout="20")
    link_list = []
    individual_investments_list = []
    row_table = browser.get_webelements(row_table_individual_investment)
    for index in range(len(row_table)):
        td = row_table[index].find_elements_by_tag_name('td')
        individual_investments_list.append({
            "UII": td[0].text,
            "bureau": td[1].text,
            "investment_title": td[2].text,
            "total": td[3].text,
            "type": td[4].text,
            "CIO_rating": td[5].text,
            "of_projects": td[6].text})
        try:
            a = td[0].find_element_by_tag_name('a')
            link = a.get_attribute('href')
            link_list.append({"file_name": td[0].text, "link": link})
        except:
            continue
    return {"table": individual_investments_list, "link_list": link_list}


def individual_investments_to_the_table(individual_investments):
    worsheet_name = "Individual Investments"
    excel.create_worksheet(worsheet_name)
    excel.set_worksheet_value("1", "1",  "UII")
    excel.set_worksheet_value("1", "2",  "Bureau")
    excel.set_worksheet_value("1", "3",  "Investment Title")
    excel.set_worksheet_value("1", "4",  "Total FY2021 Spending ($M)")
    excel.set_worksheet_value("1", "5",  "Type")
    excel.set_worksheet_value("1", "6",  "CIO Rating")
    excel.set_worksheet_value("1", "7",  "# of Projects")
    excel.append_rows_to_worksheet(
        individual_investments, "Individual Investments")
    excel.save_workbook()


def downloads_file(link_list):
    path_file = os.path.abspath(os.curdir) + "/output"
    list_file = os.listdir(path_file)
    for name in list_file:
        if name.endswith(".pdf"):
            os.remove(path_file + "/" + name)
    for index in range(len(link_list)):
        link = link_list[index]["link"]
        browser.go_to(link)
        browser.wait_until_element_is_visible(
            download_link, timeout="15")
        browser.click_link(download_link)
        path = os.path.abspath(os.curdir) + "/output/" + \
            link_list[index]["file_name"] + ".pdf"
        try:
            file.wait_until_created(path, timeout="50")
        except:
            index = index-1
            continue


if __name__ == "__main__":
    try:
        open_website()
        agencies = get_agencies()
        create_excel_table()
        agencies_to_the_table(agencies)
        result = get_individual_investments()
        individual_investments_to_the_table(result["table"])
        downloads_file(result["link_list"])
    finally:
        browser.close_browser()