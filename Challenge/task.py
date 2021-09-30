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


def open_website():
    browser.set_download_directory(os.path.abspath(os.curdir) + "/output")
    browser.open_available_browser(url)


def get_agencies():
    browser.click_element_if_visible(
        "xpath://html/body/main/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/a")
    browser.wait_until_element_is_visible(
        "xpath://html/body/main/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div[1]/div/div[2]/div/div/div")
    list_name = browser.get_webelements(
        "xpath://html/body/main/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div[1]/div/div[2]/div/div/div/div/div/div/div/div/div/a/span[1]")
    list_spending = browser.get_webelements(
        "xpath://html/body/main/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div[1]/div/div[2]/div/div/div/div/div/div/div/div/div/a/span[2]")
    list_agencies = []
    for i in range(len(list_name)):
        list_agencies.append(
            {"name": list_name[i].text, "spending": list_spending[i].text})
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
    browser.click_element_if_visible(
        "xpath://html/body/main/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div[1]/div/div[2]/div/div/div/div[1]/div[2]/div/div/div/a")
    browser.wait_until_element_is_visible(
        "xpath://html/body/main/div/div/div/div[4]/div/div/div/div[2]/div/div[1]/div/div/div/div/div[3]/div[2]/table", timeout="15")
    browser.select_from_list_by_value(
        "xpath://html/body/main/div/div/div/div[4]/div/div/div/div[2]/div/div[1]/div/div/div/div/div[2]/div[2]/div/label/select", "-1")
    browser.wait_until_page_contains_element(
        "xpath://html/body/main/div/div/div/div[4]/div/div/div/div[2]/div/div[1]/div/div/div/div/div[3]/div[2]/table/tbody/tr", limit=158, timeout="20")
    link_list = []
    individual_investments_list = []
    row_table = browser.get_webelements(
        "xpath://html/body/main/div/div/div/div[4]/div/div/div/div[2]/div/div[1]/div/div/div/div/div[3]/div[2]/table/tbody/tr")
    for i in range(len(row_table)):
        td = row_table[i].find_elements_by_tag_name('td')
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
    excel.create_worksheet("Individual Investments")
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
    for i in range(len(link_list)):
        link = link_list[i]["link"]
        browser.go_to(link)
        browser.wait_until_element_is_visible(
            "xpath://html/body/main/div/div/div/div[1]/div/div/div/div/div[1]/div/div/div/div/div[6]/a", timeout="15")
        browser.click_link(
            "xpath://html/body/main/div/div/div/div[1]/div/div/div/div/div[1]/div/div/div/div/div[6]/a")
        path = os.path.abspath(os.curdir) + "/output/" + \
            link_list[i]["file_name"] + ".pdf"
        try:
            file.wait_until_created(path, timeout="50")
        except:
            i = i-1
            continue


def main():
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


if __name__ == "__main__":
    main()
