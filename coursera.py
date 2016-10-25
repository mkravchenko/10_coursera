import json
import random
import re
import requests
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
from openpyxl import Workbook


def get_courses_list(number_of_courses=20):
    course_urls_list = []
    site_map = requests.get("https://www.coursera.org/sitemap~www~courses.xml")
    root = ET.fromstring(site_map.text)
    list_urls = root.getchildren()
    random.choice(list_urls)
    for loc in list_urls:
        if number_of_courses:
            course_urls_list.append(loc[0].text)
            number_of_courses -= 1
    return course_urls_list


def get_courses_info(course_urls_list):
    curses_base_info_list = []
    for url in course_urls_list:
        print(url)
        response = requests.get(url)
        parsed_html = BeautifulSoup(response.text, 'html.parser')

        if not parsed_html.find_all(class_='rc-TopAlertBar info'):
            curses_base_info_list.append({"name": get_course_name(parsed_html),
                                          "start_data": get_course_start_time(parsed_html),
                                          "number_of_weeks": get_number_of_weeks(parsed_html),
                                          "language": get_language(parsed_html),
                                          "user_ratings": get_user_rating(parsed_html)})
    return curses_base_info_list


def get_course_name(parsed_html):
    course_name = parsed_html.find("div", class_="title display-3-text")
    if course_name is not None:
        return course_name.get_text().encode('utf-8')
    return None


def get_course_start_time(parsed_html):
    script = parsed_html.find("script", type="application/ld+json")
    if script is not None:
        json_data = json.loads(script.contents[0])
        return json_data["hasCourseInstance"][0]["startDate"]
    return None


def get_number_of_weeks(parsed_html):
    number_of_weeks = len(parsed_html.find_all("div", class_='week'))
    if not number_of_weeks:
        return None
    else:
        return number_of_weeks


def get_user_rating(parsed_html):
    value_to_return = parsed_html.find("div", class_="ratings-text bt3-visible-xs")
    if not value_to_return:
        return None
    return re.sub("[^0-9.]", "", value_to_return.get_text())


def get_language(parsed_html):
    data = parsed_html.find(
            class_="basic-info-table bt3-table bt3-table-striped bt3-table-bordered bt3-table-responsive").find_all(
            "tr")
    value_to_return = None
    for raw in data:
        title_name = raw.find("span", class_="td-title").get_text().lower()
        if title_name == "language":
            value_to_return = raw.find(class_="td-data").get_text()
    return value_to_return


def output_courses_info_to_xlsx(data_list, filepath):
    header = [u'Name', u'Languages', u'Start data', u'Number of weeks', u'User ratings(stars)']
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Coursera courses info"
    ws1.append(header)
    for data in data_list:
        ws1.append([data["name"], data["language"], data["start_data"], data["number_of_weeks"], data["user_ratings"]])
        wb.save(filename=filepath)
    print("Successfully added information to the: {}".format(filepath))


def get_path_to_xlsx_file():
    while True:
        path_to_file = input("\nEnter path to .xlsx file: ")
        if path_to_file:
            if ".xlsx" in path_to_file:
                return path_to_file
            else:
                print("No file  extension '.xlsx'!")
        else:
            print("Path to file can't be empty!")


if __name__ == '__main__':
    path_to_xlsx_file = get_path_to_xlsx_file()
    list_of_urls = get_courses_list()
    courses_info_list = get_courses_info(list_of_urls)
    output_courses_info_to_xlsx(courses_info_list, path_to_xlsx_file)
