# -*- coding: cp1252 -*-

from bs4 import BeautifulSoup
import urllib.request
import urllib.parse
from bs4.element import NavigableString
from bs4.element import Tag
import re
import userpaths
import os

awmf_url = "https://www.awmf.org/leitlinien/aktuelle-leitlinien.html"

def get_url(soup, type):
    document_list = soup.findAll('div', {'class': 'document-row no-indent'}).find("li", recursive=False)
    long_version_list = [x for x in document_list if type in x.text]
    if len(long_version_list) == 0: return ""
    for long_version in long_version_list:
        document_li = long_version.find
    pass

def get_document_url(html_text, type):
    match_list = re.findall('href="(.*?)pdf"', html_text)
    result = ""
    for match in match_list:
        start_index = html_text.index(match)
        substring = html_text[:start_index]
        sub_match_list = re.findall('<span class="document-name">(.*?)<\/span>', substring)
        if type in sub_match_list[-1]:
            result = match
            break

    return result

def parse_guideline_page(title, url):
    title = title.replace("/", "-")
    print("%s..." % title)
    target_path = "%s\\AWMF Downloads\\Leitlinien\\%s" % (userpaths.get_my_documents(), title)
    if os.path.exists(target_path) == False:
        os.makedirs(target_path, exist_ok=True)

    with urllib.request.urlopen("https://www.awmf.org/%s" % url) as url_request:
        guideline_html = url_request.read().decode()

    long_version_url = get_document_url(guideline_html, "Langfassung")
    if len(long_version_url) > 0:
        long_version_url = "https://www.awmf.org/%spdf" % long_version_url
        print(long_version_url)
        urllib.request.urlretrieve(long_version_url, "%s\\%s_Langfassung.pdf" % (target_path, title))
    else:
        pass

    coi_url = get_document_url(guideline_html, "Interessen")
    if len(coi_url) > 0:
        coi_url = "https://www.awmf.org/%spdf" % coi_url
        print(coi_url)
        urllib.request.urlretrieve(coi_url, "%s\\%s_CoI.pdf" % (target_path, title))
    else:
        pass

with urllib.request.urlopen(awmf_url) as url_request:
    awmf_html = url_request.read().decode()

soup = BeautifulSoup(awmf_html)

guideline_div_list = soup.findAll('div', {'class': 'col-title'})
guideline_div_list = [x for x in guideline_div_list if not isinstance(x.contents[0], NavigableString)]
for guideline_div in guideline_div_list:
    guideline_title = guideline_div.contents[0].attrs['title']
    guideline_url = guideline_div.contents[0].attrs['href']
    parse_guideline_page(guideline_title, guideline_url)
    pass


