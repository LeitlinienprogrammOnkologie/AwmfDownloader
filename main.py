# -*- coding: cp1252 -*-
import win32com.client
from bs4 import BeautifulSoup
import urllib.request
import urllib.parse
from bs4.element import NavigableString
from bs4.element import Tag
import re
import userpaths
import os
from lxml import etree
import glob
import pywin32_system32

word = win32com.client.Dispatch("Word.Application")
word.visible = 0
word.DisplayAlerts = False

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

def sanitize_path(path):
    keepcharacters = (' ', '.', '_', "\\", "(", ")", "-")
    return "".join(c for c in path if c.isalnum() or c in keepcharacters).rstrip()

def try_download(url, target_path):
    if os.path.exists(target_path):
        print("EXISTS. Skpping.")
        return True

    try:
        urllib.request.urlretrieve(url, target_path)
        return True
    except:
        with open("%s\\AWMF Downloads\\download_error.txt" % userpaths.get_my_documents(), "a+", encoding="utf-8") as f:
            f.write("%s\n" % url)
        return False

def parse_guideline_page(fg, title, url, reg_nr, classification):
    print("%s..." % title)
    target_path = "%s\\AWMF Downloads\\Leitlinien\\%s\\%s_%s_%s" % (userpaths.get_my_documents(), sanitize_path(fg), reg_nr, sanitize_path(title), classification)
    if os.path.exists(target_path) == False:
        os.makedirs(target_path, exist_ok=True)

    try:
        with urllib.request.urlopen("https://www.awmf.org/%s" % url) as url_request:
            guideline_html = url_request.read().decode()
    except Exception as e:
        with open("%s\\AWMF Downloads\\file_not_found.txt" % userpaths.get_my_documents(), "a+", encoding="utf-8") as f:
            f.write("%s\n" % url)
        return

    long_version_url = get_document_url(guideline_html, "Langfassung")
    if len(long_version_url) > 0:
        long_version_url = "https://www.awmf.org/%spdf" % long_version_url
        print(long_version_url)
        if try_download(long_version_url, "%s\\%s_%s_Langfassung_%s.pdf" % (target_path, reg_nr, sanitize_path(title), classification)):
            if not os.path.exists("%s\\%s_%s_Langfassung_%s.docx" % (target_path, reg_nr, sanitize_path(title), classification)):
                word.DisplayAlerts = False
                try:
                    wb = word.Documents.Open("%s\\%s_%s_Langfassung_%s.pdf" % (target_path, reg_nr, sanitize_path(title), classification), False, False, False)
                except:
                    wb = None

                if wb is not None:
                    try:
                        wb.SaveAs2("%s\\%s_%s_Langfassung_%s.docx" % (target_path, reg_nr, sanitize_path(title), classification), FileFormat=16)
                    except:
                        with open("%s\\AWMF Downloads\\conversion_fails.txt" % userpaths.get_my_documents(), "a+",
                                  encoding="utf-8") as f:
                            f.write("%s\n" % ("%s\\%s_%s_Langfassung_%s.pdf" % (
                            target_path, reg_nr, sanitize_path(title), classification)))
                    finally:
                        wb.Close()
                else:
                    with open("%s\\AWMF Downloads\\conversion_fails.txt" % userpaths.get_my_documents(), "a+",
                              encoding="utf-8") as f:
                        f.write("%s\n" % ("%s\\%s_%s_Langfassung_%s.pdf" % (target_path, reg_nr, sanitize_path(title), classification)))
    else:
        pass

    coi_url = get_document_url(guideline_html, "Interessen")
    if len(coi_url) > 0:
        coi_url = "https://www.awmf.org/%spdf" % coi_url
        print(coi_url)
        try_download(coi_url, "%s\\%s_CoI.pdf" % (target_path, sanitize_path(title)))
    else:
        pass

with urllib.request.urlopen(awmf_url) as url_request:
    awmf_html = url_request.read().decode()

tree = etree.HTML(awmf_html)
#fg = tree.xpath('//ul[starts-with(@id, "alphabet-list") and contains(@id, "-content")]')
fg_parent = tree.xpath('//*[@id="tab-regnumbers-content"]/ul/li')

for element in fg_parent:
    fg = element.xpath('./a')[0].text
    guidelines = element.xpath('./ul/li')
    for guideline in guidelines[1:]:
        reg_nr = guideline.xpath('./div[@class="col-reg col1"]')[0].text
        classification = guideline.xpath('./div[@class="col-classification"]/span')[0].text
        title = guideline.xpath('./div[@class="col-title"]/a')[0].attrib['title']
        url = guideline.xpath('./div[@class="col-title"]/a')[0].attrib['href']
        parse_guideline_page(fg, title, url, reg_nr, classification)

word.Quit()