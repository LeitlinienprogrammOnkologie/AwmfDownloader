from docx import Document
import os
import fitz
import re
import nltk
from nltk.tokenize import sent_tokenize
from nltk.stem.snowball import SnowballStemmer
from nltk.corpus import stopwords

nltk.download('punkt')
nltk.download('stopwords')

stemmer = SnowballStemmer("german")
stop_words = set(stopwords.words("german"))

BASE_PATH = "C:\\Users\\User\\Documents\\AWMF Downloads\\Leitlinien"

pdf_file_list = []

def get_word_file_list(path):
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.endswith(".pdf") and not "CoI" in file:
                pdf_file_list.append("%s\\%s" % (root, file))

        for dir in dirs:
            get_word_file_list(dir)

get_word_file_list(BASE_PATH)

search_string_arr = ["Entscheidungsfindung", "Aufklärung", "partizipativ", "gemeinsam entscheiden", "shared decision making", "Gespräch"]
re_pattern = "(%s)" % "|".join(search_string_arr)
found_dict = {}

def get_sentences(text, match_text, start_index, end_index):
    text_before = text[:start_index].replace("\n", " ")
    text_after = text[end_index:].replace("\n", " ")

    sentences_before = sent_tokenize(text_before)
    sentences_after = sent_tokenize(text_after)

    if isinstance(sentences_before, str):
        sentence_before = ""
        sentence = sentences_before
    else:
        if len(sentences_before) > 1:
            sentence_before = sentences_before[-2]
            sentence = sentences_before[-1]
        else:
            sentence_before = ""
            sentence = ""

    sentence += " %s" % match_text
    if isinstance(sentences_after, str):
        sentence_after = ""
        sentence += " %s" % sentence_after
    else:
        if len(sentences_after) > 1:
            sentence_after = sentences_after[1]
            sentence += " %s" % sentences_after[0]
        else:
            sentence_after = ""

    return [sentence_before.replace("  ", " "), sentence.replace("  ", " "), sentence_after.replace("  ", " ")]

def analyze_doc(doc, title, fg):
    global found_dict
    toc = doc.get_toc()

    for i, page in enumerate(doc):
        page_no = i + 1
        chapter = "N/A"
        if len(toc) > 0:
            chapter = [x for x in toc if x[2] <= page_no][-1][1]
        text = page.get_text()
        matches = re.finditer(re_pattern, text, re.IGNORECASE)
        for match in matches:
            start_index = match.start()
            end_index = match.end()
            sentences = get_sentences(text, match.group(), start_index, end_index)

            if title not in found_dict:
                found_dict[title] = []
            found_dict[title].append([fg, [page_no, chapter, match.group(), sentences]])
        pass

index = 0
for word_file in pdf_file_list:
    try:
        doc = fitz.open(word_file)
        path_arr = word_file.split("\\")
        fg = path_arr[6].split("-")[1].strip()
        analyze_doc(doc, os.path.basename(word_file), fg)
    except:
        with open("C:\\Users\\User\\Documents\\AWMF Downloads\\analysis_errors.txt", "a+", encoding="utf-8") as f:
            f.write("word_file\n")

    index += 1
    print("%s/%s (%s)" % (index, len(pdf_file_list), 100*index/len(pdf_file_list)))
    #if index > 100:
    #    break


excel_out = "Fachgesellschaft|Registernummer|Leitlinie|Klassifizierung|Erkannter Begriff|Seite|Kapitel|Satz mit Begriff|Vorheriger Satz|Nachfolgender Satz\n"
for k,v in found_dict.items():
    guideline_arr = k.split("_")
    guideline_reg_no = guideline_arr[0]
    guideline_name = guideline_arr[1]
    guideline_class = guideline_arr[-1].split(".")[0]

    for match in v:
        fg = match[0]
        found_entity = match[1][2]
        excel_line = "%s|%s|%s|%s|%s|%s|%s|%s|%s|%s\n" % (fg, guideline_reg_no, guideline_name, guideline_class, match[1][2], match[1][0], match[1][1], match[1][3][1], match[1][3][0], match[1][3][2])
        excel_out += excel_line

with open("suchergebniss_awmf.csv", "w", encoding="utf-8") as f:
    f.write(excel_out)