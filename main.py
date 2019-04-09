import threading
import zipfile
from PyPDF2 import PdfFileMerger
import subprocess
from docx.shared import Cm
import comtypes.client
import sys
from shutil import copy2, copytree, rmtree, move
import pathlib
import requests
from io import BytesIO
from distutils.dir_util import copy_tree
import csv
from docx import Document
from docxcompose.composer import Composer
import os
from os import listdir
from os.path import isfile, join
wdFormatPDF = 17


def zipdir(path, ziph):
    length = len(path)

    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        folder = root[length:]  # path without "parent"
        for file in files:
            ziph.write(os.path.join(root, file), os.path.join(folder, file))
#
# def zipdir(path, ziph):
#     # ziph is zipfile handle
#     for root, dirs, files in os.walk(path):
#         for file in files:
#             ziph.write(os.path.join(root, file))
#


BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

mypath = os.path.join(BASE_DIR)
src_path = os.path.join(BASE_DIR, "src")

templates = os.path.join(BASE_DIR, src_path, "Templates")

onlyfiles = [f for f in listdir(templates) if isfile(join(templates, f))]

onlyfolder = [f for f in listdir(templates) if not isfile(join(templates, f))]


def update_toc(docx_file):
    word = comtypes.client.CreateObject("Word.Application")
    doc = word.Documents.Open(docx_file)
    toc_count = doc.TablesOfContents.Count
    if toc_count == 1:
        toc = doc.TablesOfContents(1)
        toc.Update
        print('TOC should have been updated.')
    else:
        print('TOC has not been updated for sure...')
    print(docx_file)
    doc.SaveAs(docx_file,FileFormat=16)
    doc.Close(SaveChanges=True)
    # word.Close()
    word.Quit()


def replace_word(path, cp_name, position, industry, logo):
    doc = Document(path)
    header = doc.sections[0].header

    for i, j in enumerate(header.paragraphs):

        if '[Company]' in j.text:
            header.paragraphs[i].text = j.text.replace(
                '[Company]', cp_name)
        if '[Position]' in j.text:
            header.paragraphs[i].text = j.text.replace(
                '[Position]', position)
    for g, k in enumerate(doc.paragraphs):
        if '[Company]'in k.text:
            doc.paragraphs[g].text = k.text.replace(
                '[Company]', cp_name)

        if'[Position]' in k.text:
            doc.paragraphs[g].text = k.text.replace(
                '[Position]', position)

        if'[Industry]' in k.text:
            doc.paragraphs[g].text = k.text.replace(
                '[Industry]', industry)

        if'[Company Image]' in k.text:
            doc.paragraphs[g].text = ""
            try:
                response = requests.get(logo)
            except:
                response = ""
            binary_img = BytesIO(response.content)
            pic = doc.paragraphs[g]
            run = pic.add_run()
            run.add_picture(binary_img)
    doc.save(path)
    return doc


def combine_word_documents(files, merged_name):

    merged_document = Document()

    for index, file in enumerate(files):
        sub_doc = Document(file)

        # Don't add a page break if you've reached the last file.
        if index < len(files) - 1:
            sub_doc.add_page_break()

        for element in sub_doc.element.body:
            merged_document.element.body.append(element)

    merged_document.save(merged_name)


def merged_docx(files, merged_name):
    print(files)
    master = Document()
    composer = Composer(master)
    for i in files:
        composer.append(Document(i), remove_property_fields=False)
    composer.save(os.path.join(src_path, merged_name))
    new_doc = Document(merged_name)
    for i in new_doc.sections:
        i.left_margin = Cm(1.5)
        i.right_margin = Cm(1.5)

    new_doc.save(merged_name)

    # update_toc(merged_name)
    return merged_name


def get_form(path):

    with open(os.path.join(src_path, path)) as f:
        b = csv.reader(f)
        return list(b)


def convert_pptx_to_pdf(src, dst):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    deck = powerpoint.Presentations.Open(src)
    deck.SaveAs(dst, 32)
    deck.Close()
    powerpoint.Quit()
    return True


def convert_to_pdf(src, dst):
    
    dst=dst.replace("docx","pdf" )
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    #word.DisplayAlerts = False
    doc = word.Documents.Open(src)
    doc.SaveAs(dst, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
    return True


# def convert_to_pdf(src, dst):
#     p = subprocess.Popen(["soffice", "--headless", "--convert-to", "pdf",
#                           src, "--outdir", dst])
#     (output, err) = p.communicate()
#     p_status = p.wait()


def merge_pdf(pdfs, dst):
    merger = PdfFileMerger()
    for pdf in pdfs:
        merger.append(open(pdf, 'rb'))

    with open(dst, 'wb') as fout:
        merger.write(fout)


if __name__ == '__main__':
    file_input = [os.path.join(templates, i) for i in onlyfiles]

    form = get_form("form.csv")

    for data in form[1:-1]:

        cp_name = data[0].strip()
        logo = data[1].strip()
        industry = data[3].strip()
        position = data[4].strip()
        questionsandanswers = data[-2].strip()
        interviewprocess = data[-1].strip()

        parent_dir = os.path.join(
            BASE_DIR, "src", "Output", cp_name, position)
        temp_dir = os.path.join(parent_dir, "Temp Files")

        pathlib.Path(temp_dir).mkdir(parents=True, exist_ok=True)

        pathlib.Path(os.path.join(parent_dir, "Study Guide")
                     ).mkdir(parents=True, exist_ok=True)

        pathlib.Path(os.path.join(parent_dir, "Course")
                     ).mkdir(parents=True, exist_ok=True)

        for i in file_input:
            copy2(i, temp_dir)

        new_dir = os.path.join(BASE_DIR, "src", "Output", cp_name, position,
                               "Temp Files")

        for i in onlyfolder[0:-1]:
            copy_tree(os.path.join(templates, i), os.path.join(temp_dir, i))
        word = onlyfolder[-1]
        _cp_detail = os.path.join(templates, word, "Company Details")
        _industry = os.path.join(templates, word, "Industry Details")
        _interview_process = os.path.join(templates, word, "Interview Process")
        _jd = os.path.join(templates, word, "Job Description")
        _qa = os.path.join(templates, word, "List of Questions and Answers")
        _img = os.path.join(templates, "Images", "Course")

        cp_detail_word = [f for f in listdir(
            _cp_detail) if not isfile(join(templates, f))]
        industry_word = [f for f in listdir(
            _industry) if not isfile(join(templates, f))]
        interview_process_word = [f for f in listdir(
            _interview_process) if not isfile(join(templates, f))]

        jd_word = [f for f in listdir(
            _jd) if not isfile(join(templates, f))]
        qa_word = [f for f in listdir(
            _qa) if not isfile(join(templates, f))]
        img = [f for f in listdir(
            _img) if not isfile(join(templates, f))]
        print("==========PREPERING TOCOPY====================")

        for i in cp_detail_word:
            if cp_name.lower() in i.lower():
                copy2(os.path.join(_cp_detail, i), temp_dir)
                print(f"{cp_name} copied detail")

        for i in industry_word:
            if industry.lower() in i.lower():
                copy2(os.path.join(_industry, i), temp_dir)
                print(f"{cp_name} copied industry")

        for i in interview_process_word:
            if interviewprocess.lower() in i.lower():
                copy2(os.path.join(_interview_process, i), temp_dir)
                print(f"{cp_name} copied interviewprocess")

        for i in jd_word:

            if f"{cp_name} {position}.docx".lower() == i.lower():
                copy2(os.path.join(_jd, i), temp_dir)
                print(f"{cp_name} copied jd_word")

        for i in qa_word:
            if f"{questionsandanswers}.docx".lower() == i.lower():
                copy2(os.path.join(_qa, i), temp_dir)
                print(f"{cp_name} copied qa_word")

        for i in img:
            if f"{cp_name} {position}.jpg".lower() == i.lower():
                copy2(os.path.join(_img, i), os.path.join(
                    temp_dir, "Images"))
                print(f"{cp_name} copied img")

        print("==========COPY DONE====================")

        now_only_files = [f for f in listdir(
            new_dir) if isfile(join(new_dir, f))]


        only_files_docx = [f for f in now_only_files if "docx" in f]

        only_files_docx = [f for f in only_files_docx if f[:2] != "~$"]

        study_file = [f for f in only_files_docx if "Study" in f]
        workbook_file = [f for f in only_files_docx if "Workbook" in f]
        workbook_file.sort()

        study_list = []
        workbook_list = []
        study_file.sort()
        # repalce method
        for i in only_files_docx:
            replace_word(os.path.join(new_dir, i),
                         cp_name, position, industry, logo)

        study_file.insert(1, f"{interviewprocess}.docx")
        study_file.insert(3, f"{industry}.docx")
        study_file.insert(5, f"{cp_name}.docx")
        study_file.insert(7, f"{cp_name} {position}.docx")
        study_file.insert(9, f"{questionsandanswers}.docx")
        merged_study = os.path.join(
            new_dir, f'Study Guide–{cp_name} {position} Interview preparation.docx')
        merged_workbook = os.path.join(
            new_dir, f'Workbook–{cp_name} {position} Interview preparation.docx')
        merged_study_pdf = os.path.join(
            new_dir, f'Study Guide–{cp_name} {position} Interview preparation.pdf')
        merged_workbook_pdf = os.path.join(
            new_dir, f'Workbook–{cp_name} {position} Interview preparation.pdf')

        for i in study_file:
            study_list.append(os.path.join(new_dir, i))

        pdf_study_list = [i.replace("docx", "pdf") for i in study_list]

        print(
            f"========== MERGING: Study Guide–{cp_name} {position} Interview preparation.pdf... ====================")
        for i in study_list:
            convert_to_pdf(i, os.path.join(new_dir, os.path.basename(i)))
        merge_pdf(pdf_study_list, merged_study_pdf)

        print(
            f"========== MERGING: Study Guide–{cp_name} {position} Interview preparation.docx... ====================")
        merged_docx(
            study_list, merged_study)

        print(
            f"========== MERGING: Done ====================")

        copy2(merged_study_pdf, os.path.join(
            parent_dir, "Course", f'Study Guide–{cp_name} {position} Interview preparation.pdf'))
        copy2(merged_study_pdf, os.path.join(
            parent_dir, "Study Guide", f'Study Guide–{cp_name} {position} Interview preparation.pdf'))

        workbook_file.insert(1, f"{interviewprocess}.docx")
        workbook_file.insert(3, f"{industry}.docx")
        workbook_file.insert(5, f"{cp_name}.docx")
        workbook_file.insert(7, f"{cp_name} {position}.docx")
        workbook_file.insert(9, f"{questionsandanswers}.docx")
        for i in workbook_file:
            workbook_list.append(os.path.join(new_dir, i))
        print(
            f"========== MERGING: Workbook–{cp_name} {position} Interview preparation.docx... ====================")
        merged_docx(
            workbook_list, merged_workbook)

        for i in workbook_list:
            convert_to_pdf(i, new_dir)

        pdf_workbook_list = [i.replace("docx", "pdf") for i in study_list]
        merge_pdf(pdf_workbook_list, merged_workbook_pdf)

        print(
            f"========== MERGING: Done ====================")

        copy2(merged_workbook_pdf, os.path.join(
            parent_dir, "Course", f'Workbook––{cp_name} {position} Interview preparation.pdf'))
        for i in onlyfiles:
            if "pdf" in i:
                if f"Interview preparation" in i:
                    continue
                else:
                    os.remove(os.path.join(new_dir, i))

        pptx_file = os.path.join(
            new_dir, "Course", "Slides - Coursetake Interview Preparation.pptx")
        pptx_to = os.path.join(
            new_dir, "Course", f"Slides – {cp_name} {position} Interview preparation.pptx")
        os.rename(pptx_file, pptx_to)

        convert_pptx_to_pdf(pptx_to, os.path.join(
            parent_dir, "Course", f"Slides – {cp_name} {position} Interview preparation.pptx"))

        src_course = os.path.join(new_dir, "Course")
        src_file_course = [f for f in listdir(
            src_course) if isfile(join(src_course, f))]
        for i in src_file_course:
            copy2(os.path.join(src_course, i),
                  os.path.join(parent_dir, "Course"))
        os.chdir(parent_dir)

        zipf = zipfile.ZipFile(os.path.join(
            parent_dir, f"Course – {cp_name} {position} Interview preparation.zip"), 'w', zipfile.ZIP_DEFLATED)
        zipdir(os.path.join(parent_dir, "Course"), zipf)
        zipf.close()
