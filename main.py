import ppt2video
import os
import shutil
import json
import docx2pdf

if __name__ == "__main__":
    json_file =open('ppt.json',encoding='utf-8')
    path_data = json.load(json_file)
    ppt_path = path_data["ppt_path"]
    mp4_path = path_data["mp4_path"]
    docx_path = path_data["docx_path"]
    pdf_path = path_data["pdf_path"]
    ppt2video.ppt2video_transform(ppt_path,mp4_path)
    docx2pdf.word_to_pdf(docx_path,pdf_path)