# -*- coding: utf-8 -*-
"""
Created on Fri Feb 23 13:31:41 2024

@author: pradeep_phule
"""

from reportlab.lib.pagesizes import letter
import datetime

from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, PageBreak, Spacer, Table, TableStyle
from reportlab.lib import colors
 
from string import digits
import pandas as pd
import numpy as np
import os as os
import sys


def convert_to_int(x):
    try:
        return int(''.join(filter(str.isdigit, str(x))))

    except:
        return np.nan

#filename_drug_trials = "D:\\dotNet\\integration\\merged_drug_trial.xlsx"
filename_drug_trials = sys.argv[3]+"_merged.xlsx"
druginfo_table = pd.read_excel("Drugs.xlsx",sheet_name = "Sheet1")
trials_table = pd.read_excel("Trials.xlsx",sheet_name = "Sheet1")
patient2_score = pd.read_csv(sys.argv[3],sep="\t")
patient2_score = patient2_score.dropna(subset=['P. Notation'])
# drug position formatting
druginfo_table['position'] = druginfo_table['P. Notation'].apply(lambda x: int(''.join(filter(str.isdigit,x))))
trials_table['position'] = trials_table['P. Notation'].apply(lambda x: int(''.join(filter(str.isdigit,x))))
patient2_score['position'] = patient2_score['P. Notation'].apply(convert_to_int)
# merging
patient2_merge_drug = pd.merge(patient2_score, druginfo_table, on=['Gene','Gene ID','position'],how='left')
patient2_merge_trials = pd.merge(patient2_score, trials_table, on=['Gene','Gene ID','position'],how='left')
patient2_merge_drug_trials = pd.merge(patient2_merge_drug, trials_table, on=['Gene','Gene ID','position'],how='left')
# export file
export_file = filename_drug_trials
patient2_merge_drug_trials.to_excel(export_file,"patient_drug_trial")





# %%

merge = pd.read_excel(filename_drug_trials,sheet_name="patient_drug_trial")
table_merge_drugs = [["Gene Name","P. Notation","Drug","Indication"]]
table_merge_trials = [["Gene Name","P. Notation","Trial Name","Trial ID"]]
for i in merge.index:
    if pd.isna(merge.loc[i,'Drugs']):
        continue
    
    table_merge_drugs.append([merge.loc[i,'Gene'],merge.loc[i,'type'],merge.loc[i,'Drugs'],merge.loc[i,'Indication']])
    if pd.isna(merge.loc[i,'Trial Name']):
        continue
    table_merge_trials.append([merge.loc[i,'Gene'],merge.loc[i,'type'],merge.loc[i,'Trial Name'],merge.loc[i,'Trial ID']])
    



# %%

patient_info = pd.read_excel(sys.argv[1]) ## Pass as argument
#patient_info = pd.read_excel("D:/python_practice/Patient_Details 1.xlsx") ## Pass as argument
second_col = patient_info.loc[0]
patient_info.columns = second_col


file_path = sys.argv[2] ## Pass as argument 'inputDirectory'
#file_path = 'D:/dotNet/patient2.vcf'   ## Pass as argument 'inputDirectory'
file_name = os.path.basename(file_path)
patine_data = file_name[0:-4]
#print(patine_data)
order_number = patient_info['Order Number']
order_index =0
index = 0
for order in order_number:
    if order.upper() == patine_data.upper():
        order_index = index
    index+=1

print(order_index)

# %%
order_diagnosis = patient_info['Order Number']+"_"+patient_info['Initial Diagnosis'].replace(to_replace="\s+",value="_",regex=True)
patient_first_name = "Patient Name: " + patient_info['Patient First Name']+" " + + patient_info['Patient Last Name']
patient_last_name = "Patient Last Name: " + patient_info['Patient Last Name']
gender = "Gender/Sex: " + patient_info['Gender']
date_of_birth = patient_info['DOB']
medical_record_number = "Medical Record Number: "+ patient_info['Medical Record Number']
sample_type = "Sample Type: " + patient_info['Sample Type']
sample_source = "Sample Source: "+ patient_info['Sample Source']
physician = "Physician: " + patient_info['Physician']
institute = "Institute: " + patient_info['Institution']
sample_collection_date = patient_info['Collection Date']
sample_processing_date = patient_info['Processing Date']
initial_diagnosis = "Initial Diagnosis: " + patient_info['Initial Diagnosis']
final_diagnosis = "Final Diagnosis: "+ patient_info['Final Diagnosis']

#print()


# %%
score_all=[]
table_data = [['Gene Name','Ensemble ID','P Notation','Variant Classification','Score']]
row_num =0
with open(sys.argv[3],'r') as score_file:     ## Pass as argument
#with open('D:\\dotNet\\integration\\patient2_scores.txt','r') as score_file:    
    contents = score_file.readlines()
    for content in contents:
        row_num+=1
        if row_num >1:
            data=content.split("\t")
            if data[4].strip() != '0.0':
                score_all.append(content.strip())
                table_data.append([data[0].strip(),data[1].strip(),data[2].strip(),data[3].strip(),data[4].strip()])


from reportlab.graphics.shapes import Drawing, Line
header_image_path = "semicol_image.png"

def header_content(canvas,doc):
    canvas.saveState()
    header_pos = letter[1] - 50
    canvas.drawImage(header_image_path, 50, header_pos - 50, width=300, height=80)
    canvas.drawString(50, 680, order_diagnosis[order_index])
    date = datetime.datetime.now()
    date_for_pdf = "Date: "+ str(date.month) +"/"+str(date.day)+"/"+str(date.year)
    canvas.drawString(480, 680, str(date_for_pdf))
    canvas.setFont("Helvetica", 9)
    page_number_text = "Page %d" % (canvas.getPageNumber())   
    canvas.drawRightString(570, 30, page_number_text) 
    timestamp_text = "Report generated on: %s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    canvas.drawString(50, 50, timestamp_text)
    canvas.drawString(50, 30, get_footer_text_2())
    canvas.setStrokeColor(colors.black)
    canvas.rect(20, 20, 560, 750)

def get_footer_text_1():
    return ("Name: Dr. Evelyn Stone, MD Designation: Chief Pathologist")

def get_footer_text_2():
    return ("Name: Dr. Evelyn Stone, MD Designation: Chief Pathologist Address: Stone Pathology Clinic Cityville, State ZIP USA")

def add_border(canvas,doc):
    canvas.saveState()
    canvas.setStrokeColor(colors.black)
    canvas.rect(0, 0, doc.width, doc.height0)
    canvas.restoreState()

def sections(canvas,doc):
    canvas.saveState()
    canvas.setFillColorRGB(0.85, 0.95, 1)
    canvas.rect(20, 450, 560, 30, fill=1)
    
    canvas.setFont( "Helvetica-Bold", 14)
    canvas.setFillColorRGB(0, 0, 0) # Black color 
    canvas.drawString(40, 460,"Summary Table")
    canvas.restoreState()

dob_patient = "Date of Birth: " + str(date_of_birth[order_index]).split(" ")[0]
collection_date = "Sample Collection Date: " + str(sample_collection_date[1]).split(" ")[0]
process_date = "Sample Processing Date: " + str(sample_processing_date[1]).split(" ")[0]

text_data_before_table = []
text_data_before_table.append([patient_first_name[order_index],gender[order_index]])
text_data_before_table.append([dob_patient,medical_record_number[1]])
text_data_before_table.append([sample_type[1],sample_source[1]])
text_data_before_table.append([physician[1],institute[1]])
text_data_before_table.append([collection_date,process_date])
text_data_before_table.append([initial_diagnosis[1],final_diagnosis[1]])                        


 
# Generate a PDF with header image and text

    
#pdf = SimpleDocTemplate("aa.pdf", pagesize=letter,topMargin=150)   ## Pass as argument
pdf = SimpleDocTemplate(sys.argv[4], pagesize=letter,topMargin=150)   ## Pass as argument

table_text = Table(text_data_before_table,colWidths=[80,300])

table_style_text = TableStyle([
   
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
    ('FONTSIZE', (0, 0), (-1, 0), 8),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
    ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
    ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
    ('FONTSIZE', (0, 1), (-1, -1), 8),
    ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
])

table_text.setStyle(table_style_text)


table_style = TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, 0), 8),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
    ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
    ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
    ('FONTSIZE', (0, 1), (-1, -1), 8),
    ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
    ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
    ('BOX', (0, 0), (-1, -1), 1,colors.black),
])

pdf_table_text = []



table = Table(table_data,colWidths=[80,80,80,120,80])
table.setStyle(table_style)
table.repeatRows=1
pdf_table = []

d = Drawing(10, 10)
d.add(Line(1, 160, 350, 160))

para1 = Paragraph("Patient Details")
pdf_table.append(para1)

#d.add(Line(1, 5, 350, 5))
#d.add(Paragraph("Summary Table"))
para = Paragraph("Summary Table")

pdf_table.append(table_text)

pdf_table.append(para)
pdf_table.append(d)

pdf_table.append(Spacer(1,20))

pdf_table.append(table)

pdf_table.append(Spacer(1,30))


para2 = Paragraph("Summary Table")
table_drugs = Table(table_merge_drugs,colWidths=[80,80,80,80])
table_drugs.setStyle(table_style)
table_drugs.repeatRows=1

table_trials = Table(table_merge_trials,colWidths=[80,80,80,80])
table_trials.setStyle(table_style)
table_trials.repeatRows=1

para2 = Paragraph("Drugs Table")
pdf_table.append(para2)
pdf_table.append(table_drugs)


pdf_table.append(Spacer(1,30))


para3 = Paragraph("Trial Table")
pdf_table.append(para3)


pdf_table.append(table_trials)



pdf_table.append(Spacer(1,30))
para4 = Paragraph("General Information")
pdf_table.append(para4)


custome_style = ParagraphStyle('MyStyle', fontName='Helvetica', fontSize=8)

para5 = Paragraph("Disclaimer: The information provided in this report is for general informational purposes only. It is not intended as medical " 
                  +"advice, diagnosis, or treatment. Always seek the advice of your physician or other qualified health provider with any " 
                  +"questions you may have regarding a medical condition. Never disregard professional medical advice or delay in seeking "
                  +"it because of something you have read on this website. Reliance on any information provided herein is solely at your " 
                  +"own risk. Methodology: Data analysis was conducted using thematic analysis, whereby patterns, themes, and categories "
                  +"were identified within the interview transcripts. Additionally, quantitative data were collected through surveys "
                  +"distributed to a larger sample size, allowing for broader insights and statistical analysis. "
                  +"The survey instrument was developed based on the findings from the qualitative phase and relevant literature.",style=custome_style)

pdf_table.append(para5)






pdf.build(pdf_table,onFirstPage=header_content,onLaterPages=header_content)


