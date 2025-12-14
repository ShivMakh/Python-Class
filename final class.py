# -*- coding: utf-8 -*-
"""
@author: ShivMakh
"""

#standard imports
import pandas as pd
import numpy as np

#web scraping
import requests
from bs4 import BeautifulSoup
import urllib.parse

#visuals
import seaborn as sns
import matplotlib.pyplot as plt

#saving plots as pdf
from matplotlib.backends.backend_pdf import PdfPages

#for email
import os
import win32com.client as win32


def tsy_yld_curve(dataset, date):
    dataset.dataset.set_index('Date')
    df_1day =  dataset[dataset.index==date]
    df_1day = df_1day.T
    
    #set the visual up
    fig = plt.figure(figsize=(50,25))
    plt.plot(df_1day.index, df_1day[date])
    #plt.show()
    
    fig.savefig(fr'{output_folder}\Tsy Yld Curve for {date.replace("/",".")}.png')
    
    return fig


def create_pdf(figs, date):
    pdf_name = fr'{output_folder}\TSY yld {date.replace("/",".")}.pdf'
    
    with PdfPages(pdf_name) as pdf:
        for fig in figs:
            pdf.savefig(fig)
            plt.close()
            
    return pdf_name


def email(subject='',body_text='',files=[],html_tables=[],html_tables_noindex=[],
          send_list = '',send_or_display='display'):
    
    outlook = win32.Dispatch('outlook.application')
    mail=outlook.CreateItem(0)
    
    mail.To = send_list
    
    mail.Subject = subject
    
    for file in files:
        mail.Attachments.Add(file)
        
    mail.HtmlBody = body_text + "<br>"
    
    for table in html_tables:
        mail.HtmlBody += table.to_html(index=False) + "<br>"
        

    for table in html_tables_noindex:
        mail.HtmlBody += table.to_html(index=True) + "<br>"
        
    mail.Send() if send_or_display.lower()[0]=='s' else mail.Display()
    
    return f'mail sent to {send_list}' if send_or_display.lower()[0]=='s' else 'showing email'



if __name__ == "__main__":
    
    output_folder = fr'C:\Users\i5 PC\Documents\GitHub\Python-Class'
    
    run_date  = '12/1/2017'    
    
    #%% dollar rolls
    
    reinvestment_rate = 1.7/100
    fin_rate = 1.7/100
    face = 100 #in mills
    paydown = 1 #mills
    
    