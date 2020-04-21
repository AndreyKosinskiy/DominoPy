import os
import re
import requests
import openpyxl
import logging

 #init logging
logging.basicConfig(format='%(asctime)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S',level=logging.INFO)
def action_with_data_in_file(filepath):
    """get flat data table  """
    table = [] 
    book = openpyxl.load_workbook(filepath)
    sheet = book.active
    
    for c1 in sheet[sheet.dimensions]:
        row = ''
        for c2 in c1:
            row =row+str(c2.value)+';'
        table.append(row)
    os.remove(filepath)
    logging.info(u"Delete file:: path: "+filepath)
    logging.info(u"Genereted flat table.")
    return table

def main():
    #get entity from view
    entityes = requests.get("http://localhost/testAnon.nsf/testView?ReadViewEntries&ExpandView")
    logging.info(u"Get:: http://localhost/testAnon.nsf/testView?ReadViewEntries&ExpandView")
    # get urls, using rexexp
    urls = re.findall('<text>(.*)<\/text>', entityes.text)

    for url_for_download_file in urls:
        #iterate through urls list
        download_file = requests.get(url_for_download_file)
        #get file name from url
        path_for_temp_save = url_for_download_file.split('/')[-1]
        #get document ouid name from url
        doc_ouid = url_for_download_file.split('/')[-3]
        with open(path_for_temp_save, 'wb') as file:
            file.write(download_file.content)
            logging.info(u"Download file:: path: "+path_for_temp_save+" ouid source doc: "+doc_ouid)
        #get flat interpritate of data from saved file
        table = action_with_data_in_file(path_for_temp_save)
        #build dict with key:value
        new_data = {
            'mark':'Marked',
            'ouid':doc_ouid,
            'flatcontent': table
        }
        report = requests.post("http://localhost/testAnon.nsf/ProcessedTestForm?CreateDocument",data = new_data)
        logging.info(u"Post:: http://localhost/testAnon.nsf/ProcessedTestForm?CreateDocument, ouid source doc: "+doc_ouid)

if __name__ == "__main__": 
    main()
