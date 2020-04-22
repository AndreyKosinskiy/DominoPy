import requests
import base64

SITE_AGENT = "http://localhost/testAnon.nsf/saveFile?OpenAgent"
#upload Data
def main_old():
    file_name_list = ['testforUpload.xlsx']
    headers ={
        "Content-Type":"application/octet-stream",
        }
    files= dict()
    for file_name in file_name_list:
        with open(file_name, 'rb') as file:
            files[file_name] = file.read()
            #files[file_name] = base64.encodestring(files[file_name])
            files[file_name] = base64.b64encode(files[file_name])

    # Send Data
    upload_file = requests.post(SITE_AGENT,files=files,headers=headers)
    print(upload_file.text)





def upload_file_to_document(doc_identifier,upload_file_path):
    """
    Paraments:
        doc_identifier - doc id, doc ouid, doc_uid
        upload_file_path - path to file what you wont to upload
    """
    files= dict()
    file_name = upload_file_path.split("\\")[-1]
    headers ={
        "Content-Type":"application/octet-stream",
    }
    data = {
        "doc_identifier":doc_identifier,
    }
    with open(upload_file_path, 'rb') as file:
       files[file_name] = file.read()
       files[file_name] = base64.b64encode(files[file_name])
    upload_file = requests.post(SITE_AGENT,files=files,headers=headers,data=data)
    print(upload_file.text)

upload_file_to_document("34212452",'testforUpload.xlsx')


