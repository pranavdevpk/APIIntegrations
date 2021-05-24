import pandas as pd

# To see how to generate token, refer to the OneDrive integration file
def outlookEmail(**kwargs):
    print("here")
    url ="https://graph.microsoft.com/v1.0/me/sendMail"
    token=kwargs['token']
  
    receiver_email_id = kwargs["emailAddr"]
    b64_string =convertforemail(kwargs["data"])
    reqbody = {"message": {
        "subject":"Subject" ,
        "body": {
         "contentType": "Text",
        "content": body
         },
        "Attachments": [
              {
                "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
                "Name": "menu.csv",
                 "ContentBytes": b64_string
             }
             ],
            "toRecipients": [
        {
        "emailAddress": {
            "address": receiver_email_id.strip()
         }}]}}



    headers={'Authorization': "Bearer " + token, "Content-Type": "application/json"}
    response = requests.post(url, headers=headers,data=json.dumps(reqbody))
    print(response.content)
    

def convertforemail(jsondata):
    data =pd.DataFrame(jsondata)
    ss = io.StringIO()
    data.to_csv(ss,index=False)
    attachment = MIMEText(ss.getvalue()).as_bytes()
    #kwargs["errorRows"].to_csv(ss, index=False)
    b64_bytes = base64.urlsafe_b64encode(attachment)
    b64_string = b64_bytes.decode()
    return b64_string
