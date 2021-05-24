import sys
import json
import pandas as pd


def token():
    
    # Option1 -- Store the refresh token in a file
    # myfile = open("refresh.txt", 'r')
    # refresh_token = myfile.read()
    # myfile.close()
    

    #Option 2 - Store the refresh token in an environment variable
    refresh_token = os.environ['refresh_token']
    params = {
          'grant_type': "refresh_token", 
          'client_id': "*Client ID*",
          'refresh_token': refresh_token
         }

    response = requests.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data=params)
    # print(response.content)
    

    access_token = response.json()['access_token']
    new_refresh_token = response.json()['refresh_token']
    # print(access_token)
    

    # with open("refresh.txt", "w") as tokenfile:
    #         tokenfile.write(new_refresh_token)    
    #         tokenfile.close()
    # access_token = os.environ['access_token']
    return access_token
  
 def getChildren(access_token):
    #Folder ID where the files are uploaded
    FolderID="016HLRNJNR5Q4TD5IJPRE3DY6VAN3FBFJB" #/PFA
    URL = "https://graph.microsoft.com/v1.0/me/drive/items/"+FolderID+"/children"
    headers={'Authorization': "Bearer " + access_token}
    data = requests.get("https://graph.microsoft.com/v1.0/me/", headers=headers)
    print(data.content)
    r = requests.get(URL, headers=headers)
    j = json.loads(r.text)
    # print(j)
    fileDetails={}
    for files in j['value']:
        fileDetails[files['id']] = {
        "driveID" : files["parentReference"]["driveId"],
        "fileName" : files["name"],
        "parentPath":files["parentReference"]["path"],
        "id":files["id"],
        "url":files['webUrl']
        }

    # print(j)
    return fileDetails 
 
  
 def getPayload(access_token,fileDetails):
    baseUrl = "https://graph.microsoft.com/v1.0/me/"
    #url = baseUrl+fileDetails['parentPath']+"/"+
    url = baseUrl+"/drive/items/"+fileDetails['id']+"/workbook/worksheets('Sheet1')/usedRange"
    headers={'Authorization': "Bearer " + access_token}
    response = requests.get(url, headers=headers)
    # print(json.loads(response.content))
    a = json.loads(response.content)['text']
    df = pd.DataFrame.from_records(a[1:], index=None)
    df.columns=a[0]
    # df.dropna(inplace = True)
    json_str = df.to_json(orient='records')
    return json.loads(json_str),df
  
 def main(event, lambda_context):
    access_token = token()
    filesToUpload = getChildren(access_token)
    # Perform any other process that is needed
    

 if __name__ == "__main__":
    main()
  
  
