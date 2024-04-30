import pandas as pd
import os.path
import re
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

#Take variables from user, if they are not valid values try again for 5 times 
def get_user_input(prompt, validator, error_message, max_attempts=5):
    for _ in range(max_attempts):
        try:
            user_input = input(prompt)
            if validator(user_input):
                return user_input
            else:
                raise ValueError
        except ValueError:
            print(error_message)

# New's site input
platform = get_user_input(
    "Haber Sitesi (örnek: posta.com.tr): ",
    lambda x: ".com" in x,
    "Geçerli bir haber sitesi giriniz."
)
# Start Date input 
start_date = get_user_input(
    "Başlangiç Tarihi (Format: YYYY-MM-DD): ",
    lambda x: re.match(r"\d{4}-\d{2}-\d{2}", x),
    "Geçerli bir başlangiç tarihi giriniz."
)
# End Date input
end_date = get_user_input(
    "Bitiş Tarihi (Format: YYYY-MM-DD): ",
    lambda x: re.match(r"\d{4}-\d{2}-\d{2}", x),
    "Geçerli bir bitiş tarihi giriniz.")
# Most repeated "n" word input
n = int(get_user_input(
    "En çok tekrar eden kaç kelime istediğinizi giriniz: ",
    lambda x: x.isdigit(),
    "Geçerli bir sayi giriniz."
))
# URL limit input
row_limit = int(get_user_input(
    "Hesaplanacak URL sinirlama sayisi: ",
    lambda x: x.isdigit(),
    "Geçerli bir sayi giriniz."
))
if row_limit == 0:
    row_limit = 25000


def gsc_auth(scopes):
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', scopes)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('searchconsole', 'v1', credentials=creds)

    return service

scopes = ['https://www.googleapis.com/auth/webmasters']
service = gsc_auth(scopes)

request={
  "startDate": start_date,
  "endDate": end_date,
  "dimensions": ["page"],
  "type": "discover",
  "rowLimit": row_limit
}

# Write to an Excel file  
gsc_search_analytics = service.searchanalytics().query(siteUrl=f'sc-domain:{platform}', body=request).execute()

report = pd.DataFrame(data=gsc_search_analytics['rows'])
report.to_excel('Data.xlsx', index=False)

def calculate_word_frequencies(n) : 

    data=pd.read_excel("Data.xlsx")

    dict_address={}
    #Split the url "/" sign and "-" sign and process with the last part
    for address in data["keys"]:
        address=str(address)
        address=address.split("/")[-1].split("-")
                
        for word in address:
            #Ignore short words and numbers 
            if len(word)>=4 and not word.isdigit():

                #add word to the dict as value=1,if is already in dict add 1
                if word in dict_address:
                    dict_address[word]+=1
                else:
                    dict_address[word]=1
            
                            
    # Sort and get top 'n' words  
    sorted_list=sorted(dict_address.items(), key=lambda t:t[1], reverse=True)[:n]
    return sorted_list
sorted_list=calculate_word_frequencies(n)
df_address=pd.DataFrame(data=sorted_list,columns=["Kelimeler","Tekrar Sayilari"])

# Write to an Excel file  
writing_path="Word_Frequencies.xlsx"
with pd.ExcelWriter(writing_path) as writer:
    df_address.to_excel(writer, index=False,sheet_name='URL Count Page')

