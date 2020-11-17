####################################################################################
# G Suit/Google Workspace domain konverter
#
# Ez a script módosítja az összes felhasználó és csoport levelezési címében
# a "sulinet.hu" stringet "edu.hu"-ra. A régi e-mail címet meghagyja alaias-ként.
#
# Készítette: Venczel József
#
# Debrecen, 2020. 11. 07.
####################################################################################


from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
import socket
import time
import sys

SCOPES = ['https://www.googleapis.com/auth/admin.directory.user',
          'https://www.googleapis.com/auth/admin.directory.group'
]

mit = "sulinet.hu"
mire = "edu.hu"

print("Kapcsolatfelvétel a szerverrel...")

creds = None
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)

    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

socket.setdefaulttimeout(600)

serviceAdmin = build('admin', 'directory_v1', credentials=creds)

if serviceAdmin:
    print("Szerverhez kapcsolódva.")
else:
    print("Nem sikerült kapcsolatot létesíteni a szerverrel!")
    sys.exit()


#--- Tesztelő

# lista = serviceAdmin.users().list(customer="my_customer", maxResults=10, query="email:teszt*").execute()
# for sor in lista.get("users"):
#     email=str(sor.get("primaryEmail"))
#     id=str(sor.get("id"))
#     aliases=sor.get("aliases")
#     aliases.append("hejho")
#     print(email, aliases)


# sys.exit()

def cb_modFelh(request_id, response, exception):

    if exception is not None:
        id=request_id
        felh = serviceAdmin.users().get(id).execute()
        email=str(felh.get("primaryEmail"))
        primail=email.replace(mit, mire)
        alnevek=felh.get("aliases")
        alnevek.append(email)
        felh = { "primaryEmail" : primail,
                 "aliases" : [email]
        }

        ismeteld=True
        varido=1

        while ismeteld and varido < 64:
            try:
                tanar = serviceAdmin.users().update(userKey=id, body=felh).execute()
                ismeteld=False
            except HttpError as err:
                if err.resp.status in [409]:
                    ismeteld=False
                elif err.resp.status in [403, 503]:
                    time.sleep(varido+random())
                    varido=varido*2
                else:
                    print(err)
                    sys.exit()
            print("*", end="")
        if ismeteld:
            print(err)
            sys.exit()


totalTime=time.time()

print("Felhasználók e-mail címének módosítása")
batch = serviceAdmin.new_batch_http_request(callback=cb_modFelh)
kotegmeret=30
page_token=""
db=0
while page_token or db == 0:

    lista = serviceAdmin.users().list(customer="my_customer", pageToken=page_token, maxResults=30, query="email:teszt*").execute()
    page_token=lista.get("nextPageToken")

    for sor in lista.get("users"):
        email=str(sor.get("primaryEmail"))
        id=str(sor.get("id"))
        #print(email)

        if email.find("teszt") == 0 and email.find(mit) >= 0:
            primail=email.replace(mit, mire)
            alnevek=sor.get("aliases")
            alnevek.append(email)
            felh = { "primaryEmail" : primail,
                     "aliases" : alnevek
            }
            request = serviceAdmin.users().update(userKey=id, body=felh)
            batch.add(request, request_id=id)
            kotegmeret-=1
            print(".", end ="")
            if kotegmeret < 1:
                batch.execute()
                kotegmeret=30
                batch = serviceAdmin.new_batch_http_request(callback=cb_modFelh)
            db+=1

print(".")
batch.execute()
print("\nMódosított e-mail címek száma:", db)







def cb_modCsop(request_id, response, exception):

    if exception is not None:
        id=request_id
        csop = serviceAdmin.groups().get(id).execute()
        email=str(felh.get("email"))
        primail=email.replace(mit, mire)
        alnevek=csop.get("aliases")
        alnevek.append(email)
        csop = { "email" : primail,
                 "aliases" : alnevek
        }

        ismeteld=True
        varido=1

        while ismeteld and varido < 64:
            try:
                tanar = serviceAdmin.groups().update(groupKey=id, body=csop).execute()
                ismeteld=False
            except HttpError as err:
                if err.resp.status in [409]:
                    ismeteld=False
                elif err.resp.status in [403, 503]:
                    time.sleep(varido+random())
                    varido=varido*2
                else:
                    print(err)
                    sys.exit()
            print("*", end="")
        if ismeteld:
            print(err)
            sys.exit()


print("Csoportok e-mail címének módosítása")
batch = serviceAdmin.new_batch_http_request(callback=cb_modCsop)
kotegmeret=30
page_token=""
db=0
while page_token or db == 0:

    lista = serviceAdmin.groups().list(customer="my_customer", pageToken=page_token, maxResults=30, query="email:teszt*").execute()
    page_token=lista.get("nextPageToken")

    for sor in lista.get("groups"):
        email=str(sor.get("email"))
        id=str(sor.get("id"))

        if email.find("teszt") == 0 and email.find(mit) >= 0:
            primail=email.replace(mit, mire)
            alnevek=sor.get("aliases")
            alnevek.append(email)
            csop = { "email" : primail,
                     "aliases" : alnevek
            }
            request = serviceAdmin.groups().update(groupKey=id, body=csop)
            batch.add(request, request_id=id)
            kotegmeret-=1
            print(".", end ="")
            if kotegmeret < 1:
                batch.execute()
                kotegmeret=30
                batch = serviceAdmin.new_batch_http_request(callback=cb_modCsop)
            db+=1

print(".")
batch.execute()
print("\nMódosított csoport e-mail címek száma:", db)


elapsedTime = time.time() - totalTime
print("Ez a program %.3fs alatt futott le." % elapsedTime)