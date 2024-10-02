import requests

url = "" # Url da api

headers = {
    "accept": "", # Accept da API  json
    "authorization": "" #Token da API
}

response = requests.get(url, headers=headers)

print(response.text)