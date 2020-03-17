import requests 

from lxml import html

import json


quotes = []

url = 'https://mashable.com/article/best-game-of-thrones-one-liners/'

path = '//*[@id="main"]/div[1]/div/div[1]/div/div[2]/div/article/section/p/strong'

response = requests.get(url)

byte_data = response.content 

source_code = html.fromstring(byte_data) 

tree = source_code.xpath(path) 
 
for i in range(len(tree)):
    data = tree[i].text_content().split('"')
    saying = data[1]
    speaker = data[2][2:100]
    quotes.append({
    "id" : i+1,
    "speaker" : speaker,
    "saying" : saying
    })


for i in quotes:
    print(i["id"], i["speaker"] , i["saying"])

f = open("GOT.json", "w")
json.dump(quotes,f)
f.close()
    

