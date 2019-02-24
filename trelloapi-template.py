#!/usr/bin/env python
# coding: utf-8

import datetime
now = datetime.datetime.now()
now = now.strftime("%d_%m_%Y")


# In[3]:


import requests
import pandas as pd
from pandas.io.json import json_normalize
import json
import warnings
warnings.filterwarnings("ignore")

# Insert Key & Token user, respectively

key = ""
token = ""
user = ""

# List all Boards id

base = 'https://trello.com/1/'
boards_url = base + 'members/'+user+'/boards/'
params_key_and_token = {'key':key,'token':token}
arguments = {'fields': 'all', 'lists': 'all', 'cards' : 'all'}
response = requests.get(boards_url, params=params_key_and_token, data=arguments)
response_array_of_dict = response.json()

boards_attr = pd.DataFrame.from_dict(json_normalize(response_array_of_dict), orient='columns')

# Get (needed) Boards Informations

url = base + 'boards/'
contents = ['/actions','/cards','/checklists','/lists','/members']
all_attr = []
for content in contents:
    exec('%s = []' % (content[1:]+'_links'))
    exec('%s = pd.DataFrame()' % ('df_board_'+content[1:]))
    for board in response_array_of_dict:
        cmplt = url+board['id']+content
        exec('%s.append(cmplt)' % (content[1:]+'_links'))
    exec('all_attr.append(%s)' % (content[1:]+'_links'))


for j in range(len(all_attr)):
    for i in range(len(all_attr[j])):
        response = requests.get(all_attr[j][i], params=params_key_and_token, data=arguments)
        response_array_of_dict = response.json()
        df_id = pd.DataFrame.from_dict(json_normalize(response_array_of_dict), orient='columns')
        df_id['BoardName'] = boards_attr['name'][i]
        df_id['dateLastActivityBoard'] =  boards_attr['dateLastActivity'][i]
        df_id['dateLastViewBoard'] = boards_attr['dateLastView'][i]
        exec('%s = pd.concat([%s,df_id], ignore_index = True)' % ('df_board_'+contents[j][1:],'df_board_'+contents[j][1:]))

with pd.ExcelWriter('trello_board_'+now+'.xlsx') as writer:
    for j in range(len(contents)):
        exec('%s.to_excel(writer, sheet_name=%s)' % ('df_board_'+contents[j][1:],'contents[j][1:]'))

# Get (needed) Cards Informations

url = base + 'cards/'
contents = ['/actions','/attachments','/checkItemStates','/checklists','/list','/members','/membersVoted','/pluginData']
all_attr = []
for content in contents:
    exec('%s = []' % (content[1:]+'_links'))
    exec('%s = pd.DataFrame()' % ('df_card_'+content[1:]))
    for card in df_board_cards['id']:
        cmplt = url+card+content
        exec('%s.append(cmplt)' % (content[1:]+'_links'))
    exec('all_attr.append(%s)' % (content[1:]+'_links'))


for j in range(len(all_attr)):
    for i in range(len(all_attr[j])):
        response = requests.get(all_attr[j][i], params=params_key_and_token, data=arguments)
        response_array_of_dict = response.json()
        df_id = pd.DataFrame.from_dict(json_normalize(response_array_of_dict), orient='columns')
        exec('%s = pd.concat([%s,df_id], ignore_index = True)' % ('df_card_'+contents[j][1:],'df_card_'+contents[j][1:]))

with pd.ExcelWriter('trello_card_'+now+'.xlsx') as writer:
    for j in range(len(contents)):
        exec('%s.to_excel(writer, sheet_name=%s)' % ('df_card_'+contents[j][1:],'contents[j][1:]'))

