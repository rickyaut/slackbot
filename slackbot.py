#!/usr/bin/env python
# -*- coding: utf-8 -*-
import random
import xlrd
import time
import pdb
import re
import requests
import shutil
import os

from slackclient import SlackClient
from datetime import date
from datetime import timedelta
from collections import OrderedDict

#get your personal token from https://api.slack.com/web, bottom of the page.
api_key = 'xoxb-xxxxxxxxx-086tDUascVsvtXFbHqucJJKw'
target_environments = ['PREPROD', 'UAT', 'PROD']
col_index_project = 2
row_index_head = 1

client = SlackClient(api_key)
projects = OrderedDict()

def parse_workbook(file_name):
    temp_projects = OrderedDict()
    workbook = xlrd.open_workbook(file_name)
    today = date.today()
    sheet = workbook.sheet_by_index(0)
    head_row = sheet.row(row_index_head)
    for col_index in range(col_index_project + 1, sheet.ncols):
        date_cell = sheet.cell(row_index_head, col_index)
        date_cell_value = date_cell.value
        dt_tuple = xlrd.xldate_as_tuple(date_cell_value, workbook.datemode)
        cell_date_value= date( dt_tuple[0], dt_tuple[1], dt_tuple[2])
        if cell_date_value >= today:
            for row_index in range(row_index_head + 1, sheet.nrows):
                target_cell = sheet.cell(row_index, col_index)
                target_cell_value = target_cell.value
                if(target_cell.ctype == 1 and target_cell_value != None and target_cell_value.strip() in target_environments):
                    project_cell_value = sheet.cell(row_index, col_index_project).value
                    project = temp_projects.get(project_cell_value)
                    if( project == None ):
                        project = OrderedDict()
                        temp_projects[project_cell_value] = project
                    project[cell_date_value] = target_cell_value
    return temp_projects


def get_msg_by_date(start_date):
    "get_msg_by_dates"
    deployments = OrderedDict([('PREPROD', []), ('UAT', []), ('PROD', [])])
    for project_name, project_info in projects.iteritems():
        for target_date, target in project_info.iteritems():
            if(target_date == start_date):
                deployments.get(target).append(project_name)
    msg = ''
    for target, project_names in deployments.iteritems():
        if(len(project_names)>0):
            msg += target + ": "
            for project_name in project_names:
                msg +=project_name + ", "
            msg += "\r\n"
    return msg



def get_msg_of_project_lists():
    'get_msg_of_project_lists'
    project_names = ''
    index = 1

    for project_name, project_info in projects.iteritems():
        project_names += "%s. %s\r\n" % (index, project_name)
        index += 1

    project_names += "\r\n\r\nType project number followed by '?' to know project detail"
    return project_names

def get_msg_by_project(project_index):
    "get_msg_by_project"
    index = 1
    for project_name, project_info in projects.iteritems():
        if(index == project_index):
            msg = project_name +"\r\n"
            current_target = ''
            for target_date, target in project_info.iteritems():
                if(target != current_target):
                    msg += "%s: %s\r\n" % (target, target_date.strftime('%d/%m/%Y'))
                    current_target = target
            return msg
        index += 1
    return "Sorry, I don't understand which project you are talking about"



if client.rtm_connect():
    holden_xlsx = 'abc.xlsx'
    projects = parse_workbook(holden_xlsx)
    while True:
        last_read = client.rtm_read()
        if last_read and last_read[0].get('type') == 'message':
            try:
                if last_read[0].get('file') is None or last_read[0].get('file').get('url_private') is None:
                    parsed = last_read[0]['text']
                    #reply to channel message was found in.
                    message_channel = last_read[0]['channel']
                    if parsed and 'projects' in parsed.lower():
                        msg = get_msg_of_project_lists()
                    elif parsed and len(re.compile('\d+').findall(parsed))>0:
                        msg = get_msg_by_project(int(re.compile('\d+').findall(parsed)[0]))
                    elif parsed and 'today' in parsed.lower():
                        msg = get_msg_by_date(date.today())
                    elif parsed and 'tomorrow' in parsed.lower():
                        msg = get_msg_by_date(date.today() + timedelta(days=1))
                    else:
                        msg ='Hello naughty tester, I don\'t understand what you want\r\n Please type "projects" to see the project list\r\nType "today" or "tomorrow" to see the deployment of these days'
                    client.rtm_send_message(message_channel, msg)
                elif last_read[0].get('file') is not None and last_read[0].get('file').get('url_private') is not None:
                    url = str(last_read[0]['file']['url_private'])
                    file_type = str(last_read[0]['file']['filetype'])
                    response = requests.get(url, headers={'Authorization': 'Bearer %s' % api_key}, stream = True)
                    temp_name = 'abc_temp.%s' % file_type
                    message_channel = last_read[0]['channel']
                    with open(temp_name, 'wb') as out_file:
                        response.raw.decode_content = True
                        shutil.copyfileobj(response.raw, out_file)
                    projects = parse_workbook(temp_name)
                    os.remove(holden_xlsx)
                    os.rename(temp_name, holden_xlsx)
                    client.rtm_send_message(message_channel, 'Thank you for uploading new schedule!')
                    del response
            except IOError as e:
                print e
            except:
                pass
        time.sleep(1)
