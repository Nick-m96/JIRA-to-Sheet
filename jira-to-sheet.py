# -*- coding: utf-8 -*-

import gspread
import json
from jira import JIRA
import time
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

def rename_tech(text):
    t = text.lower()
    for key in techs.keys():
        if t.find(key) != -1:
            return techs[key]
    return 'N/A'

def rename_assignee(name):
    for key in assignee.keys():
        if name.find(key) != -1:
            return assignee[key]
    return name

def rename_status(name):
    for key in status.keys():
        if name.find(key) != -1:
            return status[key]
    return name

def string_to_date(text):
    date = datetime.strptime(text, '%Y-%m-%dT%H:%M:%S.%f%z')
    return datetime.strptime(date.strftime('%m-%d-%y %H:%M:%S'), '%m-%d-%y %H:%M:%S')

def update_timeline(issues, is_mc, index):
    for issue in issues:
        if not is_excluded_status(issue.fields.status.name):
            status = rename_status(issue.fields.status.name)

            title_card = issue.fields.summary
            card_id = issue.key
            tech = rename_tech(title_card)

            try:
                assignee = rename_assignee(issue.fields.assignee.name)
            except AttributeError:
                assignee = 'Sin Asignacion'

            if is_mc:
                try:
                    board = issue.fields.components[0].name if issue.fields.components[0].name != 'General' else 'Evolutivas'
                    if board != 'Mejora Continua':
                        issue.update(fields={"components" : [{'name': 'Mejora Continua'}]})
                        board = 'Mejora Continua'
                except:
                    board = 'Mejora Continua'
                    issue.update(fields={"components" : [{'name': 'Mejora Continua'}]})
            else:       
                try:
                    board = issue.fields.components[0].name if issue.fields.components[0].name != 'General' else 'Evolutivas'
                except:
                    board = 'Sin Asignar'
                        
            row = [
                id_sprint,
                card_id,
                board,
                tech,
                title_card,
                assignee,
                status
            ]
            if (is_mc and status =='Done') or not is_mc:
                update_row(row, index, worksheet_timeline)
                index += 1
                time.sleep(1.73)
    return index

def update_row(data, index_row, ws):
    cell_list = ws.range('A{}:{}{}'.format(index_row,chr(64+len(data)),index_row))

    for i,cell in enumerate(cell_list):
        cell.value = data[i]

    ws.update_cells(cell_list,'USER_ENTERED')

def get_sort_request(sheetId,end_row, end_column):
    return {
        "requests": [
            {
            "setBasicFilter": {
                "filter": {
                "range": {
                    "sheetId": sheetId,
                    "endColumnIndex": end_column,
                    "endRowIndex": end_row,
                    "startColumnIndex": 0,
                    "startRowIndex": 0
                },
                "sortSpecs": [
                    {
                    "dimensionIndex": 1,
                    "sortOrder": "ASCENDING"
                    },
                    {
                    "dimensionIndex": 6,
                    "sortOrder": "ASCENDING"
                    }
                ]
                }
            }
            }
        ],
        "includeSpreadsheetInResponse": False,
        "responseIncludeGridData": False}

def time_in_transition(index):
    return '=IF(AND(F{}=E{}, B{}=B{}), (NETWORKDAYS(G{},G{})-1)*(Resume!$AC$2-Resume!$AC$1)+IF(NETWORKDAYS(G{},G{}),MEDIAN(MOD(G{},1),Resume!$AC$2,Resume!$AC$1),0)-MEDIAN(NETWORKDAYS(G{},G{})*MOD(G{},1),Resume!$AC$2,Resume!$AC$1),)'.format(index, index+1, index, index+1, index, index+1, index+1, index+1, index+1, index, index, index)

def is_excluded_status(status):
    return status in excluded_status

def is_qa_reject(index):
    return '=IF(AND(OR(E{}="Testing", E{}="Ready for Testing"), F{}="In Progress"), TRUE, FALSE)'.format(index,index,index)

def fill_timeline():
    values_list = worksheet_timeline.col_values(1)
    index = worksheet_timeline.find(id_sprint).row if id_sprint in values_list else len(values_list) + 1
    
    index = update_timeline(issues, False, index)
    update_timeline(issues_incident, True, index)

def fill_movements():
    values_list = worksheet_movements.col_values(1)
    index = worksheet_movements.find(id_sprint).row if id_sprint in values_list else len(values_list) + 1

    index = update_movements(issues, index)
    update_movements(issues_incident, index) 

def update_movements(issues, index): 
    start = datetime.strptime(start_sprint, '%Y-%m-%d')
    end = datetime.strptime(end_sprint, '%Y-%m-%d')
    row = []

    for issue in issues:
        card_id = issue.key
        title_card = issue.fields.summary
        tech = rename_tech(title_card)
        try:
            board = issue.fields.components[0].name if issue.fields.components[0].name != 'General' else 'Evolutivas'
        except:
            board = 'Sin Asignar'
        for history in issue.changelog.histories:
            item = history.items[-1]
            date = string_to_date(history.created)
            if item.field == 'status' and start < date < end:
                row = [
                    id_sprint,
                    card_id,
                    tech,
                    title_card,
                    item.fromString,
                    rename_status(item.toString), 
                    str(date),
                    time_in_transition(index),
                    is_qa_reject(index),
                    board
                ]
                update_row(row, index, worksheet_movements)
                index += 1
                time.sleep(1.73)

    sortRequest = spreadsheet.batch_update(get_sort_request(worksheet_movements.id, index-1, len(row)))
    return  index

if __name__ == '__main__':

    #json
    with open('variables.json') as js_file:
        data = json.load(js_file)
        techs = data['techs']
        assignee = data['assignee']
        status = data['status']
        excluded_status = data['excluded_status']
        start_sprint = data['start_sprint']
        end_sprint = data['end_sprint'] 
        id_sprint = data['number_sprint']
        jira_client = data['jira_server']
        jira_mail = data['jira_mail']
        jira_token = data['jira_token']
        jira_project = data['jira_project']

    #google sheets
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open("Dashboard")
    worksheet_movements = spreadsheet.worksheet('Movements')
    worksheet_timeline = spreadsheet.worksheet('Timeline')
    worksheet_errores = spreadsheet.worksheet('Errores')

    # JIRA
    jira = JIRA(server=jira_client, basic_auth=(jira_mail,jira_token))

    issues_story = jira.search_issues("project = {} AND status changed DURING({}, {}) AND issuetype = Story".format(jira_project, start_sprint, end_sprint), maxResults=100, expand='changelog')
    issues_incident = jira.search_issues("project = {} AND status changed DURING({}, {}) AND issuetype = Incidente".format(jira_project, start_sprint, end_sprint), maxResults=100, expand='changelog')
    issues_bug = jira.search_issues("project = {} AND status changed DURING({}, {}) AND issuetype = Bug".format(jira_project, start_sprint, end_sprint), maxResults=100, expand='changelog')
    
    print('Actualizando Timeline')
    fill_timeline()

    print('Actualizando Movements')
    fill_movements()

    