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

def update_timeline(issues, is_mc, index, worksheet):
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
                update_row(row, index, worksheet)
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

def update_movements(issues, index, worksheet, spreadsheet): 
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
                update_row(row, index, worksheet)
                index += 1
                time.sleep(1.73)

    sortRequest = spreadsheet.batch_update(get_sort_request(worksheet.id, index-1, len(row)))
    return  index

def update_bugs(issues, index, worksheet):
    for issue in issues:
        card_id = issue.key
        title_card = issue.fields.summary
        priority = issue.fields.priority.name

        issuelinks = issue.fields.issuelinks
        card_linked = next((issue.outwardIssue for issue in issuelinks if issue.type.name == 'Blocks'))
        card_linked_key = card_linked.key
        card_linked_name = card_linked.fields.summary
        card_linked_points = jira.issue(card_linked.key).fields.customfield_10005
        
        board = 'Evolutivas' if card_linked.fields.issuetype.name == 'Historia' else 'Mejora Continua'
                    
        row = [
            id_sprint,
            card_id,
            title_card,
            priority,
            card_linked_key,
            card_linked_name,
            card_linked_points,
            board
        ]

        update_row(row, index, worksheet)
        index += 1
        time.sleep(1.73)

def get_index(worksheet):
    values_list = worksheet.col_values(1)
    index = worksheet.find(id_sprint).row if id_sprint in values_list else len(values_list) + 1
    return index

def update_ev(client, jira):
    spreadsheet = client.open("Dashboard - Evolutivas")
    issues = jira.search_issues("project = {} AND status changed DURING({}, {}) AND issuetype = Story".format(jira_project, start_sprint, end_sprint), maxResults=100, expand='changelog')
    update_spreadsheet(spreadsheet, issues)

def update_mc(client, jira):
    spreadsheet = client.open("Dashboard - Mejora Continua")
    issues = jira.search_issues("project = {} AND status changed DURING({}, {}) AND issuetype = Incidente".format(jira_project, start_sprint, end_sprint), maxResults=100, expand='changelog')
    update_spreadsheet(spreadsheet, issues)

def update_spreadsheet(spreadsheet, issues):
    worksheet_movements = spreadsheet.worksheet('Movements')
    worksheet_timeline = spreadsheet.worksheet('Timeline')
    worksheet_errores = spreadsheet.worksheet('Errores')
    issues_bug = jira.search_issues("project = {} AND status changed DURING({}, {}) AND issuetype = Bug".format(jira_project, start_sprint, end_sprint), maxResults=100, expand='changelog')

    print('Actualizando Movements')
    update_movements(issues, get_index(worksheet_movements), worksheet_movements, spreadsheet)
    
    # print('Actualizando Timeline')
    # update_timeline(issues, False, get_index(worksheet_timeline))
    
    print('Actualizando Bugs')
    update_bugs(issues_bug)

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

    # JIRA
    jira = JIRA(server=jira_client, basic_auth=(jira_mail,jira_token))
    
    #Proyecto Evolutivas
    if( input('Actualizar Evolutivas? [s/n]') == 's'):
        update_ev(client, jira)

    #Proyecto Mejora Continua
    if(input('Actualizar Mejora Continua? [s/n]') == 's'):
        update_mc(client, jira)
        

    