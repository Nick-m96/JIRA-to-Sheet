# JIRA to Sheet


Creation and automation of Metrics in Projects using Jira Software.
Using Google Sheet as a support we can see the status of the functionalities, the changes of state that occurred, how long it was in each one and how many rework the team performed.

## Installing

A quick introduction of the minimal setup you need to get a hello world up &
running.

### Install Python3 if necessary

[Python Page](https://www.python.org/downloads/)

### Getting started

```shell
git clone https://github.com/NickMano/JIRA-to-Sheet.git
cd JIRA-to-Sheet

### Recomended ###
python3 -m venv venv 
source venv/bin/activate
### Recomended ### 

pip install -r requirements.txt
```

### Initial Configuration

Rename `example.json` to `variables.json` and add all the parameters you need.
