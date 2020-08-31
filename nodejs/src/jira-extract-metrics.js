#!/usr/bin/env node
const { printError,printInfo,consoleError } = require('./print');
const axios = require('axios');

const TDC_JIRA_BOARD_ID = 217
const EXCEL_FILE_NAME = "jira-report"
const JIRA_URL = "https://jira.talendforge.org/rest/api/2/search"
const TDC_JIRA_SPRINT_PAGINATION = 30

function getJIRAVariables() {
    let { JIRA_LOGIN, JIRA_PASSWORD } = process.env;
    if (!JIRA_LOGIN) {
        JIRA_LOGIN="eguillossou"
        printInfo(`CONFIG: JIRA_LOGIN environment variable set to default: `+JIRA_LOGIN)
    }
    if (!JIRA_PASSWORD) {
        printError(`CONFIG ERROR: JIRA_PASSWORD environment variable not set.`);
    }

    return { login: JIRA_LOGIN, password: JIRA_PASSWORD };
}

const parsejson = (json) => {
    printInfo(json.issues);
}

function main() {
    console.log("hello world!");
    const { login, password } = getJIRAVariables();
    axios({
        method: 'post',
        withCredentials: true,
        headers: {
            "Accept": "application/json",
            "Content-Type": "application/json"
        },
        params: {
            "expand":"changelog",
        },
        url: `${JIRA_URL}`,
        auth: {
            username: login,
            password: password
        },
        data: {
            "jql": "project = TDC AND issuetype in (Bug, \"New Feature\", \"Work Item\") AND Sprint = \"TDC Sprint 50\" ORDER BY labels ASC, RANK",
            "startAt":0,
            "maxResults":500,
        }
    }).then(function (response) {
        // parsejson(JSON.stringify(response.data))
        parsejson(response.data)
    })
    .catch(function (error) {
        consoleError(error);
    });

}

main();