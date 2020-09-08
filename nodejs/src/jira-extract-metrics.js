#!/usr/bin/env node
const { printError,printInfo,consoleError } = require('./print');
const axios = require('axios');
const ExcelJS = require('exceljs');
const { query } = require('express');
const { fill } = require('lodash');
import percentile from 'just-percentile';

const TDC_JIRA_BOARD_ID = 217
const EXCEL_FILE_NAME = "jira-report-js-full.xlsx"
const JIRA_URL = "https://jira.talendforge.org/rest/api/2/search"
const TDC_JIRA_SPRINT_PAGINATION = 30
const STR_KEY="Issue key"
const STR_TYPE="Issue Type"
const STR_SUMMARY="Issue Summary"
const STR_CREATIONDATE="Creation Date"
const STR_RESODATE="Resolution Date"
const STR_NEWDATE="In New Date"
const STR_CANDIDATEDATE="In Candidate Date"
const STR_ACCEPTDATE="In Accepted Date"
const STR_PROGRESSDATE="In in progress Date"
const STR_REVIEWDATE="In Code review Date"
const STR_VALIDDATE="In Validation Date"
const STR_MERGEDATE="In Merge Date"
const STR_FINALCDATE="In Final check Date"
const STR_DONEDATE="In Done Date"
const STR_CLOSEDDATE="In Closed Date"
const STR_ONHOLDDATE="In On hold Date"
const STR_TODODATE="In To Do Date"
const STR_BLOCKEDDATE="In Blocked Date"
const STR_REJECTEDDATE="In Rejected Date"
const STR_NEWTIME="In New Time"
const STR_CANDIDATETIME="In Candidate Time"
const STR_ACCEPTTIME="In Accepted Time"
const STR_PROGRESSTIME="In in progress Time"
const STR_REVIEWTIME="In Code review Time"
const STR_VALIDTIME="In Validation Time"
const STR_MERGETIME="In Merge Time"
const STR_FINALCTIME="In Final check Time"
const STR_DONETIME="In Done Time"
const STR_CLOSEDTIME="In Closed Time"
const STR_BLOCKEDTIME="In Blocked Time"
const STR_REJECTEDTIME="In Rejected Time"
const STR_ONHOLDTIME="In On hold Time"
const STR_TODOTIME="In To Do Time"
const STR_LEADTIME="Lead time"
const STR_CYCLETIME="Cycle time"
const STR_KEY_RESOLVED="Issue key resolved"
const STR_RESOLUTION_DATE_RESOLVED="Resolution date for resolved"
const STR_LEADTIME_RESOLVED="Lead time resolved"
const STR_CYCLETIME_RESOLVED="Cycle time resolved"
const STR_CENTILE_20TH_CYCLETIME="Centile 20th"
const STR_CENTILE_50TH_CYCLETIME="Centile 50th"
const STR_CENTILE_80TH_CYCLETIME="Centile 80th"
const STR_CYCLETIMERANGE="Cycle time Range"
const STR_CYCLETIMEDISTRIBUTION="Cycle time Distribution"

const worksheetName = 'Raw_Metrics'

const filterLowCycleTime = 0.05
const filterHighCycleTime = 90

// const [ , , ...args ] = process.argv; // remove 2 first params

const getJIRAVariables = () => {
    let { JIRA_LOGIN, JIRA_PASSWORD } = process.env;
    if (!JIRA_LOGIN) {
        JIRA_LOGIN="eguillossou"
        printInfo(`CONFIG: JIRA_LOGIN environment variable set to default: `+JIRA_LOGIN)
    }
    if (!JIRA_PASSWORD) {
        printError(`CONFIG ERROR: JIRA_PASSWORD environment variable not set.`);
    }

    return { login: JIRA_LOGIN, password: decodeURIComponent(escape( Buffer.from(JIRA_PASSWORD, 'base64').toString() )) };
}


const switchDateOrTime = (transition,isDate=true) => {
    switch(transition) {
        case 'To Do': return(isDate? STR_TODODATE:STR_TODOTIME);
        case 'New': return(isDate? STR_NEWDATE:STR_NEWTIME);
        case 'Candidate':return(isDate? STR_CANDIDATEDATE:STR_CANDIDATETIME);
        case 'Accepted':return(isDate? STR_ACCEPTDATE:STR_ACCEPTTIME);
        case 'In Progress':return(isDate? STR_PROGRESSDATE:STR_PROGRESSTIME);
        case 'Code Review':return(isDate? STR_REVIEWDATE:STR_REVIEWTIME);
        case 'Validation':return(isDate? STR_VALIDDATE:STR_VALIDTIME);
        case 'Merge':return(isDate? STR_MERGEDATE:STR_MERGETIME);
        case 'Final Check':return(isDate? STR_FINALCDATE:STR_FINALCTIME);
        case 'Done':return(isDate? STR_DONEDATE:STR_DONETIME);
        case 'Closed':return(isDate? STR_CLOSEDDATE:STR_CLOSEDTIME);
        case 'Blocked':return(isDate? STR_BLOCKEDDATE:STR_BLOCKEDTIME);
        case 'Rejected':return(isDate? STR_REJECTEDDATE:STR_REJECTEDTIME);
        case 'On hold':return(isDate? STR_ONHOLDDATE:STR_ONHOLDTIME);
        default: return(`\nWARNING: unknown transition ${isDate? "date":"time"}: ${transition}`)
    }
}
const formatDateFromDays = (nbOfDays) => {
    const nbOfRemainingDays = nbOfDays-Math.floor(nbOfDays);
    const nbHours = Math.floor(nbOfRemainingDays*24);
    const nbOfRemainingHours = nbOfRemainingDays*24 - Math.floor(nbOfRemainingDays*24);
    const nbMinutes = Math.floor(nbOfRemainingHours*60);
    const nbOfRemainingMinutes = nbOfRemainingHours*60 - Math.floor(nbOfRemainingHours*60);
    const nbSec = Math.floor(nbOfRemainingMinutes*60);
    const strNbays = Math.floor(nbOfDays);
    const strNbHours = nbHours < 10 ? `0${nbHours}` : nbHours;
    const strNbMinutes = nbMinutes < 10 ? `0${nbMinutes}` : nbMinutes;
    const strNbSec = nbSec < 10 ? `0${nbSec}` : nbSec;
    return(`${strNbays} Days ${strNbHours}:${strNbMinutes}:${strNbSec}`);
}

const jsontoexcel = async (json) => {
    // printInfo(JSON.stringify(json.issues))
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(worksheetName, {views:[{state: 'frozen', xSplit: 1, ySplit:1}]});
    sheet.columns = [
        {header: STR_KEY, key:STR_KEY, width: '25'},
        {header: STR_TYPE, key:STR_TYPE, width: '25'},
        {header: STR_SUMMARY, key:STR_SUMMARY, width: '25'},
        {header: STR_CREATIONDATE, key:STR_CREATIONDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_RESODATE, key:STR_RESODATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_NEWDATE, key:STR_NEWDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_CANDIDATEDATE, key:STR_CANDIDATEDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_ACCEPTDATE, key:STR_ACCEPTDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_PROGRESSDATE, key:STR_PROGRESSDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_REVIEWDATE, key:STR_REVIEWDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_VALIDDATE, key:STR_VALIDDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_MERGEDATE, key:STR_MERGEDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_FINALCDATE, key:STR_FINALCDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_DONEDATE, key:STR_DONEDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_CLOSEDDATE, key:STR_CLOSEDDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_ONHOLDDATE, key:STR_ONHOLDDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_TODODATE, key:STR_TODODATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_BLOCKEDDATE, key:STR_BLOCKEDDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_REJECTEDDATE, key:STR_REJECTEDDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: STR_NEWTIME, key:STR_NEWTIME, width: '25'},
        {header: STR_CANDIDATETIME, key:STR_CANDIDATETIME, width: '25'},
        {header: STR_ACCEPTTIME, key:STR_ACCEPTTIME, width: '25'},
        {header: STR_PROGRESSTIME, key:STR_PROGRESSTIME, width: '25'},
        {header: STR_REVIEWTIME, key:STR_REVIEWTIME, width: '25'},
        {header: STR_VALIDTIME, key:STR_VALIDTIME, width: '25'},
        {header: STR_MERGETIME, key:STR_MERGETIME, width: '25'},
        {header: STR_FINALCTIME, key:STR_FINALCTIME, width: '25'},
        {header: STR_DONETIME, key:STR_DONETIME, width: '25'},
        {header: STR_CLOSEDTIME, key:STR_CLOSEDTIME, width: '25'},
        {header: STR_BLOCKEDTIME, key:STR_BLOCKEDTIME, width: '25'},
        {header: STR_REJECTEDTIME, key:STR_REJECTEDTIME, width: '25'},
        {header: STR_ONHOLDTIME, key:STR_ONHOLDTIME, width: '25'},
        {header: STR_TODOTIME, key:STR_TODOTIME, width: '25'},
        {header: STR_LEADTIME, key:STR_LEADTIME, width: '25'},
        {header: STR_CYCLETIME, key:STR_CYCLETIME, width: '25'},
        {header: STR_KEY_RESOLVED, key:STR_KEY_RESOLVED, width: '25'},
        {header: STR_RESOLUTION_DATE_RESOLVED, key:STR_RESOLUTION_DATE_RESOLVED, width: '25'},
        {header: STR_LEADTIME_RESOLVED, key:STR_LEADTIME_RESOLVED, width: '25'},
        {header: STR_CYCLETIME_RESOLVED, key:STR_CYCLETIME_RESOLVED, width: '25'},
        {header: STR_CENTILE_20TH_CYCLETIME, key:STR_CENTILE_20TH_CYCLETIME, width: '25'},
        {header: STR_CENTILE_50TH_CYCLETIME, key:STR_CENTILE_50TH_CYCLETIME, width: '25'},
        {header: STR_CENTILE_80TH_CYCLETIME, key:STR_CENTILE_80TH_CYCLETIME, width: '25'},
        {header: STR_CYCLETIMERANGE, key:STR_CYCLETIMERANGE, width: '25'},
        {header: STR_CYCLETIMEDISTRIBUTION, key:STR_CYCLETIMEDISTRIBUTION, width: '25'},
    ]
    // Cross all issues retrieved from JIRA jql query
    for(let issueIdx in json.issues){
        const issue = json.issues[issueIdx];
        statusUpdate={};
        const lastRow = sheet.lastRow;
        const newRow = sheet.addRow(++(lastRow.number));
        newRow.getCell(STR_KEY).value = issue.key;
        newRow.commit();
        printInfo(`Analysing ${newRow.getCell(STR_KEY).value} item`)
        
        // # Get datetime creation
        const datetime_creation = issue.fields.created ? new Date(issue.fields.created) : undefined;
        if(datetime_creation !== undefined) {
            newRow.getCell(STR_CREATIONDATE).value = datetime_creation.toLocaleString("fr-FR",{day:"numeric",month:"numeric",year:"numeric"});
            newRow.getCell(STR_NEWDATE).value = datetime_creation.toLocaleString();
        }
        
        // # Get datetime resolution
        const datetime_resolution = issue.fields.resolutiondate ? new Date(issue.fields.resolutiondate) : undefined;
        if(datetime_resolution !== undefined) {
            newRow.getCell(STR_RESODATE).value = datetime_resolution.toLocaleString("fr-FR",{day:"numeric",month:"numeric",year:"numeric"});
            const nbOfDays = (datetime_resolution.getTime()-datetime_creation.getTime()) / (1000*60*60*24)
            // newRow.getCell(STR_LEADTIME).value = formatDateFromDays(nbOfDays);
            newRow.getCell(STR_LEADTIME).value = nbOfDays;
        }
        
        newRow.getCell(STR_TYPE).value = issue.fields.issuetype.name;
        newRow.getCell(STR_SUMMARY).value = issue.fields.summary;
        
        var previousStatusChangeDate = datetime_creation;
        let historyA = issue.changelog.histories;
        for(let historyIdx in historyA) {
            for(let itemKey in historyA[historyIdx]['items']){
                const item = historyA[historyIdx]['items'][itemKey];
                if(item.field !== null && item.field === "status") {
                    if(!(item.fromString in statusUpdate)) {
                        statusUpdate[item.fromString] = 0;
                    }
                    if(!(item.toString in statusUpdate)) {
                        statusUpdate[item.toString] = 0;
                    }
                    let dateTransition = new Date(historyA[historyIdx].created);
                    statusUpdate[item.fromString] += (dateTransition - previousStatusChangeDate);
                    previousStatusChangeDate = dateTransition;
                    
                    newRow.getCell(switchDateOrTime(item.fromString)).value = dateTransition.toLocaleString();
                    // newRow.getCell(switchDateOrTime(item.fromString,false)).value = formatDateFromDays(statusUpdate[item.fromString]/(1000*60*60*24));
                    newRow.getCell(switchDateOrTime(item.fromString,false)).value = statusUpdate[item.fromString]/(1000*60*60*24);
                }
            }
        }
        // CYCLE Time series is a sum of in progress/code review/validation/merge/final check time
        newRow.getCell(STR_CYCLETIME).value = 
        newRow.getCell(STR_PROGRESSTIME).value+
        newRow.getCell(STR_REVIEWTIME).value+
        newRow.getCell(STR_VALIDTIME).value+
        newRow.getCell(STR_MERGETIME).value+
        newRow.getCell(STR_FINALCTIME).value;
        
    }//end parsing issues

    //fill Cycle time and lead time for 
    
    //fill Distribution
    const CTCol = sheet.getColumn(STR_CYCLETIME);
    //Range to fill: 15 values, step 3
    // const range = Array(15).fill(0).map((x,y)=>x+y*3);
    // const freq = Array(range.length);
    var freq = {}
    const step = 3
    
    CTCol.eachCell(function(cell, rowNumber) {
        if(cell.value !== STR_CYCLETIME && 
            cell.value !== 0 &&
            sheet.getColumn(STR_RESODATE).values[rowNumber] !== undefined &&
            sheet.getColumn(STR_RESODATE).values[rowNumber] !== STR_RESODATE &&
            cell.value<filterHighCycleTime &&
            cell.value>=filterLowCycleTime) {
                let rangeIdx = Math.floor(cell/step);
                if(rangeIdx in freq){
                    freq[rangeIdx] = freq[rangeIdx]+1;
                } else {
                    freq[rangeIdx] = 1;
                }
        }
    });

    //fill Resolved issue metrics
    const resDateCol = sheet.getColumn(STR_RESODATE);
    let sortedColumns = []
    let indexResoDate = 2;
    resDateCol.eachCell( function(cell,rowNumber) {
        if( cell.value !== STR_RESODATE && 
            cell.value !== null ) {
                sortedColumns.push(
                    { 
                    "key":sheet.getColumn(STR_KEY).values[rowNumber],
                    "resolution":sheet.getColumn(STR_RESODATE).values[rowNumber],
                    "cycletime": sheet.getColumn(STR_CYCLETIME).values[rowNumber],
                    "leadtime": sheet.getColumn(STR_LEADTIME).values[rowNumber]
                });
                indexResoDate = indexResoDate + 1 ;
        }
    });

    //sort array by resolution date
    sortedColumns.sort((a,b) => new Date(a.resolution).getTime() - new Date(b.resolution).getTime());
    sortedColumns.forEach((value,index) => {
        currentRow = sheet.getRow(index+2);
        currentRow.getCell(STR_KEY_RESOLVED).value = value.key;
        currentRow.getCell(STR_RESOLUTION_DATE_RESOLVED).value = value.resolution;
        currentRow.getCell(STR_CYCLETIME_RESOLVED).value = (Math.round(value.cycletime * 100)/100).toFixed(2);
        currentRow.getCell(STR_LEADTIME_RESOLVED).value = (Math.round(value.leadtime * 100)/100).toFixed(2);
    });

    //fill centile 20th | 50th | 80th of cycle time
    rowTwo =  sheet.getRow(index+2);
    rowTwo.getCell(STR_CENTILE_20TH_CYCLETIME).value = ;
    rowTwo.getCell(STR_CENTILE_20TH_CYCLETIME).value = ;
    rowTwo.getCell(STR_CENTILE_20TH_CYCLETIME).value = ;
        
    const maxKey = Math.max(...Object.keys(freq));
    for(let steps = 0;steps<maxKey+2;steps++) {
        //jump first row <=> title
        currentRow = sheet.getRow(steps+2);
        currentRow.getCell(STR_CYCLETIMERANGE).value = steps*step;
        currentRow.getCell(STR_CYCLETIMEDISTRIBUTION).value = freq[steps] !== undefined ? freq[steps]:0;
    }
    
        //after filling all raw metrics, split with group row
        const STR_GRP_1 = { "title":'Raw metrics', "keyStart":STR_KEY, "keyEnd":STR_CYCLETIME};
        const STR_GRP_2 = { "title":'Raw metrics resolved issues only', "keyStart":STR_KEY_RESOLVED, "keyEnd":STR_CYCLETIME_RESOLVED};
        const STR_GRP_3 = { "title":'Cycle time distribution', "keyStart":STR_CYCLETIMERANGE, "keyEnd":STR_CYCLETIMEDISTRIBUTION};

        let grpRow = [];
        grpRow[sheet.getColumn(STR_GRP_1.keyStart).number] = STR_GRP_1.title;
        grpRow[sheet.getColumn(STR_GRP_2.keyStart).number] = STR_GRP_2.title;
        grpRow[sheet.getColumn(STR_GRP_3.keyStart).number] = STR_GRP_3.title;
        
        sheet.insertRow(1,grpRow);
        
        sheet.mergeCells(1,sheet.getColumn(STR_GRP_1.keyStart).number,1,sheet.getColumn(STR_GRP_1.keyEnd).number);
        sheet.mergeCells(1,sheet.getColumn(STR_GRP_2.keyStart).number,1,sheet.getColumn(STR_GRP_2.keyEnd).number);
        sheet.mergeCells(1,sheet.getColumn(STR_GRP_3.keyStart).number,1,sheet.getColumn(STR_GRP_3.keyEnd).number);

        const getCellForStyle = (_sheet, _key, _rowNumber) => {
            return(_sheet.getRow(_rowNumber).getCell(_sheet.getColumn(_key).number));
        }
        const cell1 = getCellForStyle(sheet,STR_GRP_1.keyStart,1);
        const cell2 = getCellForStyle(sheet,STR_GRP_2.keyStart,1);
        const cell3 = getCellForStyle(sheet,STR_GRP_3.keyStart,1);
        const centerMiddleStyle = { vertical: 'middle', horizontal: 'center' };
        cell1.alignment = centerMiddleStyle;
        cell2.alignment = centerMiddleStyle;
        cell3.alignment = centerMiddleStyle;
        cell1.fill = {
            type: 'pattern',
            pattern:'solid',
            fgColor:{argb:'ccf2ff'},
        };
        cell2.fill = {
            type: 'pattern',
            pattern:'solid',
            fgColor:{argb:'f2d9e6'},
        };
        cell3.fill = {
            type: 'pattern',
            pattern:'solid',
            fgColor:{argb:'ccffcc'},
        };

        // write to a file
        writeExcelFile(workbook);
    }

const writeExcelFile = async (workbook) => {
    try {
        await workbook.xlsx.writeFile(EXCEL_FILE_NAME);
    }
    catch (error) {
        consoleError(error);
    }
}

function main() {
    const { login, password } = getJIRAVariables();
    // axios.interceptors.request.use(request => {
    //     console.log('Starting Request', request)
    //     return request
    //   })

    axios({
        method: 'get',
        withCredentials: true,
        headers: {
            "Accept": "application/json",
            "Content-Type": "application/json"
        },
        params: {
            "expand":"changelog", // only accessible through get request and not in post jira api
            "jql": "project = TDC AND issuetype in (Bug, \"New Feature\", \"Work Item\") AND Sprint in (\"TDC Sprint 48\",\"TDC Sprint 49\",\"TDC Sprint 50\") ORDER BY labels ASC, RANK",
            "startAt":0,
            "maxResults":500,
        },
        url: `${JIRA_URL}`,
        auth: {
            username: login,
            password: password
        },
    }).then(function (response) {
        try {
            jsontoexcel(response.data);
        } catch(error) {
            consoleError(error);
        }
    })
    .catch(function (error) {
        consoleError(error);
    });

}

main();