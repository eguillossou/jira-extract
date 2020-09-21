#!/usr/bin/env node
// const axios = require('axios');
const ExcelJS = require('exceljs');
const percentile = require('just-percentile');
const constants = require('./constants')
const rest = require('./services/rest');
// const { fill } = require('lodash');
const fs = require('fs');
const path = require('path');
const excel = require('./services/excel')
const { printError,printInfo,consoleError } = require('./print');

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
        case 'To Do': return(isDate? constants.STR_TODODATE:constants.STR_TODOTIME);
        case 'New': return(isDate? constants.STR_NEWDATE:constants.STR_NEWTIME);
        case 'Candidate':return(isDate? constants.STR_CANDIDATEDATE:constants.STR_CANDIDATETIME);
        case 'Accepted':return(isDate? constants.STR_ACCEPTDATE:constants.STR_ACCEPTTIME);
        case 'In Progress':return(isDate? constants.STR_PROGRESSDATE:constants.STR_PROGRESSTIME);
        case 'Code Review':return(isDate? constants.STR_REVIEWDATE:constants.STR_REVIEWTIME);
        case 'Validation':return(isDate? constants.STR_VALIDDATE:constants.STR_VALIDTIME);
        case 'Merge':return(isDate? constants.STR_MERGEDATE:constants.STR_MERGETIME);
        case 'Final Check':return(isDate? constants.STR_FINALCDATE:constants.STR_FINALCTIME);
        case 'Done':return(isDate? constants.STR_DONEDATE:constants.STR_DONETIME);
        case 'Closed':return(isDate? constants.STR_CLOSEDDATE:constants.STR_CLOSEDTIME);
        case 'Blocked':return(isDate? constants.STR_BLOCKEDDATE:constants.STR_BLOCKEDTIME);
        case 'Rejected':return(isDate? constants.STR_REJECTEDDATE:constants.STR_REJECTEDTIME);
        case 'On hold':return(isDate? constants.STR_ONHOLDDATE:constants.STR_ONHOLDTIME);
        default: return(`\nWARNING: unknown transition ${isDate? "date":"time"}: ${transition}`)
    }
}

// const formatDateFromDays = (nbOfDays) => {
//     const nbOfRemainingDays = nbOfDays-Math.floor(nbOfDays);
//     const nbHours = Math.floor(nbOfRemainingDays*24);
//     const nbOfRemainingHours = nbOfRemainingDays*24 - Math.floor(nbOfRemainingDays*24);
//     const nbMinutes = Math.floor(nbOfRemainingHours*60);
//     const nbOfRemainingMinutes = nbOfRemainingHours*60 - Math.floor(nbOfRemainingHours*60);
//     const nbSec = Math.floor(nbOfRemainingMinutes*60);
//     const strNbays = Math.floor(nbOfDays);
//     const strNbHours = nbHours < 10 ? `0${nbHours}` : nbHours;
//     const strNbMinutes = nbMinutes < 10 ? `0${nbMinutes}` : nbMinutes;
//     const strNbSec = nbSec < 10 ? `0${nbSec}` : nbSec;
//     return(`${strNbays} Days ${strNbHours}:${strNbMinutes}:${strNbSec}`);
// }


// return an array of frequency for each values for a specific range with a step
// filter also values between a min and a high boundaries
// eg: for cycle time values, get frequency of values between 0 and 3 (step of 3 here)
const freqCellColumnByKeyColumn2 = (step, internalArray, cellColumn, columnKey ) => {
    let freq = {};
    internalArray.forEach( (value) => {
        if(value.resolutiondate!== undefined &&
            value.resolutiondate !== cellColumn) {
                let valueSelect = value.cycletime; 
                if(columnKey === constants.STR_LEADTIME) valueSelect = value.leadtime;
                if( valueSelect !== columnKey && 
                    valueSelect !== 0 &&
                    valueSelect < constants.FILTER_HIGH_CYCLETIME &&
                    valueSelect >= constants.FILTER_LOW_CYCLETIME) {
                    let rangeIdx = Math.floor(valueSelect/step);
                    if(rangeIdx in freq){
                        freq[rangeIdx] = freq[rangeIdx]+1;
                    } else {
                        freq[rangeIdx] = 1;
                    }
                }
        }
    });
    return(freq);
}

const parseIssues = (workbook, json) => {
    const sheet = workbook.getWorksheet(constants.WORKSHEET_NAME);
    const internalArray = [];
    
    // Cross all issues retrieved from JIRA jql query
    for(let issueIdx in json.issues){
        const issue = json.issues[issueIdx];
        statusUpdate={};
        const lastRow = sheet.lastRow;
        const newRow = sheet.addRow(++(lastRow.number));
        newRow.getCell(constants.STR_KEY).value = issue.key;
        newRow.commit();
        printInfo(`Analysing ${newRow.getCell(constants.STR_KEY).value} item`);
        const issueObject = {};

        issueObject.key = issue.key;
        
        // # Get datetime creation
        const datetime_creation = issue.fields.created ? new Date(issue.fields.created) : undefined;
        if(datetime_creation !== undefined) {
            newRow.getCell(constants.STR_CREATIONDATE).value = datetime_creation.toLocaleString("fr-FR",{day:"numeric",month:"numeric",year:"numeric"});
            newRow.getCell(constants.STR_NEWDATE).value = datetime_creation.toLocaleString();

            issueObject.creationdate = datetime_creation.toLocaleString("fr-FR",{day:"numeric",month:"numeric",year:"numeric"});
            issueObject.newdate = datetime_creation.toLocaleString();
        }
        
        // # Get datetime resolution
        const datetime_resolution = issue.fields.resolutiondate ? new Date(issue.fields.resolutiondate) : undefined;
        if(datetime_resolution !== undefined) {
            newRow.getCell(constants.STR_RESODATE).value = datetime_resolution.toLocaleString("fr-FR",{day:"numeric",month:"numeric",year:"numeric"});
            const nbOfDays = (datetime_resolution.getTime()-datetime_creation.getTime()) / (1000*60*60*24)
            // newRow.getCell(constants.STR_LEADTIME).value = formatDateFromDays(nbOfDays);
            newRow.getCell(constants.STR_LEADTIME).value = nbOfDays;

            issueObject.resolutiondate = datetime_resolution.toLocaleString("fr-FR",{day:"numeric",month:"numeric",year:"numeric"});
            issueObject.leadtime = nbOfDays;
        }
        
        newRow.getCell(constants.STR_TYPE).value = issue.fields.issuetype.name;
        newRow.getCell(constants.STR_SUMMARY).value = issue.fields.summary;

        issueObject.type = issue.fields.issuetype.name;
        issueObject.summary = issue.fields.summary;
        
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

                    issueObject[`${switchDateOrTime(item.fromString)}`] = dateTransition.toLocaleString();
                    issueObject[`${switchDateOrTime(item.fromString,false)}`] = statusUpdate[item.fromString]/(1000*60*60*24);
                }
            }
        }
        // CYCLE Time series is a sum of in progress/code review/validation/merge/final check time
        newRow.getCell(constants.STR_CYCLETIME).value = 
            newRow.getCell(constants.STR_PROGRESSTIME).value+
            newRow.getCell(constants.STR_REVIEWTIME).value+
            newRow.getCell(constants.STR_VALIDTIME).value+
            newRow.getCell(constants.STR_MERGETIME).value+
            newRow.getCell(constants.STR_FINALCTIME).value;
        issueObject.cycletime = 
            (issueObject[constants.STR_PROGRESSTIME] !== undefined ? issueObject[constants.STR_PROGRESSTIME]:0)+
            (issueObject[constants.STR_REVIEWTIME] !== undefined ? issueObject[constants.STR_REVIEWTIME]:0)+
            (issueObject[constants.STR_VALIDTIME] !== undefined ? issueObject[constants.STR_VALIDTIME]:0)+
            (issueObject[constants.STR_MERGETIME] !== undefined ? issueObject[constants.STR_MERGETIME]:0)+
            (issueObject[constants.STR_FINALCTIME] !== undefined ? issueObject[constants.STR_FINALCTIME]:0);
        
        internalArray.push(issueObject);
    }//end parsing issues
    //return(internalArray);
        
    //fill Distribution Cycle time
    let index = 2;
    fillDistributionCycleTime(internalArray).forEach(
        value => {
            currentRow = sheet.getRow(index);
            currentRow.getCell(constants.STR_CYCLETIMERANGE).value = value.cycletimerange;
            currentRow.getCell(constants.STR_CYCLETIMEDISTRIBUTION).value = value.cycletimedistribution;
            index = index + 1; 
        }
    );

    //fill Distribution Lead time
    index = 2;
    fillDistributionLeadTime(internalArray).forEach(
        value => {
            currentRow = sheet.getRow(index);
            currentRow.getCell(constants.STR_LEADTIMERANGE).value = value.leadtimerange;
            currentRow.getCell(constants.STR_LEADTIMEDISTRIBUTION).value = value.leadtimedistribution;    
            index = index + 1; 
        }
    );
    
    //fill Cycle time and lead time 
    //fill Resolved issue metrics
    let sortedColumns = fillSortedColumn(workbook);
    let sortedColumns = fillSortedColumn2(internalArray);
        
    //sort array by resolution date and removing Too high values and too low as well
    let filteredColumn = sortedColumns
                            .filter(a => a.cycletime > constants.FILTER_LOW_CYCLETIME && a.cycletime <= constants.FILTER_HIGH_CYCLETIME)
                            .sort((a,b) => new Date(a.resolution).getTime() - new Date(b.resolution).getTime());
    
    //fill centile 20th | 50th | 80th of cycle time
    const centileThCycleTime = (centileTh) => {
        return (percentile(
            filteredColumn
            .map((value) => value.cycletime)
            .sort((a,b)=> a-b), centileTh));
    }
    //fill centile  50th | 80th of lead time
    const centileThLeadTime = (centileTh) => {
        return (percentile(
            filteredColumn
            .map((value) => value.leadtime)
            .sort((a,b)=> a-b), centileTh));
    }
    
    filteredColumn.forEach((value,index) => {
        currentRow = sheet.getRow(index+2);
        currentRow.getCell(constants.STR_KEY_RESOLVED).value = value.key;
        currentRow.getCell(constants.STR_RESOLUTION_DATE_RESOLVED).value = value.resolution;
        currentRow.getCell(constants.STR_CYCLETIME_RESOLVED).value = Number((Math.round(value.cycletime * 100)/100).toFixed(2));
        currentRow.getCell(constants.STR_LEADTIME_RESOLVED).value = Number((Math.round(value.leadtime * 100)/100).toFixed(2));
        currentRow.getCell(constants.STR_CENTILE_20TH_CYCLETIME).value = centileThCycleTime(20);
        currentRow.getCell(constants.STR_CENTILE_50TH_CYCLETIME).value = centileThCycleTime(50);
        currentRow.getCell(constants.STR_CENTILE_80TH_CYCLETIME).value = centileThCycleTime(80);
        currentRow.getCell(constants.STR_CENTILE_50TH_LEADTIME).value = centileThLeadTime(50);
        currentRow.getCell(constants.STR_CENTILE_80TH_LEADTIME).value = centileThLeadTime(80);
    });
    
}
    
const parseIdNameFromSprints = (json) => {
    // getting first row to fill (remove group row and header row)
    const arrSprint = [];
    const filterSprints = json.sprints.filter((sprint) => 
    sprint.name.includes(constants.STR_EXP_FILTER_SPRINT) && 
    ( sprint.state.includes("CLOSED") || sprint.state.includes("ACTIVE")));
    const lastTenSprints = filterSprints.filter((_, idx) => idx >filterSprints.length-11);
    for(let sprintNb in lastTenSprints) {
        arrSprint[sprintNb] = {
            "id": lastTenSprints[sprintNb].id,
            "name":lastTenSprints[sprintNb].name
        };
    }
    return(arrSprint);
}

const fillDistributionCycleTime = (internalArray) => {
    var freqCT = {};
    const distributionCycleTime = [];
    const step = 3;
    freqCT = freqCellColumnByKeyColumn2(step,internalArray,constants.STR_RESODATE, constants.STR_CYCLETIME);

    const maxKey = Math.max(...Object.keys(freqCT));
    for (let steps = 0; steps < maxKey + 2; steps++) {
        //jump first row <=> title
        distributionCycleTime.push({
            cycletimerange : steps*step,
            cycletimedistribution: freqCT[steps] !== undefined ? freqCT[steps] : 0
        });
    }
    return(distributionCycleTime)
}
const fillDistributionLeadTime = (internalArray) => {
    var freqLT = {};
    const distributionLeadTime = [];
    const stepLT = 5;
    freqLT = freqCellColumnByKeyColumn2(stepLT,internalArray,constants.STR_RESODATE, constants.STR_LEADTIME);

    const maxKeyLT = Math.max(...Object.keys(freqLT));
    for (let steps = 0; steps < maxKeyLT + 2; steps++) {
        //jump first row <=> title
        distributionLeadTime.push({
            leadtimerange : steps*stepLT,
            leadtimedistribution: freqLT[steps] !== undefined ? freqLT[steps] : 0
        });
    }
    return(distributionLeadTime)
}
// const fillSortedColumn = (workbook) => {
//     const sheet = workbook.getWorksheet(constants.WORKSHEET_NAME);
//     const resDateCol = sheet.getColumn(constants.STR_RESODATE);
//     let sortedColumns = [];
//     let indexResoDate = 2;
//     resDateCol.eachCell(function (cell, rowNumber) {
//         if (cell.value !== constants.STR_RESODATE &&
//             cell.value !== null) {
//                 sortedColumns.push(
//                     {"key": sheet.getColumn(constants.STR_KEY).values[rowNumber],
//                 "resolution": sheet.getColumn(constants.STR_RESODATE).values[rowNumber],
//                 "cycletime": sheet.getColumn(constants.STR_CYCLETIME).values[rowNumber],
//                 "leadtime": sheet.getColumn(constants.STR_LEADTIME).values[rowNumber]
//             });
//             indexResoDate = indexResoDate + 1;
//         }
//     });
//     return sortedColumns;
// }
const fillSortedColumn2 = (internalArray) => {
    let sortedColumns = [];

    internalArray.forEach( (value) => {
        if (value.resolutiondate !== constants.STR_RESODATE &&
            value.resolutiondate !== undefined) {
            sortedColumns.push(
                {
                    "key": value.key,
                    "resolution": value.resolutiondate,
                    "cycletime": value.cycletime,
                    "leadtime": value.leadtime
                });
        }
    });
    return sortedColumns;
}
const main = async () => {
    const { login, password } = getJIRAVariables();
    // axios.interceptors.request.use(request => {
        //     console.log('Starting Request', request)
        //     return request
        //   })
        
    const workbook = excel.initExcelFile(new ExcelJS.Workbook());

    
    try {
        // working locally to avoid calls to http
        const jsonIssues = { "data": JSON.parse(fs.readFileSync(path.join(__dirname,'mock/resIssues.json'), 'utf8'))}
        const jsonSprints = { "data": JSON.parse(fs.readFileSync(path.join(__dirname,'mock/resSprints.json'), 'utf8'))}
        const jsonSprintId = { "data": JSON.parse(fs.readFileSync(path.join(__dirname,'mock/resSprintId1772.json'), 'utf8'))}
        
        // const jsonIssues = await rest.getRequest(constants.JIRA_URL,login, password, rest.paramAxiosIssues);
        // const jsonSprints = await rest.getRequest(`${constants.JIRA_GREENHOPER_URL}/${constants.TDC_JIRA_BOARD_ID}`,login, password, rest.paramAxiosSprints);

        parseIssues(workbook,jsonIssues.data);
        const arrSprint = parseIdNameFromSprints(jsonSprints.data);
        let jsonSprintDetails = []
        let jsonSprintDetail = {}
        for(const value of arrSprint) {
            if(value.id !== undefined) {
                // jsonSprintDetail = await rest.getRequest(`${constants.JIRA_URL_SPRINT_BY_ID}/${value.id}`,login, password, {});
                jsonSprintDetail = jsonSprintId;
                jsonSprintDetails.push(
                    { 
                        "id" : value.id,
                        "name" : value.name,
                        "startDate" : jsonSprintDetail.data.startDate,
                        "endDate" : jsonSprintDetail.data.endDate,
                        "completeDate" : jsonSprintDetail.data.completeDate
                    });
            }
        };
        excel.fillExcelWithSprintsDetails(workbook,jsonSprintDetails);
        excel.groupRows(workbook);
        excel.writeExcelFile(workbook);
    } catch (error) {
        consoleError(error);
    }
}

main();