#!/usr/bin/env node
const { printError,printInfo,consoleError } = require('./print');
const axios = require('axios');
const ExcelJS = require('exceljs');
const percentile = require('just-percentile');
const constants = require('./constants')

const TDC_JIRA_BOARD_ID = 217
const TDC_JIRA_SPRINT_PAGINATION = 30

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
    const sheet = workbook.addWorksheet(constants.WORKSHEET_NAME, {views:[{state: 'frozen', xSplit: 1, ySplit:1}]});
    sheet.columns = [
        {header: constants.STR_KEY, key:constants.STR_KEY, width: '25'},
        {header: constants.STR_TYPE, key:constants.STR_TYPE, width: '25'},
        {header: constants.STR_SUMMARY, key:constants.STR_SUMMARY, width: '25'},
        {header: constants.STR_CREATIONDATE, key:constants.STR_CREATIONDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_RESODATE, key:constants.STR_RESODATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_NEWDATE, key:constants.STR_NEWDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_CANDIDATEDATE, key:constants.STR_CANDIDATEDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_ACCEPTDATE, key:constants.STR_ACCEPTDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_PROGRESSDATE, key:constants.STR_PROGRESSDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_REVIEWDATE, key:constants.STR_REVIEWDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_VALIDDATE, key:constants.STR_VALIDDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_MERGEDATE, key:constants.STR_MERGEDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_FINALCDATE, key:constants.STR_FINALCDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_DONEDATE, key:constants.STR_DONEDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_CLOSEDDATE, key:constants.STR_CLOSEDDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_ONHOLDDATE, key:constants.STR_ONHOLDDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_TODODATE, key:constants.STR_TODODATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_BLOCKEDDATE, key:constants.STR_BLOCKEDDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_REJECTEDDATE, key:constants.STR_REJECTEDDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_NEWTIME, key:constants.STR_NEWTIME, width: '25'},
        {header: constants.STR_CANDIDATETIME, key:constants.STR_CANDIDATETIME, width: '25'},
        {header: constants.STR_ACCEPTTIME, key:constants.STR_ACCEPTTIME, width: '25'},
        {header: constants.STR_PROGRESSTIME, key:constants.STR_PROGRESSTIME, width: '25'},
        {header: constants.STR_REVIEWTIME, key:constants.STR_REVIEWTIME, width: '25'},
        {header: constants.STR_VALIDTIME, key:constants.STR_VALIDTIME, width: '25'},
        {header: constants.STR_MERGETIME, key:constants.STR_MERGETIME, width: '25'},
        {header: constants.STR_FINALCTIME, key:constants.STR_FINALCTIME, width: '25'},
        {header: constants.STR_DONETIME, key:constants.STR_DONETIME, width: '25'},
        {header: constants.STR_CLOSEDTIME, key:constants.STR_CLOSEDTIME, width: '25'},
        {header: constants.STR_BLOCKEDTIME, key:constants.STR_BLOCKEDTIME, width: '25'},
        {header: constants.STR_REJECTEDTIME, key:constants.STR_REJECTEDTIME, width: '25'},
        {header: constants.STR_ONHOLDTIME, key:constants.STR_ONHOLDTIME, width: '25'},
        {header: constants.STR_TODOTIME, key:constants.STR_TODOTIME, width: '25'},
        {header: constants.STR_LEADTIME, key:constants.STR_LEADTIME, width: '25'},
        {header: constants.STR_CYCLETIME, key:constants.STR_CYCLETIME, width: '25'},
        {header: constants.STR_KEY_RESOLVED, key:constants.STR_KEY_RESOLVED, width: '25'},
        {header: constants.STR_RESOLUTION_DATE_RESOLVED, key:constants.STR_RESOLUTION_DATE_RESOLVED, width: '25'},
        {header: constants.STR_CYCLETIME_RESOLVED, key:constants.STR_CYCLETIME_RESOLVED, width: '25'},
        {header: constants.STR_CENTILE_20TH_CYCLETIME, key:constants.STR_CENTILE_20TH_CYCLETIME, width: '25'},
        {header: constants.STR_CENTILE_50TH_CYCLETIME, key:constants.STR_CENTILE_50TH_CYCLETIME, width: '25'},
        {header: constants.STR_CENTILE_80TH_CYCLETIME, key:constants.STR_CENTILE_80TH_CYCLETIME, width: '25'},
        {header: constants.STR_LEADTIME_RESOLVED, key:constants.STR_LEADTIME_RESOLVED, width: '25'},
        {header: constants.STR_CENTILE_50TH_LEADTIME, key:constants.STR_CENTILE_50TH_LEADTIME, width: '25'},
        {header: constants.STR_CENTILE_80TH_LEADTIME, key:constants.STR_CENTILE_80TH_LEADTIME, width: '25'},
        {header: constants.STR_CYCLETIMERANGE, key:constants.STR_CYCLETIMERANGE, width: '25'},
        {header: constants.STR_CYCLETIMEDISTRIBUTION, key:constants.STR_CYCLETIMEDISTRIBUTION, width: '25'},
        {header: constants.STR_LEADTIMERANGE, key:constants.STR_LEADTIMERANGE, width: '25'},
        {header: constants.STR_LEADTIMEDISTRIBUTION, key:constants.STR_LEADTIMEDISTRIBUTION, width: '25'},
    ];
    // Cross all issues retrieved from JIRA jql query
    for(let issueIdx in json.issues){
        const issue = json.issues[issueIdx];
        statusUpdate={};
        const lastRow = sheet.lastRow;
        const newRow = sheet.addRow(++(lastRow.number));
        newRow.getCell(constants.STR_KEY).value = issue.key;
        newRow.commit();
        printInfo(`Analysing ${newRow.getCell(constants.STR_KEY).value} item`)
        
        // # Get datetime creation
        const datetime_creation = issue.fields.created ? new Date(issue.fields.created) : undefined;
        if(datetime_creation !== undefined) {
            newRow.getCell(constants.STR_CREATIONDATE).value = datetime_creation.toLocaleString("fr-FR",{day:"numeric",month:"numeric",year:"numeric"});
            newRow.getCell(constants.STR_NEWDATE).value = datetime_creation.toLocaleString();
        }
        
        // # Get datetime resolution
        const datetime_resolution = issue.fields.resolutiondate ? new Date(issue.fields.resolutiondate) : undefined;
        if(datetime_resolution !== undefined) {
            newRow.getCell(constants.STR_RESODATE).value = datetime_resolution.toLocaleString("fr-FR",{day:"numeric",month:"numeric",year:"numeric"});
            const nbOfDays = (datetime_resolution.getTime()-datetime_creation.getTime()) / (1000*60*60*24)
            // newRow.getCell(constants.STR_LEADTIME).value = formatDateFromDays(nbOfDays);
            newRow.getCell(constants.STR_LEADTIME).value = nbOfDays;
        }
        
        newRow.getCell(constants.STR_TYPE).value = issue.fields.issuetype.name;
        newRow.getCell(constants.STR_SUMMARY).value = issue.fields.summary;
        
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
        newRow.getCell(constants.STR_CYCLETIME).value = 
        newRow.getCell(constants.STR_PROGRESSTIME).value+
        newRow.getCell(constants.STR_REVIEWTIME).value+
        newRow.getCell(constants.STR_VALIDTIME).value+
        newRow.getCell(constants.STR_MERGETIME).value+
        newRow.getCell(constants.STR_FINALCTIME).value;
        
    }//end parsing issues
    
    
    const freqCellColumnByKeyColumn = (step, cellColumn, columnKey ) => {
        let freq = {};
        const CTCol = sheet.getColumn(columnKey);
        
        CTCol.eachCell( (cell, rowNumber) => {
            if(cell.value !== columnKey && 
                cell.value !== 0 &&
                sheet.getColumn(cellColumn).values[rowNumber] !== undefined &&
                sheet.getColumn(cellColumn).values[rowNumber] !== cellColumn &&
                cell.value<constants.FILTER_HIGH_CYCLETIME &&
                cell.value>=constants.FILTER_LOW_CYCLETIME) {
                    let rangeIdx = Math.floor(cell/step);
                    if(rangeIdx in freq){
                        freq[rangeIdx] = freq[rangeIdx]+1;
                    } else {
                        freq[rangeIdx] = 1;
                    }
                }
            });
            return(freq);
        }
        
        //fill Distribution Cycle time
        var freqCT = {};
        const step = 3;
        freqCT = freqCellColumnByKeyColumn(step, constants.STR_RESODATE, constants.STR_CYCLETIME);
        
        const maxKey = Math.max(...Object.keys(freqCT));
        for(let steps = 0;steps<maxKey+2;steps++) {
            //jump first row <=> title
            currentRow = sheet.getRow(steps+2);
            currentRow.getCell(constants.STR_CYCLETIMERANGE).value = steps*step;
            currentRow.getCell(constants.STR_CYCLETIMEDISTRIBUTION).value = freqCT[steps] !== undefined ? freqCT[steps]:0;
        }
        //fill Distribution Lead time
        var freqLT = {};
        const stepLT = 5;
        freqLT = freqCellColumnByKeyColumn(stepLT, constants.STR_RESODATE, constants.STR_LEADTIME);
        
        const maxKeyLT = Math.max(...Object.keys(freqLT));
        for(let steps = 0;steps<maxKeyLT+2;steps++) {
            //jump first row <=> title
            currentRow = sheet.getRow(steps+2);
            currentRow.getCell(constants.STR_LEADTIMERANGE).value = steps*stepLT;
            currentRow.getCell(constants.STR_LEADTIMEDISTRIBUTION).value = freqLT[steps] !== undefined ? freqLT[steps]:0;
        }
        
        
        //fill Cycle time and lead time 
        //fill Resolved issue metrics
        const resDateCol = sheet.getColumn(constants.STR_RESODATE);
        let sortedColumns = []
        let indexResoDate = 2;
        resDateCol.eachCell( function(cell,rowNumber) {
            if( cell.value !== constants.STR_RESODATE && 
                cell.value !== null ) {
                    sortedColumns.push(
                        { 
                        "key":sheet.getColumn(constants.STR_KEY).values[rowNumber],
                        "resolution":sheet.getColumn(constants.STR_RESODATE).values[rowNumber],
                        "cycletime": sheet.getColumn(constants.STR_CYCLETIME).values[rowNumber],
                        "leadtime": sheet.getColumn(constants.STR_LEADTIME).values[rowNumber]
                    });
                    indexResoDate = indexResoDate + 1 ;
                }
            });
            
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
        
        //after filling all raw metrics, split with group row
        const fillArrayTitleRow = (grp) => {
            let grpFilled = [];
            grp.forEach((grpint,index) => {
                grpFilled[sheet.getColumn(grp[index].keyStart).number] = grp[index].title;
            });
            return(grpFilled)
        };
        
        const groupRow = [
            { "title":constants.STR_GRP_RAWMETRICS, "keyStart":constants.STR_KEY, "keyEnd":constants.STR_CYCLETIME, "color":"ccf2ff"},
            { "title":constants.STR_GRP_RAWMETRICS_RESOLVED, "keyStart":constants.STR_KEY_RESOLVED, "keyEnd":constants.STR_CENTILE_80TH_LEADTIME, "color":"f2d9e6"},
            { "title":constants.STR_GRP_CYCLETIME_DISTRIBUTION, "keyStart":constants.STR_CYCLETIMERANGE, "keyEnd":constants.STR_CYCLETIMEDISTRIBUTION, "color":"ccffcc"},
            { "title":constants.STR_GRP_LEADTIME_DISTRIBUTION, "keyStart":constants.STR_LEADTIMERANGE, "keyEnd":constants.STR_LEADTIMEDISTRIBUTION, "color":"ccaacc"},
        ];
        const getCellForStyle = (_sheet, _key, _rowNumber) => {
            return(_sheet.getRow(_rowNumber).getCell(_sheet.getColumn(_key).number));
        };

        let grpRowTitle = [];
        grpRowTitle = fillArrayTitleRow(groupRow);
        sheet.insertRow(1,grpRowTitle);
        
        let cellSelected = {};
        const centerMiddleStyle = { vertical: 'middle', horizontal: 'center' };

        groupRow.forEach((_,index) => {
            sheet.mergeCells(1,sheet.getColumn(groupRow[index].keyStart).number,1,sheet.getColumn(groupRow[index].keyEnd).number);
            cellSelected = getCellForStyle(sheet,groupRow[index].keyStart,1);
            cellSelected.alignment = centerMiddleStyle;
            cellSelected.fill = {
                type: 'pattern',
                pattern:'solid',
                fgColor:{argb:groupRow[index].color},
            };
        });

        // write to a file
        writeExcelFile(workbook);
    }

const writeExcelFile = async (workbook) => {
    try {
        await workbook.xlsx.writeFile(constants.EXCEL_FILE_NAME);
    }
    catch (error) {
        consoleError(error);
    }
}

const main = async () => {
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
            "jql": constants.JIRA_QUERY,
            "startAt":0,
            "maxResults":500,
        },
        url: `${constants.JIRA_URL}`,
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