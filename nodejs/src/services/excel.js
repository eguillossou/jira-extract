const constants = require('../utils/constants')
const percentile = require('just-percentile');
const { printInfo,consoleError } = require('../utils/print');
const ExcelJS = require('exceljs');

const workbook = new ExcelJS.Workbook();

const initExcelFile = () => {
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
        {header: constants.STR_SPRINT_ID, key:constants.STR_SPRINT_ID, width: '25'},
        {header: constants.STR_SPRINT_NAME, key:constants.STR_SPRINT_NAME, width: '25'},
        {header: constants.STR_SPRINT_STARTDATE, key:constants.STR_SPRINT_STARTDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_SPRINT_ENDDATE, key:constants.STR_SPRINT_ENDDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_SPRINT_COMPLETEDATE, key:constants.STR_SPRINT_COMPLETEDATE, width: '25', style: { numFmt: 'dd/mm/yyyy  HH:mm:ss' }},
        {header: constants.STR_SPRINT_NBCOMPLETEDISSUES, key:constants.STR_SPRINT_NBCOMPLETEDISSUES, width: '25'},
        {header: constants.STR_SPRINT_NBINCOMPLETEDISSUES, key:constants.STR_SPRINT_NBINCOMPLETEDISSUES, width: '25'},
        {header: constants.STR_SPRINT_RATIOCOMPLETEDISSUES, key:constants.STR_SPRINT_RATIOCOMPLETEDISSUES, width: '25', style: { numFmt: '0%' }},
        {header: constants.STR_SPRINT_UNPLANNED_ISSUES, key:constants.STR_SPRINT_UNPLANNED_ISSUES, width: '25'},
        {header: constants.STR_SPRINT_STARTED_AND_COMPLETED_ISSUES, key:constants.STR_SPRINT_STARTED_AND_COMPLETED_ISSUES, width: '30'},
        {header: constants.STR_SPRINT_NON_STARTED_AND_COMPLETED_ISSUES, key:constants.STR_SPRINT_NON_STARTED_AND_COMPLETED_ISSUES, width: '30'},
        {header: constants.STR_SPRINT_UNESTIMATED_ITEMS, key:constants.STR_SPRINT_UNESTIMATED_ITEMS, width: '25'},
        {header: constants.STR_SPRINT_PLANNED_STORYPOINTS, key:constants.STR_SPRINT_PLANNED_STORYPOINTS, width: '25'},
        {header: constants.STR_SPRINT_COMPLETED_STORYPOINTS, key:constants.STR_SPRINT_COMPLETED_STORYPOINTS, width: '25'},
    ];
}
const fillRowValueInExcel = (rowObject, column, value ) => {
    rowObject.getCell(column).value = value;
}
const fileExcelWithRawIssues = (issueArray) => {
    const sheet = workbook.getWorksheet(constants.WORKSHEET_NAME);

    issueArray.forEach((value,index) => {
        const lastRow = sheet.lastRow;
        const newRow = sheet.addRow(++(lastRow.number));
        newRow.getCell(constants.STR_KEY).value = value.key;
        newRow.commit();
        printInfo(`Analysing ${newRow.getCell(constants.STR_KEY).value} item`);
        newRow.getCell(constants.STR_CREATIONDATE).value = value.creationdate;
        newRow.getCell(constants.STR_NEWDATE).value = value.newdate;
        newRow.getCell(constants.STR_RESODATE).value = value.resolutiondate; 
        newRow.getCell(constants.STR_LEADTIME).value = value.leadtime;
        newRow.getCell(constants.STR_TYPE).value = value.type;
        newRow.getCell(constants.STR_SUMMARY).value = value.summary;
        newRow.getCell(constants.STR_CYCLETIME).value = value.cycletime;
        newRow.getCell(constants.STR_TODODATE).value = value[`${constants.STR_TODODATE}`];
        newRow.getCell(constants.STR_TODOTIME).value = value[`${constants.STR_TODOTIME}`];
        newRow.getCell(constants.STR_NEWDATE).value = value[`${constants.STR_NEWDATE}`];
        newRow.getCell(constants.STR_NEWTIME).value = value[`${constants.STR_NEWTIME}`];
        newRow.getCell(constants.STR_CANDIDATEDATE).value = value[`${constants.STR_CANDIDATEDATE}`];
        newRow.getCell(constants.STR_CANDIDATETIME).value = value[`${constants.STR_CANDIDATETIME}`];
        newRow.getCell(constants.STR_ACCEPTDATE).value = value[`${constants.STR_ACCEPTDATE}`];
        newRow.getCell(constants.STR_ACCEPTTIME).value = value[`${constants.STR_ACCEPTTIME}`];
        newRow.getCell(constants.STR_PROGRESSDATE).value = value[`${constants.STR_PROGRESSDATE}`];
        newRow.getCell(constants.STR_PROGRESSTIME).value = value[`${constants.STR_PROGRESSTIME}`];
        newRow.getCell(constants.STR_REVIEWDATE).value = value[`${constants.STR_REVIEWDATE}`];
        newRow.getCell(constants.STR_REVIEWTIME).value = value[`${constants.STR_REVIEWTIME}`];
        newRow.getCell(constants.STR_VALIDDATE).value = value[`${constants.STR_VALIDDATE}`];
        newRow.getCell(constants.STR_VALIDTIME).value = value[`${constants.STR_VALIDTIME}`];
        newRow.getCell(constants.STR_MERGEDATE).value = value[`${constants.STR_MERGEDATE}`];
        newRow.getCell(constants.STR_MERGETIME).value = value[`${constants.STR_MERGETIME}`];
        newRow.getCell(constants.STR_FINALCDATE).value = value[`${constants.STR_FINALCDATE}`];
        newRow.getCell(constants.STR_FINALCTIME).value = value[`${constants.STR_FINALCTIME}`];
        newRow.getCell(constants.STR_DONEDATE).value = value[`${constants.STR_DONEDATE}`];
        newRow.getCell(constants.STR_DONETIME).value = value[`${constants.STR_DONETIME}`];
        newRow.getCell(constants.STR_CLOSEDDATE).value = value[`${constants.STR_CLOSEDDATE}`];
        newRow.getCell(constants.STR_CLOSEDTIME).value = value[`${constants.STR_CLOSEDTIME}`];
        newRow.getCell(constants.STR_BLOCKEDDATE).value = value[`${constants.STR_BLOCKEDDATE}`];
        newRow.getCell(constants.STR_BLOCKEDTIME).value = value[`${constants.STR_BLOCKEDTIME}`];
        newRow.getCell(constants.STR_REJECTEDDATE).value = value[`${constants.STR_REJECTEDDATE}`];
        newRow.getCell(constants.STR_REJECTEDTIME).value = value[`${constants.STR_REJECTEDTIME}`];
        newRow.getCell(constants.STR_ONHOLDDATE).value = value[`${constants.STR_ONHOLDDATE}`];
        newRow.getCell(constants.STR_ONHOLDTIME).value = value[`${constants.STR_ONHOLDTIME}`];
    });
}
const fillExcelWithCyleTimeDistribution = (distributionArray) => {
    const sheet = workbook.getWorksheet(constants.WORKSHEET_NAME);
    let index = 2;
    distributionArray.forEach(
        value => {
            currentRow = sheet.getRow(index);
            currentRow.getCell(constants.STR_CYCLETIMERANGE).value = value.cycletimerange;
            currentRow.getCell(constants.STR_CYCLETIMEDISTRIBUTION).value = value.cycletimedistribution;
            index = index + 1;
        }
    );
}
const fillExcelWithLeadTimeDistribution = (distributionArray) => {
    const sheet = workbook.getWorksheet(constants.WORKSHEET_NAME);
    index = 2;
    distributionArray.forEach(
        value => {
            currentRow = sheet.getRow(index);
            currentRow.getCell(constants.STR_LEADTIMERANGE).value = value.leadtimerange;
            currentRow.getCell(constants.STR_LEADTIMEDISTRIBUTION).value = value.leadtimedistribution;    
            index = index + 1; 
        }
    );
}
const fillExcelWithResolvedIssuesOnly = (issueArray) => {
    const sheet = workbook.getWorksheet(constants.WORKSHEET_NAME);
    let filteredColumn = issueArray.filter(value => {
        return(
        value.cycletime > constants.FILTER_LOW_CYCLETIME && 
        value.cycletime <= constants.FILTER_HIGH_CYCLETIME &&
        value.resolutiondate !== undefined)
    })
    .sort((a,b) => new Date(a.resolutiondate).getTime() - new Date(b.resolutiondate).getTime());

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
        currentRow.getCell(constants.STR_RESOLUTION_DATE_RESOLVED).value = value.resolutiondate;
        currentRow.getCell(constants.STR_CYCLETIME_RESOLVED).value = Number((Math.round(value.cycletime * 100)/100).toFixed(2));
        currentRow.getCell(constants.STR_LEADTIME_RESOLVED).value = Number((Math.round(value.leadtime * 100)/100).toFixed(2));
        currentRow.getCell(constants.STR_CENTILE_20TH_CYCLETIME).value = centileThCycleTime(20);
        currentRow.getCell(constants.STR_CENTILE_50TH_CYCLETIME).value = centileThCycleTime(50);
        currentRow.getCell(constants.STR_CENTILE_80TH_CYCLETIME).value = centileThCycleTime(80);
        currentRow.getCell(constants.STR_CENTILE_50TH_LEADTIME).value = centileThLeadTime(50);
        currentRow.getCell(constants.STR_CENTILE_80TH_LEADTIME).value = centileThLeadTime(80);
    });

}
const fillExcelWithSprintsDetails = (sprintDetails) => {
    const sheet = workbook.getWorksheet(constants.WORKSHEET_NAME);
    //skipping titles : index 2
    let indexRow = 2;
    for(let details of sprintDetails) {
        let rowObject = sheet.getRow(indexRow);
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_ID, details.id );
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_NAME, details.name );
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_STARTDATE, new Date(details.startdate).toLocaleString("fr-FR",{day:"numeric",month:"numeric",year:"numeric",hour:"numeric", minute:"numeric"}) );
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_ENDDATE, new Date(details.enddate).toLocaleString("fr-FR",{day:"numeric",month:"numeric",year:"numeric",hour:"numeric", minute:"numeric"}) );
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_COMPLETEDATE, new Date(details.completedate).toLocaleString("fr-FR",{day:"numeric",month:"numeric",year:"numeric",hour:"numeric", minute:"numeric"}) );
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_NBCOMPLETEDISSUES, details.completedissues );
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_NBINCOMPLETEDISSUES, details.incompletedissues );
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_RATIOCOMPLETEDISSUES, details.incompletedissues+details.completedissues !== 0 ?(details.completedissues/(details.incompletedissues+details.completedissues) ):1 );
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_UNPLANNED_ISSUES, details.unplannedissues );
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_STARTED_AND_COMPLETED_ISSUES, details.startedandcompletedissues );
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_NON_STARTED_AND_COMPLETED_ISSUES, details.nonstartedandcompletedissues );
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_UNESTIMATED_ITEMS, details.unestimatedissues );
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_PLANNED_STORYPOINTS, details.plannedstorypoints );
        fillRowValueInExcel(rowObject, constants.STR_SPRINT_COMPLETED_STORYPOINTS, details.completedstorypoints );
        indexRow = indexRow + 1;
    }
}
const groupRows = () => {
    const sheet = workbook.getWorksheet(constants.WORKSHEET_NAME);
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
        { "title":constants.STR_GRP_SPRINT, "keyStart":constants.STR_SPRINT_ID, "keyEnd":constants.STR_SPRINT_COMPLETED_STORYPOINTS, "color":"ccaaff"},
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
}
const writeExcelFile = async () => {
    try {
        await workbook.xlsx.writeFile(constants.EXCEL_FILE_NAME);
    }
    catch (error) {
        consoleError(error);
    }
}

initExcelFile();

module.exports = {
    initExcelFile,
    fileExcelWithRawIssues,
    fillExcelWithSprintsDetails,
    fillRowValueInExcel,
    fillExcelWithCyleTimeDistribution,
    fillExcelWithLeadTimeDistribution,
    fillExcelWithResolvedIssuesOnly,
    groupRows,
    writeExcelFile
}