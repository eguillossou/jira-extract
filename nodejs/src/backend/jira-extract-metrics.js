const express = require('express');
const fs = require('fs');
const path = require('path');

const excel = require('../services/excel')
const constants = require('../utils/constants')
const rest = require('../services/rest');
const { printError,printInfo,consoleError } = require('../utils/print');

const getJIRAVariables = () => {
    let { JIRA_LOGIN, JIRA_PASSWORD } = process.env;
    let LOGIN='';
    if (!JIRA_LOGIN) {
        LOGIN="eguillossou";
        printInfo(`CONFIG: JIRA_LOGIN environment variable set to default: `+LOGIN);
    } else {
        LOGIN=JIRA_LOGIN;
    }
    if (!JIRA_PASSWORD) {
        printError(`CONFIG ERROR: JIRA_PASSWORD environment variable not set.`);
    }

    return { login: LOGIN, password: decodeURIComponent(escape( Buffer.from(JIRA_PASSWORD, 'base64').toString() )) };
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

// return an array of frequency for each values for a specific range with a step
// filter also values between a min and a high boundaries
// eg: for cycle time values, get frequency of values between 0 and 3 (step of 3 here)
const freqCellColumnByKeyColumn = (step, internalArray, cellColumn, columnKey ) => {
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

const parseIssues = ( json ) => {
    const internalArray = [];
    
    // Cross all issues retrieved from JIRA jql query
    for(let issueIdx in json.issues){
        const issue = json.issues[issueIdx];
        statusUpdate={};
        const issueObject = {};

        issueObject.key = issue.key;
        
        // # Get datetime creation
        const datetime_creation = issue.fields.created ? new Date(issue.fields.created) : undefined;
        if(datetime_creation !== undefined) {
            issueObject.creationdate = datetime_creation.toLocaleString("fr-FR",{day:"numeric",month:"numeric",year:"numeric"});
            issueObject.newdate = datetime_creation.toLocaleString();
        }
        
        // # Get datetime resolution
        const datetime_resolution = issue.fields.resolutiondate ? new Date(issue.fields.resolutiondate) : undefined;
        if(datetime_resolution !== undefined) {
            const nbOfDays = (datetime_resolution.getTime()-datetime_creation.getTime()) / (1000*60*60*24)

            issueObject.resolutiondate = datetime_resolution.toLocaleString("fr-FR",{day:"numeric",month:"numeric",year:"numeric"});
            issueObject.leadtime = nbOfDays;
        }
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
                    
                    issueObject[`${switchDateOrTime(item.fromString)}`] = dateTransition.toLocaleString();
                    issueObject[`${switchDateOrTime(item.fromString,false)}`] = statusUpdate[item.fromString]/(1000*60*60*24);
                }
            }
        }
        // CYCLE Time series is a sum of in progress/code review/validation/merge/final check time
        issueObject.cycletime = 
        (issueObject[constants.STR_PROGRESSTIME] !== undefined ? issueObject[constants.STR_PROGRESSTIME]:0)+
        (issueObject[constants.STR_REVIEWTIME] !== undefined ? issueObject[constants.STR_REVIEWTIME]:0)+
        (issueObject[constants.STR_VALIDTIME] !== undefined ? issueObject[constants.STR_VALIDTIME]:0)+
        (issueObject[constants.STR_MERGETIME] !== undefined ? issueObject[constants.STR_MERGETIME]:0)+
        (issueObject[constants.STR_FINALCTIME] !== undefined ? issueObject[constants.STR_FINALCTIME]:0);
        
        issueObject.sprintlist = issue.fields.customfield_11070.map(value => value.match("name=[^,]+")[0].replace("name=",""));
        issueObject.storypoint = issue.fields.customfield_10150;
        
        internalArray.push(issueObject);
    }//end parsing issues
    return(internalArray);
}
// Getting last ten Sprint list with start and end date
// return last 10 sprints matching case "TDC Sprint XX" in closed
const parseIdNameFromSprints = (json) => {
    // getting first row to fill (remove group row and header row)
    const arrSprint = [];
    const filterSprints = json.sprints.filter((sprint) => 
    sprint.name.includes(constants.STR_EXP_FILTER_SPRINT) && sprint.state.includes("CLOSED"));
    const lastTenSprints = (filterSprints.filter((_, idx) => idx >filterSprints.length-11))
    .sort((a, b) => {
        if (a.name > b.name) {
          return 1;
        }
        if (a.name < b.name) {
          return -1;
        }
        return 0;
    });
    for(let sprintNb in lastTenSprints) {
        arrSprint[sprintNb] = {
            "id": lastTenSprints[sprintNb].id,
            "name":lastTenSprints[sprintNb].name
        };
    }
    return(arrSprint);
}
const calculateDistributionCycleTime = (internalArray) => {
    var freqCT = {};
    const distributionCycleTime = [];
    const step = 3;
    freqCT = freqCellColumnByKeyColumn(step,internalArray,constants.STR_RESODATE, constants.STR_CYCLETIME);
    
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
const calculateDistributionLeadTime = (internalArray) => {
    var freqLT = {};
    const distributionLeadTime = [];
    const stepLT = 5;
    freqLT = freqCellColumnByKeyColumn(stepLT,internalArray,constants.STR_RESODATE, constants.STR_LEADTIME);

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
const getCompleteAndUnCompleteIssueBySprint = (issueArray,jsonSprintDetails) => {
    const frequencyComplete = [];
    issueArray.forEach( issue => {
        jsonSprintDetails.forEach(sprint => {
            if(!frequencyComplete.find(val => val.id === sprint.id)) {
                frequencyComplete.push(
                    {
                        id: sprint.id,
                        completedissues : 0,
                        incompletedissues : 0,
                        unplannedissues : 0,
                        startedandcompletedissues: 0,
                        nonstartedandcompletedissues: 0,
                        plannedstorypoints: 0,
                        completedstorypoints: 0,
                        unestimatedissues: 0
                    });
            }
            if( issue.sprintlist.includes(sprint.name)) {
                const selectedSprint = frequencyComplete.find(v => v.id === sprint.id);
                if( issue.creationdate !== undefined &&
                    new Date(issue.creationdate)>= new Date(sprint.startdate)) {
                        selectedSprint.unplannedissues = selectedSprint.unplannedissues+1;
                }
                if( issue[constants.STR_PROGRESSDATE] !== undefined &&
                    issue.resolutiondate !== undefined &&
                    new Date(issue.resolutiondate)<= new Date(sprint.completedate)) {
                    if( new Date(issue[constants.STR_PROGRESSDATE])>= new Date(sprint.startdate)) {
                        selectedSprint.startedandcompletedissues = selectedSprint.startedandcompletedissues+1;
                    } else {
                        selectedSprint.nonstartedandcompletedissues = selectedSprint.nonstartedandcompletedissues+1;
                    }
                }
                if( issue.resolutiondate !== undefined &&
                new Date(issue.resolutiondate) <= new Date(sprint.completedate)) {
                    selectedSprint.completedissues = selectedSprint.completedissues +1;
                    if(issue.storypoint !== null) {
                        selectedSprint.completedstorypoints = selectedSprint.completedstorypoints +issue.storypoint;
                    }
                } else {
                    selectedSprint.incompletedissues = selectedSprint.incompletedissues +1;
                }
                if(issue.storypoint !== null) {
                    selectedSprint.plannedstorypoints = selectedSprint.plannedstorypoints +issue.storypoint;
                } else {
                    selectedSprint.unestimatedissues = selectedSprint.unestimatedissues+1;
                }
            }
        })
    });
    return(frequencyComplete);
}
const main = async () => {
    const app = express();
    const { login, password } = getJIRAVariables();
    
    try {
        let jsonSprints = {};
        // MOCK INSIDE
        let isMock = false;
        // working locally to avoid calls to http
        if(isMock) {
            jsonSprints = { "data": JSON.parse(fs.readFileSync(path.join(__dirname,'../mock/resSprints.json'), 'utf8'))}
        } else {
            printInfo(`Getting last ten Sprint list ${new Date().toLocaleString()}`);
            jsonSprints = await rest.getRequest(`${constants.JIRA_GREENHOPER_URL}/${constants.TDC_JIRA_BOARD_ID}`,login, password, rest.paramAxiosSprints);
        }
        
        const arrSprint = parseIdNameFromSprints(jsonSprints.data);
        let jsonSprintDetails = [];
        let jsonSprintDetail = {};
        
        for(const value of arrSprint) {
            if(value.id !== undefined) {
                // MOCK INSIDE
                // working locally to avoid calls to http
                if(isMock) {
                    jsonSprintDetail = { "data": JSON.parse(fs.readFileSync(path.join(__dirname,'../mock/resSprintIds.json'), 'utf8'))};
                } else {
                    jsonSprintDetail = await rest.getRequest(`${constants.JIRA_URL_SPRINT_BY_ID}/${value.id}`,login, password, {});
                }
                
                jsonSprintDetails.push(
                { 
                    "id" : value.id,
                    "name" : value.name,
                    "startdate" : jsonSprintDetail.data.startDate,
                    "enddate" : jsonSprintDetail.data.endDate,
                    "completedate" : jsonSprintDetail.data.completeDate
                });
            }
        };
        
        let jsonSprintIssues = {};
        if(isMock) {
             jsonSprintIssues = { "data": JSON.parse(fs.readFileSync(path.join(__dirname,'../mock/resIssues.json'), 'utf8'))}
        } else {
            jql = constants.JIRA_QUERY.replace("${value}",`${jsonSprintDetails.map(v => v.id).join(', ')}`);
            printInfo(`Start searching issues ${new Date().toLocaleString()} with pagination`);
            let startat = 0;
            jsonSprintIssues = await rest.getRequest(constants.JIRA_SEARCH_URL, login, password, rest.paramAxiosIssues(jql, startat, constants.TDC_JIRA_ISSUE_PAGINATION));
            const listFn = [];
            for(let i = 1;i<=Math.floor(jsonSprintIssues.data.total/constants.TDC_JIRA_ISSUE_PAGINATION);i++) {
                listFn.push(rest.getRequest(constants.JIRA_SEARCH_URL, login, password, rest.paramAxiosIssues(jql, constants.TDC_JIRA_ISSUE_PAGINATION*i, constants.TDC_JIRA_ISSUE_PAGINATION)));
            }
            const issuesPaginated = await Promise.all(listFn.map(fn => fn));
            for (const response in issuesPaginated) {
                jsonSprintIssues.data.issues = [...jsonSprintIssues.data.issues,...issuesPaginated[response].data.issues];
            }
            printInfo(`End searching issues ${new Date().toLocaleString()}`);
        }
        let issueArray = parseIssues(jsonSprintIssues.data);

        //fill raw issues to Excel file
        excel.fileExcelWithRawIssues(issueArray);

        const filterSprints = (item) => {
            return(item.sprintlist.includes(jsonSprintDetails[jsonSprintDetails.length-1].name) ||
            item.sprintlist.includes(jsonSprintDetails[jsonSprintDetails.length-2].name) || 
            item.sprintlist.includes(jsonSprintDetails[jsonSprintDetails.length-3].name));
        };
        const issuesFromThreeLastSprintOnly = issueArray.filter((item) => filterSprints(item));
        //fill Distribution Cycle time in Excel file
        excel.fillExcelWithCyleTimeDistribution(calculateDistributionCycleTime(issuesFromThreeLastSprintOnly));
        //fill Distribution Lead time in Excel file
        excel.fillExcelWithLeadTimeDistribution(calculateDistributionLeadTime(issuesFromThreeLastSprintOnly));
        //fill Resolved issue metrics in Excel file
        excel.fillExcelWithResolvedIssuesOnly(issuesFromThreeLastSprintOnly);

        let sprintCompleteAndInComplete = getCompleteAndUnCompleteIssueBySprint(issueArray,jsonSprintDetails);

        const sprintCompleteAndInCompleteUpdated = sprintCompleteAndInComplete.map(({id, ...good}) => good );
        const jsonSprintDetailsUpdate  = jsonSprintDetails.map((item,index) => item = {...item, ...sprintCompleteAndInCompleteUpdated[index]})
        excel.fillExcelWithSprintsDetails(jsonSprintDetailsUpdate);
        excel.groupRows();
        excel.writeExcelFile();

        app.get('/', (req, res) => {
            res.send('Hello World');
        });
        app.get('/api/sprints', (req, res) => {
            res.send(jsonSprintDetailsUpdate);
        });
    } catch (error) {
        consoleError(error);
    }
    
    if (!module.parent) {
    //app.listen(3000);
    //console.log('Express started on port 3000');
    }
}

main();