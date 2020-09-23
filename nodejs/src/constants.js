function define(name, value) {
    Object.defineProperty(exports, name, {
        value:      value,
        enumerable: true
    });
}

define("TDC_JIRA_BOARD_ID",217);
define("JIRA_SEARCH_URL","https://jira.talendforge.org/rest/api/2/search");
define("JIRA_URL_SPRINT_BY_ID","https://jira.talendforge.org/rest/agile/1.0/sprint");
define("JIRA_GREENHOPER_URL",`https://jira.talendforge.org/rest/greenhopper/1.0/sprintquery`);
define("JIRA_QUERY","project = TDC AND issuetype in (Bug, \"New Feature\", \"Work Item\") AND Sprint in (${value}) ORDER BY labels ASC, RANK");
define("JIRA_QUERY_SPRINTS","project = TDC AND sprint in (closedSprints(),futureSprints(),openSprints())");
define("STR_EXP_FILTER_SPRINT","TDC Sprint");
define("TDC_JIRA_SPRINT_PAGINATION",30);
define("TDC_JIRA_ISSUE_PAGINATION",100);

define("EXCEL_FILE_NAME","jira-report-js-full.xlsx");

define("STR_KEY","Issue key");
define("STR_TYPE","Issue Type");
define("STR_SUMMARY","Issue Summary");
define("STR_CREATIONDATE","Creation Date");
define("STR_RESODATE","Resolution Date");
define("STR_NEWDATE","In New Date");
define("STR_CANDIDATEDATE","In Candidate Date");
define("STR_ACCEPTDATE","In Accepted Date");
define("STR_PROGRESSDATE","In in progress Date");
define("STR_REVIEWDATE","In Code review Date");
define("STR_VALIDDATE","In Validation Date");
define("STR_MERGEDATE","In Merge Date");
define("STR_FINALCDATE","In Final check Date");
define("STR_DONEDATE","In Done Date");
define("STR_CLOSEDDATE","In Closed Date");
define("STR_ONHOLDDATE","In On hold Date");
define("STR_TODODATE","In To Do Date");
define("STR_BLOCKEDDATE","In Blocked Date");
define("STR_REJECTEDDATE","In Rejected Date");
define("STR_NEWTIME","In New Time");
define("STR_CANDIDATETIME","In Candidate Time");
define("STR_ACCEPTTIME","In Accepted Time");
define("STR_PROGRESSTIME","In in progress Time");
define("STR_REVIEWTIME","In Code review Time");
define("STR_VALIDTIME","In Validation Time");
define("STR_MERGETIME","In Merge Time");
define("STR_FINALCTIME","In Final check Time");
define("STR_DONETIME","In Done Time");
define("STR_CLOSEDTIME","In Closed Time");
define("STR_BLOCKEDTIME","In Blocked Time");
define("STR_REJECTEDTIME","In Rejected Time");
define("STR_ONHOLDTIME","In On hold Time");
define("STR_TODOTIME","In To Do Time");
define("STR_LEADTIME","Lead time");
define("STR_CYCLETIME","Cycle time");
define("STR_KEY_RESOLVED","Issue key resolved");
define("STR_RESOLUTION_DATE_RESOLVED","Resolution date for resolved");
define("STR_LEADTIME_RESOLVED","Lead time resolved");
define("STR_CYCLETIME_RESOLVED","Cycle time resolved");
define("STR_CENTILE_20TH_CYCLETIME","Centile 20th");
define("STR_CENTILE_50TH_CYCLETIME","Centile 50th");
define("STR_CENTILE_80TH_CYCLETIME","Centile 80th");
define("STR_CENTILE_50TH_LEADTIME","Centile 50th Lead Time");
define("STR_CENTILE_80TH_LEADTIME","Centile 80th Lead Time");
define("STR_CYCLETIMERANGE","Cycle time Range");
define("STR_CYCLETIMEDISTRIBUTION","Cycle time Distribution");
define("STR_LEADTIMERANGE","Lead time Range");
define("STR_LEADTIMEDISTRIBUTION","Lead time Distribution");

define("STR_SPRINT_ID","Sprint id");
define("STR_SPRINT_NAME","Sprint name");
define("STR_SPRINT_STARTDATE","Sprint start date");
define("STR_SPRINT_ENDDATE","Sprint end date");
define("STR_SPRINT_COMPLETEDATE","Sprint completed date");
define("STR_SPRINT_ENDMONTH","Sprint end month");
define("STR_SPRINT_ENDWEEK","Sprint end week");
define("STR_SPRINT_NBCOMPLETEDISSUES","Sprint completed issues");
define("STR_SPRINT_NBINCOMPLETEDISSUES","Sprint incompleted issues");

define("STR_GRP_RAWMETRICS","Raw metrics")
define("STR_GRP_RAWMETRICS_RESOLVED","Raw metrics resolved issues only")
define("STR_GRP_CYCLETIME_DISTRIBUTION","Cycle time distribution")
define("STR_GRP_LEADTIME_DISTRIBUTION","Lead time distribution")
define("STR_GRP_SPRINT","Sprint metrics")

define("WORKSHEET_NAME",'Raw_Metrics');

define("FILTER_LOW_CYCLETIME",0.05);
define("FILTER_HIGH_CYCLETIME",90);