const axios = require('axios');
const constants = require('../constants')

const paramAxiosIssues = (jql, startat, maxresults) =>
{    return({
                "params": {
                "expand":"changelog", // only accessible through get request and not in post jira api
                "jql": jql,
                "startAt":startat,
                "maxResults": maxresults,
                "fields":"created,resolutiondate,issuetype,summary,customfield_11070"
                }
    });
};
const paramAxiosSprints = 
{    
    "params": {
        "includeFutureSprints":true,
        "includeHistoricSprints":false
    }
};

const getAxiosConfig = (url, login, password, params ) => {
    return(
        {        
            "method": 'get',
            "withCredentials": true,
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            },
            "url": url,
            "auth": {
                "username": login,
                "password": password
            },
            ...params
        }
    )
}
const getRequest = async (url,login,password,params) => {
    return(await axios(getAxiosConfig(url,login, password, params)));
}
module.exports = {
    getAxiosConfig,
    getRequest,
    paramAxiosIssues,
    paramAxiosSprints,
}