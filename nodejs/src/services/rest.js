const axios = require('axios');
const constants = require('../constants')
const { printInfo,consoleError } = require('../print');


const getAllIssues = async (login, password) => {
    try {
        const response = await axios({
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
        });
        return(response.data);
    } catch (err) {
        consoleError(err);
    }
//     await axios({
//     method: 'get',
//     withCredentials: true,
//     headers: {
//         "Accept": "application/json",
//         "Content-Type": "application/json"
//     },
//     params: {
//         "expand":"changelog", // only accessible through get request and not in post jira api
//         "jql": constants.JIRA_QUERY,
//         "startAt":0,
//         "maxResults":500,
//     },
//     url: `${constants.JIRA_URL}`,
//     auth: {
//         username: login,
//         password: password
//     },
// }).then(function (response) {
//         return(response.data);
// })
// .catch(function (error) {
//     consoleError(error);
// });

};

module.exports = {
    getAllIssues,
}