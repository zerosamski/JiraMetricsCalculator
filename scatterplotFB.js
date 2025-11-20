// This code sample uses the 'node-fetch' library:
// https://www.npmjs.com/package/node-fetch
const fetch = require('node-fetch');
const reader = require('xlsx');

// Define variables
let values;
let changelogItem;
let object;
let objectIssues
let projectID = 'FBDDA';
let jiraEmailAndToken = 'samuel.mulkens@schiphol.nl:ATATT3xFfGF0YCCZYadK5shOsEHW2q55VRawSCLiUFaSUSttp29ld1O7cgNiJUIAVLMua-qkGOeDL917BsT3WuKtmB4FMik23Pe5aousPbdAiYTOlTlRwm_eyfuzmQC5vy8YGkDq23Pp0tKtda9eo4WN31ofSG5TN-KbuZbtyICk6tZILBWUSzk=93FB846B'
let itemList = [];
let csvList = [];

const delay = ms => new Promise(res => setTimeout(res, ms));
const getCsv = async () => {

    //Get List of Issues --> resolved %26resolved>="2023/04/01"
    fetch(`https://schiphol.atlassian.net/rest/api/3/search?maxResults=100&jql=project=${projectID}%26issueType=Story%26status='done'%26resolved>"2024/07/01"`, {
        method: 'GET',
        headers: {
            'Authorization': `Basic ${Buffer.from(
               jiraEmailAndToken
            ).toString('base64')}`,
            'Accept': 'application/json'
        }
    })
        .then(response => {
            return response.text();
        })
        .then(text => {
            objectIssues = JSON.parse(text)
            for (var issueItem of objectIssues.issues) {
                let epic = 'no epic'
                if (issueItem?.fields?.parent?.fields?.summary !== undefined) {
                    epic = issueItem.fields.parent.fields.summary.toString();
                }
                itemList.push([issueItem.key.toString(), issueItem.fields.issuetype.name.toString(), issueItem.fields.summary.toString(), epic.toString()])
            }
            getIssueData(itemList)
        })
        .catch(err => console.error(err));

    //Get Issue Data
    function getIssueData(itemList) {
        for (let item of itemList) {
            issueKey = item[0].split('-').pop()

            fetch(`https://schiphol.atlassian.net/rest/api/3/issue/${projectID}-${issueKey}/changelog`, {
                method: 'GET',
                headers: {
                    'Authorization': `Basic ${Buffer.from(
                        jiraEmailAndToken
                    ).toString('base64')}`,
                    'Accept': 'application/json'
                }
            })
                .then(response => {
                    return response.text();
                })
                .then(text => {
                    object = JSON.parse(text)
                    getCycleTime(object, item)
                })
                .catch(err => console.error(err));
        }
    }

    function getCycleTime(ctObject, item) {
        values = ctObject.values
        let cycleTime = 0;
        let startTime = 0;
        let finishTime = 0;

        //get in progress time and finish time for work item
        for (let inner of Object.entries(values)) {
            changelogItem = inner[1]
            
            if (startTime == 0 && changelogItem.items[0].field == 'status' && (changelogItem.items[0].fromString == 'New' || changelogItem.items[0].fromString == 'Backlog' || changelogItem.items[0].fromString == 'Analyse' || changelogItem.items[0].fromString == 'In Refinement' || changelogItem.items[0].fromString == 'Ready (For Sprint)' || changelogItem.items[0].fromString == 'Ready For Poker' || changelogItem.items[0].fromString == 'To Do') && (changelogItem.items[0].toString == 'In Progress')) {
                startTime = new Date(changelogItem.created);
            }
            
            //get finish time
            if (changelogItem.items[0].field == 'resolution' && changelogItem.items[0].toString == 'Done' && startTime != 0) {
                finishTime = new Date(changelogItem.created);
            }
        }

        //Calculate cycle time
        if (startTime != 0 && finishTime != 0) {
            cycleTime = Math.ceil((finishTime.getTime() - startTime.getTime()) / 86400000);
            item.push(finishTime.toLocaleDateString('en-US'), cycleTime.toString())
            csvList.push(item)
        }
    }
    //write CSV list to excel file
    console.log("Exporting data from JIRA and creating Excel file.")
    await delay(10000)

    const file = reader.readFile('./scatterplotFB.xlsx', { type: 'type', cellDates: true, dateNF: "dd/mm/yy" })
    const worksheet = file.Sheets[file.SheetNames[0]]

    reader.utils.sheet_add_aoa(worksheet, csvList, { origin: { r: 1, c: 0 } })
    formatColumn(worksheet, 5, "n")
    formatColumn(worksheet, 4, "d")
    reader.writeFile(file, "./scatterplotFB.xlsx")

    console.log("Finished!")
}
//fuction to format excel columns
function formatColumn(ws, col, fmt) {
    var range = reader.utils.decode_range(ws['!ref']);
    for (var R = range.s.r + 1; R <= range.e.r; ++R) {
        var cell_address = { c: col, r: R };
        var cell_ref = reader.utils.encode_cell(cell_address);
        if (ws[cell_ref]) {
            ws[cell_ref].t = fmt;
            if (fmt == "n") {
                ws[cell_ref].z = "0.00";
            }
        }
    }
}

getCsv();




