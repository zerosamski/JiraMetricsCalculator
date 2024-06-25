//Libraries
const fetch = require('node-fetch');
const reader = require('xlsx');

// Define variables
let values;
let changelogItem;
let object;
let objectIssues;
let projectID = '';
let jiraEmailAndToken = 'email:token'
let itemList = [];
let csvList = [];

/**
 * 
 */
const delay = ms => new Promise(res => setTimeout(res, ms));
const getCsv = async () => {
    let totalIssues = 0;
    let startAt = 0;
    let issueCount = 0;

    console.log("Fetching 1st batch")
    getIssues();

    /**
     * This function uses the Jira API to create a list of issues that were resolved after a certain date. ProjectID is configurable in the above variables.
     * The API call fetches 100 issues at max (default is 50). If more remain, it recursively calls itself again untill all issues have been fetched.  
     */
    function getIssues() {
        //Get List of Issues
        fetch(`https://schiphol.atlassian.net/rest/api/3/search?maxResults=100&jql=project=${projectID}%26issueType!="Sub-task"%26status='done'%26resolved>="2024/04/01"&startAt=${startAt}`, {
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
                if (issueCount == 0) {
                    issueCount = objectIssues.total
                }
                totalIssues = objectIssues.total
                startAt += objectIssues.issues.length;
                for (var issueItem of objectIssues.issues) {
                    let component = 'no component'
                    issueCount--
                    if (issueItem?.fields?.components !== undefined) {
                        component = issueItem.fields.components[0].name.toString();
                    }
                    if (issueItem.fields.issuetype.name.toString() != "Epic") {
                        itemList.push([issueItem.key.toString(), issueItem.fields.issuetype.name.toString(), issueItem.fields.summary.toString(), component.toString()],)
                    }
                }
                if (issueCount == 0) {
                    console.log("Finished with fetching all issues")
                    getIssueData(itemList)
                } else {
                    console.log("Fetching next batch. Remaining issues: " + issueCount)
                    getIssues()
                }
            })
            .catch(err => console.error(err));
    }

    /**
     * This function interates over the above created list of issues and uses the Jira API to get the detailed information for each issue. 
     */
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
                    getTimeStamps(object, item)
                })
                .catch(err => console.error(err));
        }
    }

    /** This function gets the timestamps for all the status transitions. It then calculates lead time(Backlog to Done), process lead time(commitment point to Done), and cycle time (In Progress to Done).
     * Calculations are based on transition time, so for example the time an item spent in 'In Progress' is calculated by substracting the date the item was moved to 'In Progress' from the time the item was moved to 'Test'. 
     * If transition times are not available (because an item skipped a column) the calculations is done by adding the times the item spent in relevant columns. 
     * For example: process lead time is calculated as 'finish date' - 'to do date'. If 'to do date' is not available because that column was skipped, it is calcuated by adding the time spent in all columns between 'To Do' and 'Done'  
     **/
    function getTimeStamps(ct, item) {
        values = ct.values
        let newTime = 0;
        let backlogTime = 0;
        let inRefinementTime = 0;
        let readyForSprintTime = 0;
        let toDoTime = 0;
        let toDoTimeDate = 0;
        let analyseTime = 0;
        let inProgressTime = 0;
        let inProgressTimeDate = 0;
        let testTime = 0;
        let inReviewTime = 0;
        let doneTime = 0;
        let processLeadTime = 0;
        let cycleTime = 0;
        let resolution = "";

        /**
         * This section of the function finds all the timestamps for status changes. 
         */
        for (let inner of Object.entries(values)) {
            changelogItem = inner[1]

            //get time that ticket was created 
            if (newTime == 0) {
                newTime = new Date(changelogItem.created);
            }

            //get time that ticket was moved to 'In Refinement' (In Refinement)
            if (changelogItem.items[0].field == 'status' && changelogItem.items[0].toString == 'In Refinement' && changelogItem.items[0].fromString == 'Backlog') {
                inRefinementTime = new Date(changelogItem.created);
            }

            //get time that ticket was moved to 'Ready for Sprint' (Refined)
            if (changelogItem.items[0].field == 'status' && changelogItem.items[0].toString == 'Ready for Sprint' && (changelogItem.items[0].fromString == 'Backlog' || changelogItem.items[0].fromString == 'In Refinement')) {
                readyForSprintTime = new Date(changelogItem.created);
            }

            //get time that ticket was moved to 'To Do'(To Do)
            if (changelogItem.items[0].field == 'status' && changelogItem.items[0].toString == 'To Do' && (changelogItem.items[0].fromString == 'Backlog' || changelogItem.items[0].fromString == 'In Refinement' || changelogItem.items[0].fromString == 'Ready for Sprint')) {
                toDoTime = toDoTimeDate = new Date(changelogItem.created)
                toDoTimeDate = new Date(changelogItem.created)
            }

            //get time that ticket was moved to 'Analyse' (Analyse)
            if (changelogItem.items[0].field == 'status' && changelogItem.items[0].toString == 'Analyse' && (changelogItem.items[0].fromString == 'Backlog' || changelogItem.items[0].fromString == 'In Refinement' || changelogItem.items[0].fromString == 'Ready for Sprint' || changelogItem.items[0].fromString == 'To Do')) {
                analyseTime = new Date(changelogItem.created)
            }

            //get time that ticket was moved to 'In progress' (Realisatie)
            if (changelogItem.items[0].field == 'status' && changelogItem.items[0].toString == 'In Progress' && (changelogItem.items[0].fromString == 'Backlog' || changelogItem.items[0].fromString == 'In Refinement' || changelogItem.items[0].fromString == 'Ready for Sprint' || changelogItem.items[0].fromString == 'To Do' || changelogItem.items[0].toString == 'Analyse')) {
                inProgressTime = new Date(changelogItem.created)
                inProgressTimeDate = new Date(changelogItem.created)
            }

            //get time that ticket was moved to 'Test'' (Testen)
            if (changelogItem.items[0].field == 'status' && changelogItem.items[0].toString == 'Test' && (changelogItem.items[0].fromString == 'Backlog' || changelogItem.items[0].fromString == 'In Refinement' || changelogItem.items[0].fromString == 'Ready for Sprint' || changelogItem.items[0].fromString == 'To Do' || changelogItem.items[0].fromString == 'Analyse' || changelogItem.items[0].fromString == 'In Progress')) {
                testTime = new Date(changelogItem.created)
            }

            //get time that ticket was moved to 'In Reivew (Validatie klant)
            if (changelogItem.items[0].field == 'status' && changelogItem.items[0].toString == 'In Review' && (changelogItem.items[0].fromString == 'Backlog' || changelogItem.items[0].fromString == 'In Refinement' || changelogItem.items[0].fromString == 'Ready for Sprint' || changelogItem.items[0].fromString == 'To Do' || changelogItem.items[0].fromString == 'Analyse' || changelogItem.items[0].fromString == 'In Progress' || changelogItem.items[0].fromString == 'Test')) {
                inReviewTime = new Date(changelogItem.created)
            }

            //get finish time
            if (changelogItem.items[0].field == 'resolution' && changelogItem.items[0].toString == "Done" && newTime != 0) {
                doneTime = new Date(changelogItem.created);
                resolution = changelogItem.items[0].toString
        }
    }

        /**
         * This section calculates the time an issue spent in each status, cycle time, process lead time, and lead time. 
         */

        if (newTime != 0 && doneTime != 0) {

            //get time in 'Backlog'
            if (inRefinementTime != 0) {
                backlogTime = Math.floor((inRefinementTime.getTime() - newTime.getTime()) / 86400000);
            } else if (readyForSprintTime != 0) {
                backlogTime = Math.floor((readyForSprintTime.getTime() - newTime.getTime()) / 86400000);
            } else if (toDoTime != 0) {
                backlogTime = Math.floor((toDoTime.getTime() - newTime.getTime()) / 86400000);
            } else if (analyseTime != 0) {
                backlogTime = Math.floor((analyseTime.getTime() - newTime.getTime()) / 86400000);
            } else if (inProgressTime != 0) {
                backlogTime = Math.floor((inProgressTime.getTime() - newTime.getTime()) / 86400000);
            } else if (testTime != 0) {
                backlogTime = Math.floor((testTime.getTime() - newTime.getTime()) / 86400000);
            } else if (inReviewTime != 0) {
                backlogTime = Math.floor((inReviewTime.getTime() - newTime.getTime()) / 86400000);
            } else if (doneTime != 0) {
                backlogTime = Math.floor((doneTime.getTime() - newTime.getTime()) / 86400000);
            }

            //get time in 'In Refinement'
            if (inRefinementTime != 0) {
                if (readyForSprintTime != 0) {
                    inRefinementTime = Math.ceil((readyForSprintTime.getTime() - inRefinementTime.getTime()) / 86400000);
                } else if (toDoTime != 0) {
                    inRefinementTime = Math.ceil((toDoTime.getTime() - inRefinementTime.getTime()) / 86400000);
                } else if (analyseTime != 0) {
                    inRefinementTime = Math.ceil((analyseTime.getTime() - inRefinementTime.getTime()) / 86400000);
                } else if (inProgressTime != 0) {
                    inRefinementTime = Math.ceil((inProgressTime.getTime() - inRefinementTime.getTime()) / 86400000);
                } else if (testTime != 0) {
                    inRefinementTime = Math.ceil((testTime.getTime() - inRefinementTime.getTime()) / 86400000);
                } else if (inReviewTime != 0) {
                    inRefinementTime = Math.ceil((inReviewTime.getTime() - inRefinementTime.getTime()) / 86400000);
                } else if (doneTime != 0) {
                    inRefinementTime = Math.ceil((doneTime.getTime() - inRefinementTime.getTime()) / 86400000);
                }
            }

            //get time in 'Ready for Sprint'
            if (readyForSprintTime != 0) {
                if (toDoTime != 0) {
                    readyForSprintTime = Math.ceil((toDoTime.getTime() - readyForSprintTime.getTime()) / 86400000);
                } else if (analyseTime != 0) {
                    readyForSprintTime = Math.ceil((analyseTime.getTime() - readyForSprintTime.getTime()) / 86400000);
                } else if (inProgressTime != 0) {
                    readyForSprintTime = Math.ceil((inProgressTime.getTime() - readyForSprintTime.getTime()) / 86400000);
                } else if (testTime != 0) {
                    readyForSprintTime = Math.ceil((testTime.getTime() - readyForSprintTime.getTime()) / 86400000);
                } else if (inReviewTime != 0) {
                    readyForSprintTime = Math.ceil((inReviewTime.getTime() - readyForSprintTime.getTime()) / 86400000);
                } else if (doneTime != 0) {
                    readyForSprintTime = Math.ceil((doneTime.getTime() - readyForSprintTime.getTime()) / 86400000);
                }
            }

            //get time in 'To Do'
            if (toDoTime != 0) {
                if (analyseTime != 0) {
                    toDoTime = Math.ceil((analyseTime.getTime() - toDoTime.getTime()) / 86400000);
                } else if (inProgressTime != 0) {
                    toDoTime = Math.ceil((inProgressTime.getTime() - toDoTime.getTime()) / 86400000);
                } else if (testTime != 0) {
                    toDoTime = Math.ceil((testTime.getTime() - toDoTime.getTime()) / 86400000);
                } else if (inReviewTime != 0) {
                    toDoTime = Math.ceil((inReviewTime.getTime() - toDoTime.getTime()) / 86400000);
                } else if (doneTime != 0) {
                    toDoTime = Math.ceil((doneTime.getTime() - toDoTime.getTime()) / 86400000);
                }
            }

            //get time in 'Analyse'
            if (analyseTime != 0) {
                if (inProgressTime != 0) {
                    analyseTime = Math.ceil((inProgressTime.getTime() - analyseTime.getTime()) / 86400000);
                } else if (testTime != 0) {
                    analyseTime = Math.ceil((testTime.getTime() - analyseTime.getTime()) / 86400000);
                } else if (inReviewTime != 0) {
                    analyseTime = Math.ceil((inReviewTime.getTime() - analyseTime.getTime()) / 86400000);
                } else if (doneTime != 0) {
                    analyseTime = Math.ceil((doneTime.getTime() - analyseTime.getTime()) / 86400000);
                }
            }
            //get time in 'In Progress'
            if (inProgressTime != 0) {
                if (testTime != 0) {
                    inProgressTime = Math.ceil((testTime.getTime() - inProgressTime.getTime()) / 86400000);
                } else if (inReviewTime != 0) {
                    inProgressTime = Math.ceil((inReviewTime.getTime() - inProgressTime.getTime()) / 86400000);
                } else if (doneTime != 0) {
                    inProgressTime = Math.ceil((doneTime.getTime() - inProgressTime.getTime()) / 86400000);
                }
            }

            //get time in 'Test''
            if (testTime != 0) {
                if (inReviewTime != 0) {
                    testTime = Math.ceil((inReviewTime.getTime() - testTime.getTime()) / 86400000);
                } else if (doneTime != 0) {
                    testTime = Math.ceil((doneTime.getTime() - testTime.getTime()) / 86400000);
                }
            }

            //get time in 'In Review'
            if (inReviewTime != 0) {
                if (doneTime != 0) {
                    inReviewTime = Math.ceil((doneTime.getTime() - inReviewTime.getTime()) / 86400000);
                }
            }

            let leadTime = Math.ceil((doneTime.getTime() - newTime.getTime()) / 86400000);
            if (toDoTimeDate != 0) {
                processLeadTime = Math.ceil((doneTime.getTime() - toDoTimeDate.getTime()) / 86400000);
            } else {
                processLeadTime = toDoTime + inProgressTime + testTime + inReviewTime;
            }

            if (inProgressTimeDate != 0) {
                cycleTime = Math.ceil((doneTime.getTime() - inProgressTimeDate.getTime()) / 86400000);
            } else {
                cycleTime = inProgressTime + testTime + inReviewTime;
            }
            item.push(resolution, doneTime.toLocaleDateString('en-GB'), cycleTime.toString(), processLeadTime.toString(), leadTime.toString(), backlogTime.toString(), inRefinementTime.toString(), readyForSprintTime.toString(), toDoTime.toString(), analyseTime.toString(), inProgressTime.toString(), testTime.toString(), inReviewTime.toString())
            csvList.push(item)
        } else {  }


    }

    /** This writes the created CSV list to an excel file. Make sure the excel file exists, is closed and in the same location as the script.
     * Row A should have the following content: Issue Key	Type	Title	Component	Resolution	FinishDate	CycleTime	ProcessLeadTime	LeadTime	TIS Backlog	TIS Refinement	TIS Refined	TIS To Do	TIS Analyse	TIS Realisatie	TIS Testen	TIS Validatie Klant
     * The delay is necessary to make sure all calculations have finished before attempting to write the excel. 
     */
    await delay(25000)
    console.log("Exporting data from JIRA and creating Excel file.")
    const file = reader.readFile('./MetricsCalculator.xlsx', { type: 'type', cellDates: true, dateNF: "dd/mm/yy" })
    const worksheet = file.Sheets[file.SheetNames[0]]

    reader.utils.sheet_add_aoa(worksheet, csvList, { origin: { r: 1, c: 0 } })
    reader.writeFile(file, "./MetricsCalculator.xlsx")

    console.log("Finished!")
}

getCsv();




