//variable that stores the row the information starts on.
let startRow = 3;
//stores the total rows in the sheet
let totalRows = 1000;
//stores the oversight column location
let oversightColumn = 2;
//stores the organization column location
let organizationColumn = 3;
//stores the calling column location
let callingColumn = 4;
//stores the member column location
let memberNameColumn = 5;
//stores the starting column for the status
let statusStartColumn = 6;
//stores the total number of various statuses.
let statusColumnsWidth = 13;
//Creates a new instance of the spreadsheet app so we can access the methods associated with it
var activeSpreadsheetApp = SpreadsheetApp.getActiveSpreadsheet();
//Creates a new instance of a single sheet in the app to access the associated methods.
var wardCallingsSheet = activeSpreadsheetApp.getSheetByName("Ward Callings by Organization");
let prayerListSheet = activeSpreadsheetApp.getSheetByName("Prayer List");
let sustainingsSheet = activeSpreadsheetApp.getSheetByName("Sustainings to Announce");
let addLCRSheet = activeSpreadsheetApp.getSheetByName("To Add To LCR");
let setApartSheet = activeSpreadsheetApp.getSheetByName("Set Apart(LCR)");


//Creates several parallel arrays. 
//E.G. oversightArray[i] will contain the oversight while callingArray[i] will contain the associated calling.
var oversightArray = wardCallingsSheet.getRange(startRow, oversightColumn, totalRows, 1).getValues();
var organizationArray = wardCallingsSheet.getRange(startRow, organizationColumn, totalRows, 1).getValues();
var callingArray = wardCallingsSheet.getRange(startRow, callingColumn, totalRows, 1).getValues();
var nameArray = wardCallingsSheet.getRange(startRow, memberNameColumn, totalRows, 1).getValues();
var sustainDate = wardCallingsSheet.getRange(startRow, memberNameColumn + 5, totalRows, 1).getValues();

//creates a status array that is an array of arrays. Status array contains arrays of Xs.
//example statusArray[i] = [x,x,x,x,,,,];
var statusArray = wardCallingsSheet.getRange(startRow, statusStartColumn, totalRows, statusColumnsWidth).getValues();

//Create parallel arrays to store information for each report.
var prayerlistNames = [];
var prayerlistCallings = [];
var prayerlistOrganization = [];

var appointmentNames = [];
var appointmentCallings = [];
var appointmentOrganization = [];
var appointmentOversight = [];

var toExtendNames = [];
var toExtendCallings = [];
var toExtendOrganization = [];
var toExtendOversight = [];

var sustainNames = [];
var sustainCallings = [];
var sustainOrganization = [];

var releaseNames = [];
var releaseCallings = [];
var releaseOrganization = [];

var addLcrNames = [];
var addLcrCallings = [];
var addLcrOrganization = [];
var addLcrSustainDate = [];

var settingApartNames = [];
var settingApartCallings = [];
var settingApartOrganization = [];
var settingApartOversight = [];

var lcrFinalizedNames = [];
var lcrFinalizedCallings = [];
var lcrFinalizedOrganization = [];

var lcrRemoveNames = [];
var lcrRemoveCallings = [];
var lcrRemoveOrganization = [];

function r2d2() {
    clearReports();
    createReportArrays(statusArray);
    outputNoOversight();
    outputOversight();
}

function outputOversight() {
    let tempBishopNames = [];
    let tempBishopCalling = [];
    let tempBishopOrganization = [];
    let temp1stNames = [];
    let temp1stCalling = [];
    let temp1stOrganization = [];
    let temp2ndNames = [];
    let temp2ndCalling = [];
    let temp2ndOrganization = [];
    let tempEqpNames = [];
    let tempEqpCalling = [];
    let tempEqpOrganization = [];


    for (let i = 0; i < appointmentOversight.length; i++) {

        if (appointmentOversight[i][0] === "Bishop") {
            tempBishopNames.push(appointmentNames[i]);
            tempBishopCalling.push(appointmentCallings[i]);
            tempBishopOrganization.push(appointmentOrganization[i]);
        } else if (appointmentOversight[i][0] === "1st Counselor") {
            temp1stNames.push(appointmentNames[i]);
            temp1stCalling.push(appointmentCallings[i]);
            temp1stOrganization.push(appointmentOrganization[i]);
        } else if (appointmentOversight[i][0] === "2nd Counselor") {
            temp2ndNames.push(appointmentNames[i]);
            temp2ndCalling.push(appointmentCallings[i]);
            temp2ndOrganization.push(appointmentOrganization[i]);
        } else {
            tempEqpNames.push(appointmentNames[i]);
            tempEqpCalling.push(appointmentCallings[i]);
            tempEqpOrganization.push(appointmentOrganization[i]);
        }
    }

    if (tempBishopNames.length > 0) {
        activeSpreadsheetApp.getSheetByName("Appointments to Set").getRange(3, 1, tempBishopNames.length, 1).setValues(tempBishopNames);
        activeSpreadsheetApp.getSheetByName("Appointments to Set").getRange(3, 2, tempBishopNames.length, 1).setValues(tempBishopCalling);
        activeSpreadsheetApp.getSheetByName("Appointments to Set").getRange(3, 3, tempBishopNames.length, 1).setValues(tempBishopOrganization);
    }

    if (temp1stNames.length > 0) {
        activeSpreadsheetApp.getSheetByName("Appointments to Set").getRange(3, 5, temp1stNames.length, 1).setValues(temp1stNames);
        activeSpreadsheetApp.getSheetByName("Appointments to Set").getRange(3, 6, temp1stNames.length, 1).setValues(temp1stCalling);
        activeSpreadsheetApp.getSheetByName("Appointments to Set").getRange(3, 7, temp1stNames.length, 1).setValues(temp1stOrganization);
    }

    if (temp2ndNames.length > 0) {
        activeSpreadsheetApp.getSheetByName("Appointments to Set").getRange(3, 9, temp2ndNames.length, 1).setValues(temp2ndNames);
        activeSpreadsheetApp.getSheetByName("Appointments to Set").getRange(3, 10, temp2ndNames.length, 1).setValues(temp2ndCalling);
        activeSpreadsheetApp.getSheetByName("Appointments to Set").getRange(3, 11, temp2ndNames.length, 1).setValues(temp2ndOrganization);
    }

    if (tempEqpNames.length > 0) {
        activeSpreadsheetApp.getSheetByName("Appointments to Set").getRange(3, 13, tempEqpNames.length, 1).setValues(tempEqpNames);
        activeSpreadsheetApp.getSheetByName("Appointments to Set").getRange(3, 14, tempEqpNames.length, 1).setValues(tempEqpCalling);
        activeSpreadsheetApp.getSheetByName("Appointments to Set").getRange(3, 15, tempEqpNames.length, 1).setValues(tempEqpOrganization);
    }

    tempBishopNames = [];
    tempBishopCalling = [];
    tempBishopOrganization = [];
    temp1stNames = [];
    temp1stCalling = [];
    temp1stOrganization = [];
    temp2ndNames = [];
    temp2ndCalling = [];
    temp2ndOrganization = [];
    tempEqpNames = [];
    tempEqpCalling = [];
    tempEqpOrganization = [];

    for (let i = 0; i < toExtendOversight.length; i++) {

        if (toExtendOversight[i][0] === "Bishop") {
            tempBishopNames.push(toExtendNames[i]);
            tempBishopCalling.push(toExtendCallings[i]);
            tempBishopOrganization.push(toExtendOrganization[i]);
        } else if (toExtendOversight[i][0] === "1st Counselor") {
            temp1stNames.push(toExtendNames[i]);
            temp1stCalling.push(toExtendCallings[i]);
            temp1stOrganization.push(toExtendOrganization[i]);
        } else if (toExtendOversight[i][0] === "2nd Counselor") {
            temp2ndNames.push(toExtendNames[i]);
            temp2ndCalling.push(toExtendCallings[i]);
            temp2ndOrganization.push(toExtendOrganization[i]);
        } else {
            tempEqpNames.push(toExtendNames[i]);
            tempEqpCalling.push(toExtendCallings[i]);
            tempEqpOrganization.push(toExtendOrganization[i]);
        }
    }

    if (tempBishopNames.length > 0) {
        activeSpreadsheetApp.getSheetByName("Callings to Extend").getRange(3, 1, tempBishopNames.length, 1).setValues(tempBishopNames);
        activeSpreadsheetApp.getSheetByName("Callings to Extend").getRange(3, 2, tempBishopNames.length, 1).setValues(tempBishopCalling);
        activeSpreadsheetApp.getSheetByName("Callings to Extend").getRange(3, 3, tempBishopNames.length, 1).setValues(tempBishopOrganization);
    }

    if (temp1stNames.length > 0) {
        activeSpreadsheetApp.getSheetByName("Callings to Extend").getRange(3, 5, temp1stNames.length, 1).setValues(temp1stNames);
        activeSpreadsheetApp.getSheetByName("Callings to Extend").getRange(3, 6, temp1stNames.length, 1).setValues(temp1stCalling);
        activeSpreadsheetApp.getSheetByName("Callings to Extend").getRange(3, 7, temp1stNames.length, 1).setValues(temp1stOrganization);
    }

    if (temp2ndNames.length > 0) {
        activeSpreadsheetApp.getSheetByName("Callings to Extend").getRange(3, 9, temp2ndNames.length, 1).setValues(temp2ndNames);
        activeSpreadsheetApp.getSheetByName("Callings to Extend").getRange(3, 10, temp2ndNames.length, 1).setValues(temp2ndCalling);
        activeSpreadsheetApp.getSheetByName("Callings to Extend").getRange(3, 11, temp2ndNames.length, 1).setValues(temp2ndOrganization);
    }

    if (tempEqpNames.length > 0) {
        activeSpreadsheetApp.getSheetByName("Callings to Extend").getRange(3, 13, tempEqpNames.length, 1).setValues(tempEqpNames);
        activeSpreadsheetApp.getSheetByName("Callings to Extend").getRange(3, 14, tempEqpNames.length, 1).setValues(tempEqpCalling);
        activeSpreadsheetApp.getSheetByName("Callings to Extend").getRange(3, 15, tempEqpNames.length, 1).setValues(tempEqpOrganization);
    }



    tempBishopNames = [];
    tempBishopCalling = [];
    tempBishopOrganization = [];
    temp1stNames = [];
    temp1stCalling = [];
    temp1stOrganization = [];
    temp2ndNames = [];
    temp2ndCalling = [];
    temp2ndOrganization = [];
    tempEqpNames = [];
    tempEqpCalling = [];
    tempEqpOrganization = [];

    for (let i = 0; i < settingApartOversight.length; i++) {

        if (settingApartOversight[i][0] === "Bishop") {
            tempBishopNames.push(settingApartNames[i]);
            tempBishopCalling.push(settingApartCallings[i]);
            tempBishopOrganization.push(settingApartOrganization[i]);
        } else if (settingApartOversight[i][0] === "1st Counselor") {
            temp1stNames.push(settingApartNames[i]);
            temp1stCalling.push(settingApartCallings[i]);
            temp1stOrganization.push(settingApartOrganization[i]);
        } else if (settingApartOversight[i][0] === "2nd Counselor") {
            temp2ndNames.push(settingApartNames[i]);
            temp2ndCalling.push(settingApartCallings[i]);
            temp2ndOrganization.push(settingApartOrganization[i]);
        } else {
            tempEqpNames.push(settingApartNames[i]);
            tempEqpCalling.push(settingApartCallings[i]);
            tempEqpOrganization.push(settingApartOrganization[i]);
        }
    }

    if (tempBishopNames.length > 0) {
        activeSpreadsheetApp.getSheetByName("To Be Set Apart").getRange(3, 1, tempBishopNames.length, 1).setValues(tempBishopNames);
        activeSpreadsheetApp.getSheetByName("To Be Set Apart").getRange(3, 2, tempBishopNames.length, 1).setValues(tempBishopCalling);
        activeSpreadsheetApp.getSheetByName("To Be Set Apart").getRange(3, 3, tempBishopNames.length, 1).setValues(tempBishopOrganization);
    }

    if (temp1stNames.length > 0) {
        activeSpreadsheetApp.getSheetByName("To Be Set Apart").getRange(3, 5, temp1stNames.length, 1).setValues(temp1stNames);
        activeSpreadsheetApp.getSheetByName("To Be Set Apart").getRange(3, 6, temp1stNames.length, 1).setValues(temp1stCalling);
        activeSpreadsheetApp.getSheetByName("To Be Set Apart").getRange(3, 7, temp1stNames.length, 1).setValues(temp1stOrganization);
    }

    if (temp2ndNames.length > 0) {
        activeSpreadsheetApp.getSheetByName("To Be Set Apart").getRange(3, 9, temp2ndNames.length, 1).setValues(temp2ndNames);
        activeSpreadsheetApp.getSheetByName("To Be Set Apart").getRange(3, 10, temp2ndNames.length, 1).setValues(temp2ndCalling);
        activeSpreadsheetApp.getSheetByName("To Be Set Apart").getRange(3, 11, temp2ndNames.length, 1).setValues(temp2ndOrganization);
    }

    if (tempEqpNames.length > 0) {
        activeSpreadsheetApp.getSheetByName("To Be Set Apart").getRange(3, 13, tempEqpNames.length, 1).setValues(tempEqpNames);
        activeSpreadsheetApp.getSheetByName("To Be Set Apart").getRange(3, 14, tempEqpNames.length, 1).setValues(tempEqpCalling);
        activeSpreadsheetApp.getSheetByName("To Be Set Apart").getRange(3, 15, tempEqpNames.length, 1).setValues(tempEqpOrganization);
    }
}

function createReportArrays(pStatusArray) {
    //loop through the status array. All the arrays are parallel so what we do to one array we need to do to all the arrays.
    for (let i = 0; i < pStatusArray.length; i++) {
        let status = statusChecker(pStatusArray[i]).toLowerCase();

        switch (status) {
            case "agreed":
                prayerlistNames.push(nameArray[i]);
                prayerlistOrganization.push(organizationArray[i]);
                prayerlistCallings.push(callingArray[i]);
                break;
            case "prayed about":
                appointmentNames.push(nameArray[i]);
                appointmentOrganization.push(organizationArray[i]);
                appointmentCallings.push(callingArray[i]);
                appointmentOversight.push(oversightArray[i]);
                break;

            case "appt set":
                toExtendNames.push(nameArray[i]);
                toExtendOrganization.push(organizationArray[i]);
                toExtendCallings.push(callingArray[i]);
                toExtendOversight.push(oversightArray[i]);
                break;

            case "called":
                sustainNames.push(nameArray[i]);
                sustainOrganization.push(organizationArray[i]);
                sustainCallings.push(callingArray[i]);
                break;

            case "sustained":
                addLcrNames.push(nameArray[i]);
                addLcrOrganization.push(organizationArray[i]);
                addLcrCallings.push(callingArray[i]);
                addLcrSustainDate.push(sustainDate[i]);
                break;

            case "lcr entered":
                settingApartNames.push(nameArray[i]);
                settingApartOrganization.push(organizationArray[i]);
                settingApartCallings.push(callingArray[i]);
                settingApartOversight.push(oversightArray[i]);

                break;

            case "set apart":
                lcrFinalizedNames.push(nameArray[i]);
                lcrFinalizedOrganization.push(organizationArray[i]);
                lcrFinalizedCallings.push(callingArray[i]);
                break;
            case "approved release":
                appointmentNames.push(nameArray[i]);
                organizationArray[i][0] += " - Release";
                appointmentOrganization.push(organizationArray[i]);
                appointmentCallings.push(callingArray[i]);
                appointmentOversight.push(oversightArray[i]);
                break;

            case "release appointment":
                toExtendNames.push(nameArray[i]);
                organizationArray[i][0] += " - Release";
                toExtendOrganization.push(organizationArray[i]);
                toExtendCallings.push(callingArray[i]);
                toExtendOversight.push(oversightArray[i]);
                break;

            case "extended release":
                releaseNames.push(nameArray[i]);
                releaseOrganization.push(organizationArray[i]);
                releaseCallings.push(callingArray[i]);
                break;

            case "released":
                lcrRemoveNames.push(nameArray[i]);
                lcrRemoveOrganization.push(organizationArray[i]);
                lcrRemoveCallings.push(callingArray[i]);
                break;

            default:
                //this means everyone is set apart and in LCR or it isn't a calling slot
                //so we won't do anything. I'm just writing this statment in case I ever need
                //to do something later.
                break;
        }

    }
}



//returns the status based on the Xs.
function statusChecker(pStatusArray) {
    let status = "";
    //gets the status array of Xs and returns the status of a calling.
    switch (true) {
        case pStatusArray[0] != "" && pStatusArray[1] === "":
            status = "agreed";
            break;

        case pStatusArray[1] != "" && pStatusArray[2] === "":
            status = "prayed about";
            break;

        case pStatusArray[2] != "" && pStatusArray[3] === "":
            status = "appt set";
            break;

        case pStatusArray[3] != "" && pStatusArray[4] === "":
            status = "called";
            break;

        case pStatusArray[4] != "" && pStatusArray[5] === "":
            status = "sustained";
            break;

        case pStatusArray[5] != "" && pStatusArray[6] === "":
            status = "LCR Entered";
            break;

        case pStatusArray[6] != "" && pStatusArray[7] === "":
            status = "Set Apart";
            break;

        case pStatusArray[7] != "" && pStatusArray[8] === "":
            status = "LCR Set Apart";
            break;

        case pStatusArray[8] != "" && pStatusArray[9] === "":
            status = "approved release";
            break;

        case pStatusArray[9] != "" && pStatusArray[10] === "":
            status = "release appointment";
            break;

        case pStatusArray[10] != "" && pStatusArray[11] === "":
            status = "extended release";
            break;

        case pStatusArray[11] != "" && pStatusArray[12] === "":
            status = "released"
            break;

        default:
            //I don't really need it to do anything if it isn't in one of the current statuses.
            break;
    }
    return status;
}

function clearReports() {
    prayerListSheet.getRange(2, 1, totalRows, 3).clear();
    activeSpreadsheetApp.getSheetByName("Appointments to Set").getRange(3, 1, totalRows, 15).clear();
    activeSpreadsheetApp.getSheetByName("Callings to Extend").getRange(3, 1, totalRows, 15).clear();
    sustainingsSheet.getRange(2, 1, totalRows, 3).clear();
    sustainingsSheet.getRange(2, 5, totalRows, 3).clear();
    activeSpreadsheetApp.getSheetByName("To Be Set Apart").getRange(3, 1, totalRows, 15).clear();
    addLCRSheet.getRange(2, 1, totalRows, 3).clear();
    activeSpreadsheetApp.getSheetByName("Set Apart(LCR)").getRange(2, 1, totalRows, 3).clear();
    activeSpreadsheetApp.getSheetByName("Remove From LCR").getRange(2, 1, totalRows, 3).clear();


}

//fills in the sheets that don't need to be split by oversight
function outputNoOversight() {
    //creates output row lengths so we don't get an error with zero items
    let prayerOutputRows = prayerlistNames.length;
    let sustainingsOutputRows = sustainNames.length;
    let toAddToLCRRows = addLcrNames.length;
    let setApartLCRRows = lcrFinalizedNames.length;
    let releaseRows = releaseNames.length;
    let lcrRemoveRows = lcrRemoveNames.length;

    if (prayerOutputRows >= 1) {
        prayerListSheet.getRange(2, 1, prayerOutputRows, 1).setValues(prayerlistNames);
        prayerListSheet.getRange(2, 2, prayerOutputRows, 1).setValues(prayerlistCallings);
        prayerListSheet.getRange(2, 3, prayerOutputRows, 1).setValues(prayerlistOrganization);
    }

    if (sustainingsOutputRows >= 1) {
        sustainingsSheet.getRange(2, 1, sustainingsOutputRows, 1).setValues(sustainNames);
        sustainingsSheet.getRange(2, 2, sustainingsOutputRows, 1).setValues(sustainCallings);
        sustainingsSheet.getRange(2, 3, sustainingsOutputRows, 1).setValues(sustainOrganization);
    }

    if (toAddToLCRRows >= 1) {
        addLCRSheet.getRange(2, 1, toAddToLCRRows, 1).setValues(addLcrNames);
        addLCRSheet.getRange(2, 2, toAddToLCRRows, 1).setValues(addLcrCallings);
        addLCRSheet.getRange(2, 3, toAddToLCRRows, 1).setValues(addLcrOrganization);
        addLCRSheet.getRange(2, 4, toAddToLCRRows, 1).setValues(addLcrSustainDate);
    }

    if (setApartLCRRows >= 1) {
        setApartSheet.getRange(2, 1, setApartLCRRows, 1).setValues(lcrFinalizedNames);
        setApartSheet.getRange(2, 2, setApartLCRRows, 1).setValues(lcrFinalizedCallings);
        setApartSheet.getRange(2, 3, setApartLCRRows, 1).setValues(lcrFinalizedOrganization);
    }

    if (releaseRows >= 1) {
        sustainingsSheet.getRange(2, 5, releaseRows, 1).setValues(releaseNames);
        sustainingsSheet.getRange(2, 6, releaseRows, 1).setValues(releaseCallings);
        sustainingsSheet.getRange(2, 7, releaseRows, 1).setValues(releaseOrganization);
    }

    if (lcrRemoveRows >= 1) {
        activeSpreadsheetApp.getSheetByName("Remove From LCR").getRange(2, 1, lcrRemoveRows, 1).setValues(lcrRemoveNames);
        activeSpreadsheetApp.getSheetByName("Remove From LCR").getRange(2, 2, lcrRemoveRows, 1).setValues(lcrRemoveCallings);
        activeSpreadsheetApp.getSheetByName("Remove From LCR").getRange(2, 3, lcrRemoveRows, 1).setValues(lcrRemoveOrganization);
    }
}