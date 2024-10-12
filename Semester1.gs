/**********
  Semester1.gs
  Krish A. Patel

  Helps auto-calculate my Fall 2024 freshman semester volunteering hours for the NJIT Honors College and generate a summary.
**********/

let ss = SpreadsheetApp.getActiveSheet();
let TYPE_COL = "D";
let START_COL = "E";
let END_COL = "F";
let HOURS_COL = "G";
let INT_RANGE = "K2:K2";
let EXT_RANGE = "K3:K3";
let INT_EXT_REF = { "Internal": 0, "External": 1 };     // index in drop-down for internal/external hours (in case I decide to change that)

function onOpen() { updateAll(); }

function onEdit() { updateAll(); }

function updateAll() {
    let range = ss.getDataRange();

    // calc & update hours
    for (let r = 2; r <= range.getLastRow(); r++)
    {
        let startTimeRaw = ss.getRange(START_COL + r + ":" +  START_COL + r).getValue();
        let endTimeRaw = ss.getRange(END_COL + r + ":" +  END_COL + r).getValue();

        let timeMs = new Date(endTimeRaw) - new Date(startTimeRaw);
        let timeHr = timeMs / 3.6e6;
        
        let hoursCell = ss.getRange(HOURS_COL + r + ":" + HOURS_COL + r).getCell(1, 1);
        hoursCell.setValue(timeHr);
    }

    // update summary
    let intCell = ss.getRange(INT_RANGE).getCell(1, 1);
    let extCell = ss.getRange(EXT_RANGE).getCell(1, 1);
    let totalIntHrs = 0, totalExtHrs = 0;

    for (let r = 2; r <= range.getLastRow(); r++)
    {
        let typeCell = ss.getRange(TYPE_COL + r + ":" +  TYPE_COL + r).getCell(1, 1)
        let type = typeCell.getValue();
        let hours = ss.getRange(HOURS_COL + r + ":" + HOURS_COL + r).getValue();
        let rule = typeCell.getDataValidation();
        let options = rule.getCriteriaValues()[0];

        if (type == options[INT_EXT_REF["Internal"]])
            totalIntHrs += hours;
        else if (type == options[INT_EXT_REF["External"]])
            totalExtHrs += hours;
    }

    intCell.setValue(totalIntHrs);
    extCell.setValue(totalExtHrs);
}