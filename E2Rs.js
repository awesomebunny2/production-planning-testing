import globalVar from "./globalVar.js";
import { deactivateEvents, activateEvents, conditionalFormatting, refreshPivotTable, loadError } from "./universalFunctions.js";
import { buildTabulatorTables, organizeData } from "./tabulatorTables.js";
import { updateDataFromTable } from "./pressSchedulingInfo.js";


// ===================================================================================================================================================
    //#region Between Form Number Logic --------------------------------------------------------------------------------------------------------------
        /**
         * Determines if the form number is in range of min and max.
         * @param {Number} min The beginning of the form range
         * @param {Number} max The end of the form range
         * @param {Number} test Current form number
         * @returns {Boolean} True/False
         */
        function isBetween(min, max, test) {
            return Number(test) > min && Number(test) < max;
        }
    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
// ===================================================================================================================================================

//====================================================================================================================================================
    //#region UPDATE E2RS FROM TASKPANE TABULATOR DATA -----------------------------------------------------------------------------------------------

        /**
         * Loads all the current E2R data in from Excel, runs a function to update the data in the proper E2R info, then writes new info to Excel
         * @param {String} type The type of E2R that is being proccesed (Silk, Text, or Digital)
         * @param {Number} rowForm The current row number of the the changed form
         */
        async function updateE2RFromTaskpane(type, rowForm) {
            await Excel.run(async (context) => {

                //====================================================================================================================================
                    //#region ASSIGNING E2R SHEET VARIABLES ------------------------------------------------------------------------------------------

                        const silkE2RSheet = context.workbook.worksheets.getItem("SilkE2R").load("name");
                        const textE2RSheet = context.workbook.worksheets.getItem("TextE2R").load("name");
                        const digE2RSheet = context.workbook.worksheets.getItem("DIGE2R").load("name");

                        const silkE2RTable = silkE2RSheet.tables.getItem("SilkE2R");
                        const textE2RTable = textE2RSheet.tables.getItem("TextE2R");
                        const digE2RTable = digE2RSheet.tables.getItem("DIGE2R");

                        const silkE2RBodyRangeUpdate = silkE2RTable.getDataBodyRange().load("values");
                        const textE2RBodyRangeUpdate = textE2RTable.getDataBodyRange().load("values");
                        const digE2RBodyRangeUpdate = digE2RTable.getDataBodyRange().load("values");

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                await context.sync();

                //====================================================================================================================================
                    //#region LOADING E2R SHEET VARIABLES --------------------------------------------------------------------------------------------

                        let silkE2RArr = silkE2RBodyRangeUpdate.values;
                        let textE2RArr = textE2RBodyRangeUpdate.values;
                        let digE2RArr = digE2RBodyRangeUpdate.values;

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region UPDATING E2R VALUES ----------------------------------------------------------------------------------------------------

                        if (type == "Silk") {
                            let newSilkArr = updateDayPressInE2RArr(globalVar.silkDataSet, silkE2RArr, rowForm);
                            silkE2RBodyRangeUpdate.values = newSilkArr;
                        };

                        if (type == "Text") {
                            let newTextArr = updateDayPressInE2RArr(globalVar.textDataSet, textE2RArr, rowForm);
                            textE2RBodyRangeUpdate.values = newTextArr;
                        };

                        if (type == "Digital") {
                            let newDigArr = updateDayPressInE2RArr(globalVar.digDataSet, digE2RArr, rowForm);
                            digE2RBodyRangeUpdate.values = newDigArr;
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

            });
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region UPDATE DAY AND PRESS COLUMNS IN E2RS FUNCTION ------------------------------------------------------------------------------------------
            
        /**
         * Searches for the proper match in the day and press columns in the E2R
         * @param {Array} dataSet An array of objects containing all the neccessary info for each row of the data for the current E2R
         * @param {Array} arr An array of arrays containing all the table data for the current E2R
         * @param {Number} rowForm The current row number of the changed form
         * @returns 
         */
        function updateDayPressInE2RArr(dataSet, arr, rowForm) {
            let dayUpdate;
            let pressUpdate;
            let found = false;

            //* For each row in the dataSet, tries to match the form of the row to the rowForm number. If it finds a match, updates the day and press
            //* variables to the values in the matched row and sets found to true.
            dataSet.forEach((dataRow) => {
                if (dataRow.form == rowForm) {
                    dayUpdate = dataRow.day;
                    pressUpdate = dataRow.press;
                    found = true;
                };
            });

            //* If the row was found in the previous step, finds the rowNum in the E2R table data array and updates the day and press values there.
            if (found) {
                for (let rowIndex in arr) {
                    let E2RForm = arr[rowIndex][1];

                    if (rowForm == E2RForm) {
                        arr[rowIndex][5] = dayUpdate;
                        arr[rowIndex][6] = pressUpdate;
                        break;
                    };
                };
            } else {
                console.log("Could not find a match for the form number from the taskpane in the E2R to update");
            };

            return arr;
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region E2R CHANGE EVENT HANDLER ---------------------------------------------------------------------------------------------------------------
            
        /**
         * When data is changed in an E2R, updates the proper data in the Press Scheduling Info table and then also in the tabulator taskpane elements
         * @param {Event} event The properties of the change event 
         */
        async function E2RHandler(event) {

            deactivateEvents();

            await Excel.run(async (context) => {

                //====================================================================================================================================
                    //#region HANDLE REMOTE CHANGES --------------------------------------------------------------------------------------------------

                        if (event.source == "Remote") {
                            console.log("Content was changed by a remote user, exiting E2RHandler Event");
                            return;
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region ASSIGN SHEET VARIABLES -------------------------------------------------------------------------------------------------

                        let details = event.details;
                        let address = event.address;
                        let changeType = event.changeType;

                        let changedWorksheet = context.workbook.worksheets.getItem(event.worksheetId).load("name");
                        let changedAddress = changedWorksheet.getRange(address);
                        changedAddress.load("columnIndex");
                        changedAddress.load("rowIndex");

                        let changedTable = context.workbook.tables.getItem(event.tableId).load("name");
                        let changedTableColumns = changedTable.columns
                        changedTableColumns.load("items/name");
                        let changedTableRows = changedTable.rows;
                        changedTableRows.load("items");

                        const validationSheet = context.workbook.worksheets.getItem("Validation").load("name");
                        const silkE2RSheet = context.workbook.worksheets.getItem("SilkE2R").load("name");
                        const textE2RSheet = context.workbook.worksheets.getItem("TextE2R").load("name");
                        const digE2RSheet = context.workbook.worksheets.getItem("DIGE2R").load("name");

                        const pressSchedulingInfo = validationSheet.tables.getItem("PressSchedulingInfo");
                        const silkE2RTable = silkE2RSheet.tables.getItem("SilkE2R");
                        const textE2RTable = textE2RSheet.tables.getItem("TextE2R");
                        const digE2RTable = digE2RSheet.tables.getItem("DIGE2R");

                        const pressSchedulingBodyRange = pressSchedulingInfo.getDataBodyRange().load("values");
                        const silkE2RBodyRangeUpdate = silkE2RTable.getDataBodyRange().load("values");
                        const textE2RBodyRangeUpdate = textE2RTable.getDataBodyRange().load("values");
                        const digE2RBodyRangeUpdate = digE2RTable.getDataBodyRange().load("values");

                        let bodyRange = changedTable.getDataBodyRange().load("values");
                        let headerRange = changedTable.getHeaderRowRange().load("values");

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                await context.sync();

                //====================================================================================================================================
                    //#region LOAD SHEET VARIABLES ---------------------------------------------------------------------------------------------------

                        let pressScheduleArr = pressSchedulingBodyRange.values;
                        let silkE2RArr = silkE2RBodyRangeUpdate.values;
                        let textE2RArr = textE2RBodyRangeUpdate.values;
                        let digE2RArr = digE2RBodyRangeUpdate.values;

                        let tableContent = bodyRange.values;
                        let head = headerRange.values;


                        let tableRowIndex = changedAddress.rowIndex - 2;
                        let changedRowValues = changedTableRows.items[tableRowIndex].values

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region ASSIGN FORM, DAY, & PRESS VALUES FROM CHANGEDROWVALUES TO THEIR OWN VARIABLES ------------------------------------------

                        let changedRowForm = changedRowValues[0][1];
                        let changedRowDay = changedRowValues[0][5];
                        let changedRowPress = changedRowValues[0][6];

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region UPDATE DATA IN SILK PRESS SCHEDULING ARRAY -----------------------------------------------------------------------------

                        if (changedWorksheet.name == "SilkE2R") {

                            for (let rowIndex in pressScheduleArr) {
                                if ((pressScheduleArr[rowIndex][0] == "Silk") && (pressScheduleArr[rowIndex][2] == changedRowForm)) {
                                    pressScheduleArr[rowIndex][6] = changedRowDay;
                                    pressScheduleArr[rowIndex][7] = changedRowPress;
                                    break;
                                };
                            };
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region UPDATE DATA IN TEXT PRESS SCHEDULING ARRAY -----------------------------------------------------------------------------
                            
                        if (changedWorksheet.name == "TextE2R") {

                            for (let rowIndex in pressScheduleArr) {
                                if ((pressScheduleArr[rowIndex][0] == "Text") && (pressScheduleArr[rowIndex][2] == changedRowForm)) {
                                    pressScheduleArr[rowIndex][6] = changedRowDay;
                                    pressScheduleArr[rowIndex][7] = changedRowPress;
                                    break;
                                };
                            };
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region UPDATE DATA IN DIGITAL PRESS SCHEDULING ARRAY --------------------------------------------------------------------------

                        if (changedWorksheet.name == "DIGE2R") {

                            for (let rowIndex in pressScheduleArr) {
                                if ((pressScheduleArr[rowIndex][0] == "Digital") && (pressScheduleArr[rowIndex][2] == changedRowForm)) {
                                    pressScheduleArr[rowIndex][6] = changedRowDay;
                                    pressScheduleArr[rowIndex][7] = changedRowPress;
                                    break;
                                };
                            };
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region WRITE NEW VALUES TO PRESS SCHEDULING TABLE IN VALIDATION ---------------------------------------------------------------

                        pressSchedulingBodyRange.values = pressScheduleArr;
                    
                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================
                
                await context.sync();

                //====================================================================================================================================
                    //#region UPDATE THE DATA SETS FROM THE PRESS SCHEDULING INFO TABLE --------------------------------------------------------------

                        //Replaces the object info in each of the different data set types with the info from the press scheduling info table
                        updateDataFromTable(pressScheduleArr);

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region REBUILD THE TABULATOR TABLES IN THE TASKPANE WITH NEW DATA SETS --------------------------------------------------------

                        globalVar.silkTable = buildTabulatorTables("silk-form", globalVar.silkTable, globalVar.silkDataSet);
                        globalVar.textTable = buildTabulatorTables("text-form", globalVar.textTable, globalVar.textDataSet);
                        globalVar.digTable = buildTabulatorTables("dig-form", globalVar.digTable, globalVar.digDataSet);

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region UPDATE STATIC HTML TABLE WITH UPDATED DATA -----------------------------------------------------------------------------

                        organizeData();

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                globalVar.scrollErr.scrollTop = globalVar.scrollHeight; //? Fixes the scroll jumping issue I think?

                console.log("E2RHandler was fired, which updated both the Press Scheduling Info table and the Taskpane");

                refreshPivotTable(); //refreshes the pivot tables

            });

            activateEvents();

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region POPULATE EASY TO READS -----------------------------------------------------------------------------------------------------------------

        /**
         * Takes the data from the E2R table submitted and populates the form, form quantity, sheets, and hours columns in the returned table array
         * @param {Array} tableArray Array of the E2R table values
         * @param {Range} tableRows Range of the rows of the E2R table
         * @param {Object} worksheet The worksheet object containing the E2R table
         * @param {Array} sheetHourArr Array containing the sheet hour values from the Validation sheet
         * @returns Array of arrays
         */
        function easyToReads(tableArray, tableRows, worksheet, sheetHourArr) {

            let i = 0;
            let form;

            for (var row of tableArray) { //for each row in the E2R Table...

                let rowRange = tableRows.getItemAt(i).getRange();

                let firstNumber;

                //====================================================================================================================================
                    //#region POPULATE COLUMNS 2-7 IN E2R IF IT IS A LAYOUT/FORM/TUBE ROW ------------------------------------------------------------

                        //if the row starts with the word "Layout", we know that a new form is starting. Or "Form" if it's looking at Digital
                        if (row[0].startsWith("Layout") || row[0].startsWith("Form") || row[0].startsWith("Tube")) {

                            firstNumber = row[0].match(/[0-9]+/); //the first number in this cell will always be the form number, so let's grab that

                            if (firstNumber) { //if a form number exists, write the value to the form variable for use later

                                form = firstNumber[0];

                                let fQ;

                                //====================================================================================================================
                                    //#region ASSIGNING WASTE VALUES ---------------------------------------------------------------------------------

                                        let waste = 0;

                                        //if the worksheet's name is TextE2R, we need to offset the forms by 100, so we turn form into a number and add
                                        //100 to the result. Also defining the waste amount for non-digital text and silk products
                                        if (worksheet.name == "TextE2R") {

                                            form = Number(form) + 100; //augment form by 100
                                            waste = globalVar.wasteData["Text"]["Waste"]; //text waste

                                        } else if (worksheet.name == "SilkE2R") {

                                            waste = globalVar.wasteData["Silk"]["Waste"]; //silk waste

                                        } else if (worksheet.name == "DIGE2R") {

                                            if (isBetween(200, 251, form)) { //digital silk

                                                waste = globalVar.wasteData["Digital Silk"]["Waste"];

                                            } else if (isBetween(250, 301, form)) { //digital text

                                                waste = globalVar.wasteData["Digital Text"]["Waste"];

                                            } else if (isBetween(300, 351, form)) { //digital husky

                                                waste = globalVar.wasteData["Digital Text"]["Waste"];

                                            } else if (isBetween(400, 451, form)) { //envelope windows

                                                waste = globalVar.wasteData["Digital Text"]["Waste"];

                                            } else if (isBetween(450, 501, form)) { //envelope no windows

                                                waste = globalVar.wasteData["Digital Text"]["Waste"];

                                            } else if (isBetween(500, 601, form)) { //wide format tubes

                                                waste = globalVar.wasteData["Wide Format"]["Waste"];

                                            } else if (isBetween(600, 651, form)) { //digital variable silk

                                                waste = globalVar.wasteData["Digital Silk"]["Waste"];

                                            } else if (isBetween(650, 701, form)) { //digital variable text

                                                waste = globalVar.wasteData["Digital Text"]["Waste"];

                                            };

                                        }

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                //====================================================================================================================
                                    //#region DEFINE DIFFERENT 1-SIDED SPELLING VARIABLES ------------------------------------------------------------
                                        
                                        globalVar.singleSided = false;

                                        const numSidedArr = ["1-sided", "1-Sided", "1-SIDED"]; //with number
                                        const oneSidedArr = ["One-Sided", "One-sided", "one-sided", "ONE-SIDED", "one-Sided"]; //without number

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                //====================================================================================================================
                                    //#region PROCESS ALL NON-TUBE FORMS -----------------------------------------------------------------------------

                                        //* Only do the following for all non-tube forms
                                        if (!row[0].startsWith("Tube")) { 

                                            //========================================================================================================
                                                //#region FORM QUANTITY AND 1-SIDED WASTE AUGMENTATION -----------------------------------------------

                                                    let formQuantity = 0;

                                                    if (worksheet.name == "DIGE2R") {

                                                        //============================================================================================
                                                            //#region DIGITAL FORM QUANTITY AND 1-SIDED WASTE AUGMENT --------------------------------

                                                                //index of the position where the form number starts
                                                                let startIndexOfFormNumber = row[0].indexOf(firstNumber); 

                                                                //adds the form number start index to the length of the number itself
                                                                let indexAfterForm = startIndexOfFormNumber + (firstNumber[0].length);

                                                                //isolates all the text after the form number to the end of the string
                                                                let textAfterForm = row[0].substring(indexAfterForm).trim();

                                                                //====================================================================================
                                                                    //#region ACCOUNT FOR MULTIPLE "1-SIDED" SPELLINGS -------------------------------

                                                                        // globalVar.singleSided = false;

                                                                        // const numSidedArr = ["1-sided", "1-Sided", "1-SIDED"]; //with number

                                                                        //without number
                                                                        /* const oneSidedArr = [
                                                                            "One-Sided", "One-sided", "one-sided", "ONE-SIDED", "oneSided"
                                                                        ];
                                                                        */ 

                                                                        //============================================================================
                                                                            //#region TRY DIFFERENT SPELLINGS INCLUDING NUMBER -----------------------

                                                                                for (var nSItem of numSidedArr) {

                                                                                    // let theItem = numSidedArr[nSItem];

                                                                                    //modifying waste if the form is a 1-sided form and also adjusts 
                                                                                    if (textAfterForm.startsWith(nSItem)) { 

                                                                                        waste = Math.ceil(waste / 2);

                                                                                        //removes "1-sided" from the textAfterForm string
                                                                                        textAfterForm = textAfterForm.replace(nSItem, "").trim(); 

                                                                                        //the quantity at the end of the string will be the only 
                                                                                        //number left after the form number & 1-sided is removed, 
                                                                                        //which we did above
                                                                                        formQuantity = (textAfterForm.match(/[0-9]+/))[0];

                                                                                        fQ = Number(formQuantity);

                                                                                        globalVar.singleSided = true;

                                                                                    };

                                                                                };

                                                                            //#endregion -------------------------------------------------------------
                                                                        //============================================================================

                                                                        //============================================================================
                                                                            //#region TRY DIFFERENT SPELLINGS WITHOUT NUMBER -------------------------

                                                                                //just adjusts waste for single sided & finds form quantity 
                                                                                for (var oSItem of oneSidedArr) { 

                                                                                    // let zeItem = oneSidedArr[oSItem];

                                                                                    if (textAfterForm.startsWith(oSItem)) {

                                                                                        waste = Math.ceil(waste / 2);

                                                                                        //no need to remove "One-Sided" from string since there is no 
                                                                                        //number present other than the quantity now
                                                                                        formQuantity = (textAfterForm.match(/[0-9]+/))[0];

                                                                                        fQ = Number(formQuantity);

                                                                                        globalVar.singleSided = true;

                                                                                    };

                                                                                };

                                                                            //#endregion -------------------------------------------------------------
                                                                        //============================================================================

                                                                    //#endregion ---------------------------------------------------------------------
                                                                //====================================================================================                    

                                                            //#endregion -----------------------------------------------------------------------------
                                                        //============================================================================================

                                                        //============================================================================================
                                                            //#region FORM QUANTITY FOR DOUBLE SIDED DIGITAL -----------------------------------------

                                                                if (globalVar.singleSided == false) {
                                                                    formQuantity = (textAfterForm.match(/[0-9]+/))[0];
                                                                    fQ = Number(formQuantity);
                                                                    globalVar.singleSided = false;
                                                                    // fQ = fQ * 2;
                                                                };

                                                            //#endregion -----------------------------------------------------------------------------
                                                        //============================================================================================

                                                    } else {

                                                        //============================================================================================
                                                            //#region TEXT AND SILK FORM QUANTITY AND 1-SIDED WASTE AUGMENT --------------------------

                                                                //finds the word before "House Stock". The ./s+ accounts for all spaces before 
                                                                //House Stock and a single character (the comma)
                                                                const houseStockRegex = /\w+(?=.\s+House Stock)/;
                                                                var wordBeforeHouseStock = row[0].match(houseStockRegex);

                                                                //all double sided forms should use the word "Sheetwise" before "House Stock"
                                                                if (wordBeforeHouseStock == "Sheetwise") { 

                                                                    //================================================================================
                                                                        //#region DOUBLE-SIDED FORMS (STANDARD) --------------------------------------

                                                                            //slice takes the characters between a start index and end index. The 
                                                                            //start has a +3 so that way the result does not include the 3 characters 
                                                                            //we are looking for. The end has a /-2 so that it does not include ", "
                                                                            // characters after qty.
                                                                            formQuantity = row[0].slice((row[0].indexOf("), ") + 3), 
                                                                            (row[0].indexOf("Sheetwise") - 2));

                                                                            //removes all commas and converts from string to number
                                                                            fQ = Number(formQuantity.replace(/,/g, "")); 

                                                                            globalVar.singleSided = false;

                                                                            // fQ = fQ * 2;

                                                                        //#endregion -----------------------------------------------------------------
                                                                    //================================================================================

                                                                //means it is using the new press and should cut waste by 200 for both text & silk and
                                                                //half the hours
                                                                } else if (wordBeforeHouseStock == "Perfected") { 

                                                                    //================================================================================
                                                                        //#region PERFECTED FORMS ----------------------------------------------------

                                                                            //slice takes the characters between a start index and end index. The 
                                                                            //start has a +3 so that way the result does not include the 3 characters 
                                                                            //we are looking for. The end has a -2 so that it does not include ", " 
                                                                            //characters after qty.
                                                                            formQuantity = row[0].slice((row[0].indexOf("), ") + 3), 
                                                                            (row[0].indexOf("Perfected") - 2));

                                                                            //removes all commas and converts from string to number
                                                                            fQ = Number(formQuantity.replace(/,/g, "")); 
                                                                            globalVar.singleSided = false;

                                                                            waste = waste - 200;

                                                                            //hours will be halved down below outside of this area
                                                                            // fQ = fQ * 2;

                                                                        //#endregion -----------------------------------------------------------------
                                                                    //================================================================================

                                                                } else {

                                                                    //================================================================================
                                                                        //#region SINGLE SIDED FORMS (ABNORMAL) --------------------------------------

                                                                            //try for any other variable other than Sheetwise that should appear. 
                                                                            //This one also tells /us if it's a 1-sided form. If this fails, 
                                                                            //catch the error
                                                                            try {

                                                                                //slice takes the characters between a start index and end index. The 
                                                                                //start has a +3 so that way the result does not include the 3 
                                                                                //characters we are looking for. The end has a -2 so that it does not 
                                                                                //include ", " characters after qty.
                                                                                for (let numSidedItem of numSidedArr) {
                                                                                    if (row[0].includes(numSidedItem)) {
                                                                                        formQuantity = row[0].slice((row[0].indexOf("), ") + 3), 
                                                                                        (row[0].indexOf(numSidedItem) - 2));
                                                                                    };
                                                                                };

                                                                                for (let oneSidedItem of oneSidedArr) {
                                                                                    if (row[0].includes(oneSidedItem)) {
                                                                                        formQuantity = row[0].slice((row[0].indexOf("), ") + 3), 
                                                                                        (row[0].indexOf(oneSidedItem) - 2));
                                                                                    };
                                                                                };


                                                                                //if this try succeeds, then we know this is a single sided form and 
                                                                                //the waste needs to be cut in half
                                                                                waste = Math.ceil(waste / 2); //Always round up for paper purposes

                                                                                //removes all commas and converts from string to number
                                                                                fQ = Number(formQuantity.replace(/,/g, "")); 

                                                                                globalVar.singleSided = true;

                                                                            } catch (err) {
                                                                                console.error(err);
                                                                                loadError(err.stack)
                                                                            };

                                                                        //#endregion -----------------------------------------------------------------
                                                                    //================================================================================

                                                                };

                                                            //#endregion -----------------------------------------------------------------------------
                                                        //============================================================================================

                                                    };

                                                //#endregion -----------------------------------------------------------------------------------------
                                            //========================================================================================================

                                            //========================================================================================================
                                                //#region FIGURE OUT SHEET QUANTITY (FORM QUANTITY + WASTE [DOUBLE AND SINGLE SIDED]) ----------------

                                                    let sheets = fQ + waste; //form quantity plus waste

                                                    let sheetsAdj4Side;

                                                    if (globalVar.singleSided == false) {

                                                        //double the fQ is 2-sided. To be used for the hours calcuation only
                                                        sheetsAdj4Side = (fQ * 2) + waste; 

                                                        //sheets variable is left alone since we want the value that is output to the sheet to remain 
                                                        //unaffected by sides

                                                    } else {

                                                        sheetsAdj4Side = sheets; //don't double fQ since it is single sided. 

                                                        //waste variable defaults to a 2-sided calculation, and eariler we adjusted it for 1-sided 
                                                        //within the handleOneSided function

                                                    };

                                                //#endregion -----------------------------------------------------------------------------------------
                                            //========================================================================================================

                                            //========================================================================================================
                                                //#region FIGURE OUT HOURS ---------------------------------------------------------------------------

                                                    let hours;
                                                    let roundHours;

                                                    for (var z = 0; z < sheetHourArr.length; z++) {

                                                        if (
                                                            sheets >= globalVar.sheetHourData["Sheets (Min)"][z] 
                                                            && 
                                                            sheets <= globalVar.sheetHourData["Sheets (Max)"][z]
                                                        ){

                                                            let divideBy = globalVar.sheetHourData["Prints Per Hour"][z];
                                                            hours = (sheetsAdj4Side) / divideBy;

                                                            //if word before House Stock is "Perfected", then we need to cut the hours in half
                                                            if (wordBeforeHouseStock == "Perfected") {
                                                                hours = hours / 2;
                                                            };

                                                            roundHours = Math.round((hours + Number.EPSILON) * 100) / 100;
                                                        }
                                                    };

                                                //#endregion -----------------------------------------------------------------------------------------
                                            //========================================================================================================

                                            //========================================================================================================
                                                //#region SET DEFAULT VALUES FOR DAY AND PRESS COLUMNS -----------------------------------------------

                                                    //set values for E2Rs, and default values for days and presses in E2Rs
                                                    let dayVal;
                                                    let pressVal;

                                                    //if day is empty, replace with a "-", otherwise use existing value
                                                    if (row[5] == "") {
                                                        dayVal = "-";
                                                    } else {
                                                        dayVal = row[5];
                                                    };

                                                    //if on Digital E2R, autofill all presses with "Digital", otherwise auto-fill presses with 1.
                                                    //if a value already exists though, use that instead
                                                    if (worksheet.name == "DIGE2R") {
                                                        pressVal = "Digital";
                                                    } else if (row[6] == "") {
                                                        pressVal = "1";
                                                    } else {
                                                        pressVal = row[6];
                                                    };

                                                //#endregion -----------------------------------------------------------------------------------------
                                            //========================================================================================================

                                            //========================================================================================================
                                                //#region UPDATE THE CELL VALUES FOR THE ROW ---------------------------------------------------------
                                                    
                                                    row[1] = form;
                                                    row[2] = formQuantity;
                                                    row[3] = sheets;
                                                    row[4] = roundHours;
                                                    row[5] = dayVal;
                                                    row[6] = pressVal;

                                                //#endregion -----------------------------------------------------------------------------------------
                                            //========================================================================================================

                                            //========================================================================================================
                                                //#region SET DATA VALIDATION FOR DAY AND PRESS COLUMNS ----------------------------------------------

                                                    try {

                                                        //============================================================================================
                                                            //#region STRINGIFY AND REPLACE BLANKS WITH "-" IN DOW & PRESSES ARRAYS ------------------

                                                                //====================================================================================
                                                                    //#region DAYS OF WEEK -----------------------------------------------------------

                                                                        var daysForDataVal = JSON.parse(JSON.stringify(globalVar.daysOfWeek));

                                                                        for (let entry in daysForDataVal) {
                                                                            if (daysForDataVal[entry] == " ") {
                                                                                daysForDataVal[entry] = "-";
                                                                            };
                                                                        };

                                                                        let daysString = daysForDataVal.join();

                                                                    //#endregion ---------------------------------------------------------------------
                                                                //====================================================================================

                                                                //====================================================================================
                                                                    //#region PRESSES ----------------------------------------------------------------

                                                                        var pressesForDataVal = JSON.parse(JSON.stringify(globalVar.presses));

                                                                        for (let press in pressesForDataVal) {
                                                                            if (pressesForDataVal[press] == " ") {
                                                                                pressesForDataVal[press] = "-";
                                                                            };
                                                                        };

                                                                        let pressString = pressesForDataVal.join();

                                                                    //#endregion ---------------------------------------------------------------------
                                                                //====================================================================================

                                                            //#endregion -----------------------------------------------------------------------------
                                                        //============================================================================================

                                                        //============================================================================================
                                                            //#region SET THE DATA VALIDATION --------------------------------------------------------

                                                                //====================================================================================
                                                                    //#region DOW DATA VALIDATION ----------------------------------------------------

                                                                        let directDayRange = "F" + (i + 3);

                                                                        let dayRange = worksheet.getRange(directDayRange);

                                                                        dayRange.dataValidation.clear();

                                                                        let dvDay = {
                                                                            list: {
                                                                                inCellDropdown: true,
                                                                                source: daysString
                                                                            }
                                                                        };

                                                                        dayRange.dataValidation.rule = dvDay;

                                                                        //makes everything centered (so "-" will be centered upon initally showing up)
                                                                        dayRange.format.horizontalAlignment = "Center"; 

                                                                    //#endregion ---------------------------------------------------------------------
                                                                //====================================================================================

                                                                //====================================================================================
                                                                    //#region PRESSES DATA VALIDATION ------------------------------------------------

                                                                        let directPressRange = "G" + (i + 3);

                                                                        let pressRange = worksheet.getRange(directPressRange);

                                                                        pressRange.dataValidation.clear();


                                                                        let dvPress = {
                                                                            list: {
                                                                                inCellDropdown: true,
                                                                                source: pressString
                                                                            }
                                                                        };

                                                                        pressRange.dataValidation.rule = dvPress;

                                                                        //makes everything centered (so "-" will be centered upon initally showing up)
                                                                        pressRange.format.horizontalAlignment = "Center";

                                                                    //#endregion ---------------------------------------------------------------------
                                                                //====================================================================================

                                                            //#endregion -----------------------------------------------------------------------------
                                                        //============================================================================================

                                                    } catch (e) {
                                                        console.log(e);
                                                        loadError(e.stack);
                                                    };

                                                //#endregion -----------------------------------------------------------------------------------------
                                            //========================================================================================================

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                //====================================================================================================================
                                    //#region PROCESS TUBE FORMS -------------------------------------------------------------------------------------

                                        //* Only print form for tubes, just like non form lines
                                        } else { 
                                            row[1] = form;
                                            row[2] = "-";
                                            row[3] = "-";
                                            row[4] = "-";
                                            row[5] = "-";
                                            row[6] = "-";
                                        };

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                            } else {
                                console.log(`This layout does not contain a form number...\n${row[0]}`);
                            };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region IF METRIX INFO CELL IS EMPTY -------------------------------------------------------------------------------------------

                        //if the row doesn't start with "Layout", "Form", or "Tube", simply fill the cells with nothing (they're pieces on the layout)
                        } else if (row[0] == "") {

                            row[1] = "";
                            row[2] = "";
                            row[3] = "";
                            row[4] = "";

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region IF A FORM NUMBER EXISTS ------------------------------------------------------------------------------------------------

                        //if a form number exists (which almost always will at this point), add it to the second cell in the row
                        } else if (form) { 

                            //if metrix info has the word "RUSH" in it, set globalVar.rushItem to true, which will tell us later to highlight this row in E2R
                            if (row[0].includes("RUSH")) {
                                console.log("This item includes a RUSH tag:");
                                console.log(row[0]);
                                globalVar.rushItem = true;
                            };

                            row[1] = form;
                            row[2] = "-";
                            row[3] = "-";
                            row[4] = "-";
                            row[5] = "-";
                            row[6] = "-";

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region IF ALL ELSE FAILS, CONSOLE LOG THE ROW MISSING THE FORM ----------------------------------------------------------------

                        } else {
                            console.log(`The following record does not belong to a form: \n${row[0]}`)
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region DO THE CONDITIONAL FORMATTING ------------------------------------------------------------------------------------------

                        //============================================================================================================================
                            //#region ASSIGN INDIVIDUAL CELL ADDRESSES FOR EACH COLUMN TO AN OBJECT FOR CONDITIONAL FORMATTING -----------------------

                                let metrixCell = worksheet.getCell(i + 2, 0);
                                let formCell = worksheet.getCell(i + 2, 1);
                                let quantityCell = worksheet.getCell(i + 2, 2);
                                let sheetsCell = worksheet.getCell(i + 2, 3);
                                let hoursCell = worksheet.getCell(i + 2, 4);
                                let dayCell = worksheet.getCell(i + 2, 5);
                                let pressCell = worksheet.getCell(i + 2, 6);

                                let certainAddresses = {
                                    0: metrixCell,
                                    1: formCell,
                                    2: quantityCell,
                                    3: sheetsCell,
                                    4: hoursCell,
                                    5: dayCell,
                                    6: pressCell
                                };

                            //#endregion -------------------------------------------------------------------------------------------------------------
                        //============================================================================================================================

                        //conditional formatting...
                        conditionalFormatting(worksheet, rowRange, row, certainAddresses);

                        i = i + 1; //increments i, which is specifically for the console.log above that is probably commented out at the moment...

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

            };

            return tableArray;

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

export { updateE2RFromTaskpane, E2RHandler, easyToReads };