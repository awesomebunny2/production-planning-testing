import globalVar from "./globalVar.js";
import { deactivateEvents, activateEvents, refreshPivotTable, loadError } from "./universalFunctions.js";
import { buildTabulatorTables, organizeData } from "./tabulatorTables.js";

//====================================================================================================================================================
    //#region CLEAR PRESS SCHEDULING INFO ------------------------------------------------------------------------------------------------------------
            
        /**
         * Clears out all the data and deletes all the rows in the Press Scheduling Info table in the Validation sheet
         */
        async function clearPSInfo() {

                await Excel.run(async (context) => {

                    deactivateEvents(); //turns off workbook events

                    //================================================================================================================================
                        //#region ASSIGN SHEET VARIABLES ---------------------------------------------------------------------------------------------

                            const validation = context.workbook.worksheets.getItem("Validation");
                            const pressSchedulingInfo = validation.tables.getItem("PressSchedulingInfo");
                            const pressSchedulingInfoRows = pressSchedulingInfo.rows.load("count");

                        //#endregion -----------------------------------------------------------------------------------------------------------------
                    //================================================================================================================================

                    await context.sync();

                    //================================================================================================================================
                        //#region LOAD VARIABLES AND DELETE ALL BUT FIRST ROW ------------------------------------------------------------------------
                            
                            //can't delete all rows from a table, so grab entire count minus 1 to leave one row at the top
                            let rowCount = pressSchedulingInfoRows.count - 1;

                            //delete all but one row in the table
                            pressSchedulingInfoRows.deleteRowsAt(0, rowCount);

                        //#endregion -----------------------------------------------------------------------------------------------------------------
                    //================================================================================================================================

                    await context.sync();

                    //================================================================================================================================
                        //#region DELETE DATA FROM REMAINING ROW -------------------------------------------------------------------------------------

                            //deletes just the data from the first row in the table (the only row that should be left at this point)
                            pressSchedulingInfoRows.getItemAt(0).delete();

                        //#endregion -----------------------------------------------------------------------------------------------------------------
                    //================================================================================================================================

                    await context.sync();

                    activateEvents(); //turns workbook events back on

                    location.reload(); //reloads the workbook

                });
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region PUSH DATA TO PRESS SCHEDULING TABLE IN VALIDATION --------------------------------------------------------------------------------------

        /**
         * Updates the Press Scheduling Info Table appropriately based on where the change event took place
         * @param {String} trigger A string of text that tells the function what kind of event caused this function to fire 
         * (Populate, Taskpane, or "")
         */
        async function pressSchedulingInfoTable(trigger) {

            await Excel.run(async (context) => {

                //====================================================================================================================================
                    //#region ASSIGN SHEET VARIABLES -------------------------------------------------------------------------------------------------

                        const validation = context.workbook.worksheets.getItem("Validation");
                        const pressSchedulingInfo = validation.tables.getItem("PressSchedulingInfo");
                        const pressSchedulingInfoBodyRange = pressSchedulingInfo.getDataBodyRange().load("values");
                        const pressSchedulingInfoHeaderRange = pressSchedulingInfo.getHeaderRowRange().load("values");
                        const pressSchedulingInfoRows = pressSchedulingInfo.rows.load("count");

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                await context.sync();

                //====================================================================================================================================
                    //#region LOAD SHEET VARIABLES ---------------------------------------------------------------------------------------------------

                        let pressSchedulingArr = pressSchedulingInfoBodyRange.values;

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                let silkSchedulingInfo = [];
                let textSchedulingInfo = [];
                let digSchedulingInfo = [];

                let oneBigArr = [];

                let updateDataWithArrInfo = false;

                //====================================================================================================================================
                    //#region ON POPULATE FORMS BUTTON PRESS -----------------------------------------------------------------------------------------

                        if (trigger == "Populate") {

                            let emptyCell = false;
                            let silkEmpty = true;
                            let textEmpty = true;
                            let digitalEmpty = true;

                            //if a single cell in the row is not empty, force emptyCell to be false
                            for (let cell of pressSchedulingArr[0]) {
                                if (cell == "") {
                                    emptyCell = true;
                                } else {
                                    emptyCell = false;
                                    break;
                                }
                            };

                            //empties these variables just in case
                            silkSchedulingInfo = [];
                            textSchedulingInfo = [];
                            digSchedulingInfo = [];

                            //========================================================================================================================
                                //#region ON TABLE EMPTY ---------------------------------------------------------------------------------------------

                                    //if the press scheduling table is only 1 row and said row is empty, simply add the E2R data to it
                                    if (pressSchedulingArr.length === 1 && emptyCell == true) {

                                        console.log("Press Scheduling Info table is empty!");

                                        silkSchedulingInfo = createArrFromObj(globalVar.silkDataSet);
                                        textSchedulingInfo = createArrFromObj(globalVar.textDataSet);
                                        digSchedulingInfo = createArrFromObj(globalVar.digDataSet);

                                        //empties this array everytime this loop goes through
                                        oneBigArr = [];

                                        pushToBigArr(silkSchedulingInfo, oneBigArr);
                                        pushToBigArr(textSchedulingInfo, oneBigArr);
                                        pushToBigArr(digSchedulingInfo, oneBigArr);

                                        await context.sync();

                                        pressSchedulingInfo.rows.add(
                                            null,
                                            oneBigArr,
                                            true,
                                        );

                                //#endregion ---------------------------------------------------------------------------------------------------------
                            //========================================================================================================================

                            //========================================================================================================================
                                //#region IF TABLE IS NOT EMPTY --------------------------------------------------------------------------------------

                                    } else {

                                        console.log("populate was pressed and the table is not empty!!!");

                                        let i = 0;
                                        let cheeseSoup = [];
                                        let snailGloves = [];

                                        let pressSchedArrUpdated = [];
                                        let newPressSchedArrUpdated = [];

                                        //============================================================================================================
                                            //#region PUSH SILK INFO TO PRESS SCHEDULING INFO TABLE --------------------------------------------------

                                                //match the info in the silk data set to the press scheduling table
                                                let silkInfo = matchRewrite(pressSchedulingArr, i, globalVar.silkDataSet, "Silk");

                                                let silkArr = silkInfo.arr;

                                                i = silkInfo.index;

                                                silkArr.forEach((row) => {
                                                    pressSchedArrUpdated.push(row);
                                                });

                                                pressSchedulingArr.forEach((row) => {
                                                    if (row[0] !== "Silk") {
                                                        pressSchedArrUpdated.push(row);
                                                    };
                                                });

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                        //============================================================================================================
                                            //#region PUSH TEXT INFO TO PRESS SCHEDULING INFO TABLE --------------------------------------------------

                                                //match info in text data set to press scheduling info
                                                let textInfo = matchRewrite(pressSchedArrUpdated, i, globalVar.textDataSet, "Text");

                                                let textArr = textInfo.arr;

                                                i = textInfo.index;


                                                //push silk info first since it goes above text in the validation table
                                                silkArr.forEach((row) => {
                                                    newPressSchedArrUpdated.push(row);
                                                });

                                                //next push text info
                                                textArr.forEach((row) => {
                                                    newPressSchedArrUpdated.push(row);
                                                });

                                                pressSchedulingArr.forEach((row) => {

                                                    if (row[0] !== "Silk" && row[0] !== "Text") {
                                                        newPressSchedArrUpdated.push(row);
                                                    };

                                                });

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                        //============================================================================================================
                                            //#region PUSH DIGITAL INFO TO PRESS SCHEDULING INFO TABLE -----------------------------------------------

                                                //match info in digital data set to press scheduling info
                                                let digInfo = matchRewrite(newPressSchedArrUpdated, i, globalVar.digDataSet, "Digital");

                                                let digArr = digInfo.arr;

                                                i = digInfo.index;

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                        //============================================================================================================
                                            //#region COMBINE ALL TYPE ARRAYS INTO ONE BIG ARRAY TO PUSH TO TABLE ------------------------------------

                                                let bigBoi = [];

                                                pushToBigArr(silkArr, bigBoi);
                                                pushToBigArr(textArr, bigBoi);
                                                pushToBigArr(digArr, bigBoi);

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                        //============================================================================================================
                                            //#region DELETE PRESS SCHEDULING TABLE AND REMAKE IT WITH NEW VALUES ------------------------------------

                                                let rowCount = pressSchedulingInfoRows.count - 1;

                                                pressSchedulingInfoRows.deleteRowsAt(0, rowCount);

                                                await context.sync();

                                                pressSchedulingInfoRows.getItemAt(0).delete();

                                                await context.sync();

                                                pressSchedulingInfo.rows.add(
                                                    null,
                                                    bigBoi,
                                                    true,
                                                );

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                    };

                                //#endregion ---------------------------------------------------------------------------------------------------------
                            //========================================================================================================================

                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region ON TASKPANE CHANGE -----------------------------------------------------------------------------------------------------

                        //replace table array with new dataSet info and then push to excel table
                        if (trigger == "Taskpane") {

                            let rowCount = pressSchedulingInfoRows.count - 1;

                            pressSchedulingInfoRows.deleteRowsAt(0, rowCount);

                            silkSchedulingInfo = createArrFromObj(globalVar.silkDataSet);
                            textSchedulingInfo = createArrFromObj(globalVar.textDataSet);
                            digSchedulingInfo = createArrFromObj(globalVar.digDataSet);

                            oneBigArr = [];

                            pushToBigArr(silkSchedulingInfo, oneBigArr);
                            pushToBigArr(textSchedulingInfo, oneBigArr);
                            pushToBigArr(digSchedulingInfo, oneBigArr);

                            await context.sync();

                            pressSchedulingInfoRows.getItemAt(0).delete();

                            await context.sync();

                            pressSchedulingInfo.rows.add(
                                null,
                                oneBigArr,
                                true,
                            );


                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region IF TRIGGER IS NOT FROM A TASKPANE ELEMENT ------------------------------------------------------------------------------

                            //this should only fire when the val table changes and after the content is matched when values exist in table when 
                        //populate button is pressed
                        if (trigger !== "Taskpane") {

                            const pressSchedUpdateBodyRange = pressSchedulingInfo.getDataBodyRange().load("values");

                            await context.sync();

                            let pressSchedUpdate = pressSchedUpdateBodyRange.values;

                            //updates the data sets from the press scheduling table
                            updateDataFromTable(pressSchedUpdate);

                            //update taskpane tabulator tables
                            globalVar.silkTable = buildTabulatorTables("silk-form", globalVar.silkTable, globalVar.silkDataSet);
                            globalVar.textTable = buildTabulatorTables("text-form", globalVar.textTable, globalVar.textDataSet);
                            globalVar.digTable = buildTabulatorTables("dig-form", globalVar.digTable, globalVar.digDataSet);

                            organizeData(); //update static HTML tables values

                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                refreshPivotTable(); //refreshes the pivot tables

            });

            activateEvents();

            globalVar.scrollErr.scrollTop = globalVar.scrollHeight;

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region PRESS SCHEDULING INFO CHANGE EVENT HANDLER ---------------------------------------------------------------------------------------------
            
        /**
         * The function that handles what changes to make in the workbook if a value in the Press Scheduling Info table is changed directly
         * @param {Event} event The change event
         */
        async function pressSchedulerHandler(event) {

            deactivateEvents();

            await Excel.run(async (context) => {

                console.log("val table was changed!");

                //====================================================================================================================================
                    //#region ASSIGN SHEET VARIABLES -------------------------------------------------------------------------------------------------
                            
                        let details = event.details;
                        let address = event.address;
                        let changeType = event.changeType;

                        let changedWorksheet = context.workbook.worksheets.getItem(event.worksheetId).load("name");
                        let changedAddress = changedWorksheet.getRange(address);
                        changedAddress.load("columnIndex");
                        changedAddress.load("rowIndex");

                        var allTables = context.workbook.tables;
                        allTables.load("items/name");
                        let changedTable = context.workbook.tables.getItem(event.tableId).load("name");
                        let changedTableColumns = changedTable.columns
                        changedTableColumns.load("items/name");
                        let changedTableRows = changedTable.rows;
                        changedTableRows.load("items");

                        const validationSheet = context.workbook.worksheets.getItem("Validation").load("name");
                        const pressSchedulingInfo = validationSheet.tables.getItem("PressSchedulingInfo");
                        const pressSchedulingBodyRange = pressSchedulingInfo.getDataBodyRange().load("values");

                        let bodyRange = changedTable.getDataBodyRange().load("values");
                        let headerRange = changedTable.getHeaderRowRange().load("values");

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                await context.sync();

                //====================================================================================================================================
                    //#region LOAD SHEET VARIABLES ---------------------------------------------------------------------------------------------------

                        let pressSchedArr = pressSchedulingBodyRange.values;

                        let tableContent = bodyRange.values;
                        let head = headerRange.values;

                        //changedAddress.rowIndex is from a worksheet level, which is 0 indexed. In the worksheet, the first row (0) is the title row,  
                        //then the next row (1) is the header. The content doesn't start until (2). However, we want the row index according to the
                        // table, which would have the first row start at 0. Since the title and header will always be the way they are, we can simply 
                        //subtract 2 from the worksheet row index to get the table row index
                        let tableRowIndex = changedAddress.rowIndex - 2;

                        let changedRowValues = changedTableRows.items[tableRowIndex].values

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region ASSIGN COLUMN VALUES TO VARIABLES --------------------------------------------------------------------------------------

                        //just to make it easier on myself for now, I am assuming that the position of the columns will not change, so I am 
                        //referencing the order in which the content should be. However, if I ever want to go back and make it more dynamic with 
                        //column placement, then I will need to load in the headers and stuff like I am doing in the art queue and assign each feild 
                        //in an object to each header and then create an array of object for every row in the table. That's a lot of effort for now 
                        //(even though i have done it before), so I am just not going to do that for now and assume that the columns will never move
                        let changedRowType = changedRowValues[0][0];
                        let changedRowForm = changedRowValues[0][2];
                        let changedRowDay = changedRowValues[0][6];
                        let changedRowPress = changedRowValues[0][7];

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region FIGURE OUT TABLE TYPE --------------------------------------------------------------------------------------------------

                        let E2RTableName;

                        if (changedRowType == "Silk") {
                            E2RTableName = "SilkE2R"
                        };

                        if (changedRowType == "Text") {
                            E2RTableName = "TextE2R"
                        };

                        if (changedRowType == "Digital") {
                            E2RTableName = "DIGE2R"
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region ASSIGN THE TABLE'S ITEMS, ROWs, AND VALUES -----------------------------------------------------------------------------

                        const E2RTable = context.workbook.tables.getItem(E2RTableName);
                        const E2RBodyRange = E2RTable.getDataBodyRange().load("values");

                        const E2RRows = E2RTable.rows;
                        E2RRows.load("items");

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                await context.sync();

                //====================================================================================================================================
                    //#region LOAD THE TABLE'S VALUES ------------------------------------------------------------------------------------------------

                        const E2RValues = E2RBodyRange.values;

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region UPDATE THE E2R TABLE'S VALUES TO MATCH THE PRESS SCHEDULING INFO TABLE -------------------------------------------------

                        for (let rowIndex in E2RValues) {

                            let E2RForm = E2RValues[rowIndex][1];

                            //if the form number in the PSI table matches the E2R form number, update the day and press values
                            if (changedRowForm == E2RForm) {
                                E2RValues[rowIndex][5] = changedRowDay;
                                E2RValues[rowIndex][6] = changedRowPress;
                                break;
                            };

                        };

                        E2RBodyRange.values = E2RValues; //writes the new values to the E2R table in Excel

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                await context.sync();

                //====================================================================================================================================
                    //#region UPDATE THE DATA SETS BASED ON THE PRESS SCHEDULING INFO TABLE ----------------------------------------------------------

                        updateDataFromTable(pressSchedArr);

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                await context.sync();

                //====================================================================================================================================
                    //#region UPDATE THE TASKPANE TABULATOR TABLES BASED ON THE PRESS SCHEDULING INFO TABLE ------------------------------------------

                        globalVar.silkTable = buildTabulatorTables("silk-form", globalVar.silkTable, globalVar.silkDataSet);
                        globalVar.textTable = buildTabulatorTables("text-form", globalVar.textTable, globalVar.textDataSet);
                        globalVar.digTable = buildTabulatorTables("dig-form", globalVar.digTable, globalVar.digDataSet);

                        //updates the static HTML tables
                        organizeData();

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                globalVar.scrollErr.scrollTop = globalVar.scrollHeight; //? Fixes the scroll issue?

                console.log("pressSchedulerHandler was fired, which updated the E2R associated with the changed value and the Taskpane");

                refreshPivotTable(); //refreshes the pivot tables

            });

            activateEvents();

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region MATCH REWRITE --------------------------------------------------------------------------------------------------------------------------
        /**
         * Tries to match the new info in the data set with the existing info in the press scheduling info table, if there is any
         * @param {Array} tableArray An array of arrays containing all the table data for the Press Scheduling table
         * @param {Number} tableRowIndex A number representing the current table row index
         * @param {Array} dataSet An array of objects containing all the updated values 
         * @param {String} tableType A string indicating which E2R we are currently using (Silk, Text, or Digital)
         * @returns 
         */
        function matchRewrite(tableArray, tableRowIndex, dataSet, tableType) {

            let newArr = [];
            let leTempArr = [];

            for (let rowNum in dataSet) {

                //====================================================================================================================================
                    //#region HANDLE IF ROW IS ADDED THAT PREVIOUSLTY DIDN'T EXIST IN PSINFO TABLE ---------------------------------------------------

                        //if this passes, this probably means that the product type previous had no info in the table, but the updated info now has it
                        if (!tableArray[tableRowIndex]) {
                            console.log("This info didn't previously exist in the val table, adding it in now!");
                            leTempArr.push(dataSet[rowNum].type);
                            leTempArr.push(dataSet[rowNum].priority);
                            leTempArr.push(dataSet[rowNum].form);
                            leTempArr.push(dataSet[rowNum].formQuantity);
                            leTempArr.push(dataSet[rowNum].sheets);
                            leTempArr.push(dataSet[rowNum].hours);
                            leTempArr.push(dataSet[rowNum].day);
                            leTempArr.push(dataSet[rowNum].press);
                            // leTempArr.push(dataSet[rowNum].operator);

                            newArr.push(leTempArr);

                            leTempArr = [];

                            tableRowIndex++;

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region MATCH TABLE TYPE AND TRY TO MATCH FORM NUMBER --------------------------------------------------------------------------

                        } else {

                            //if the table type row in tableArray matches the tableType variable, push the data to the tempArr
                            if (tableArray[tableRowIndex][0] == tableType) {

                                leTempArr.push(dataSet[rowNum].type);
                                leTempArr.push(dataSet[rowNum].priority);
                                leTempArr.push(dataSet[rowNum].form);
                                leTempArr.push(dataSet[rowNum].formQuantity);
                                leTempArr.push(dataSet[rowNum].sheets);
                                leTempArr.push(dataSet[rowNum].hours);
                                leTempArr.push(tableArray[tableRowIndex][6]);
                                leTempArr.push(tableArray[tableRowIndex][7]);

                                //if the form number in the data set match the form number in the tableArry, push the temp data to newArr
                                if (dataSet[rowNum].form == tableArray[tableRowIndex][2]) {

                                    newArr.push(leTempArr);

                                    leTempArr = [];

                                    tableRowIndex++;

                                } else {

                                    console.log(`A form number in the dataSet did not align with the val table info: \n
                                    TableType: ${tableType} \n
                                    dataSet: rowNum: ${rowNum}, form: ${dataSet[rowNum].form} \n
                                    valTableArray: rowIndex: ${tableRowIndex}, form: ${tableArray[tableRowIndex][1]}`);

                                    newArr.push(tableArray[tableRowIndex]);

                                    tableRowIndex++;

                                };

                            } else {

                                //====================================================================================================================
                                    //#region IF TABLE TYPE DOESN'T MATCH, ADD TO BOTTOM OF NEW ARRAY ------------------------------------------------
                    
                                        console.log(`
                                            The val table's type (${tableType}) does not match the new data's type (${tableArray[tableRowIndex][0]}).
                                        `);
                                        // leTempArr.splice(tableRowIndex, 0, dataSet[rowNum].type)
                                        leTempArr.push(dataSet[rowNum].type);
                                        leTempArr.push(dataSet[rowNum].priority);
                                        leTempArr.push(dataSet[rowNum].form);
                                        leTempArr.push(dataSet[rowNum].formQuantity);
                                        leTempArr.push(dataSet[rowNum].sheets);
                                        leTempArr.push(dataSet[rowNum].hours);
                                        leTempArr.push(dataSet[rowNum].day);
                                        leTempArr.push(dataSet[rowNum].press);
                                        newArr.push(leTempArr);

                                        leTempArr = [];

                                        tableRowIndex++;

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                            };

                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                leTempArr = [];

            };

            return {
                arr: newArr,
                index: tableRowIndex
            };

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region REPLACE OBJECT INFO IN DATA SETS WITH TABLE ARRAY INFO ---------------------------------------------------------------------------------

        /**
         * Replaces the object info in each of the different data set types with the info from the press scheduling info table
         * @param {Array} tableArray The array of the values of the Press Scheduling Info table
         */
        function updateDataFromTable(tableArray) {

            globalVar.silkDataSet = []; let sI = 1;
            globalVar.textDataSet = []; let tI = 1;
            globalVar.digDataSet = []; let dI = 1;

            let emptyCell = false;

            //if all cells in the first row of the PSI table are empty, set emptyCell to true
            for (let cell of tableArray[0]) {
                if (cell == "") {
                    emptyCell = true;
                } else {
                    emptyCell = false;
                    break;
                }
            };

            //PSI table is empty already if the number of rows in the table is one and all the cells are empty. So we just log that it is empty and
            //clear out all the data sets before exiting the function
            if (tableArray.length === 1 && emptyCell == true) {

                console.log("Press Scheduling Info table has been emptied!");

                globalVar.silkDataSet = [];
                globalVar.textDataSet = [];
                globalVar.digDataSet = [];

                return;

            };

            //for each row in the val table, update the temp obj info with the row info, then based on the table (silk, text, or digital) of the row, 
            //push this new data to the newly, emptied dataSet for that table type 
            for (let t = 0; t < tableArray.length; t++) {

                let zeObj = {};

                zeObj = {
                    // id: tId,
                    type: tableArray[t][0],
                    priority: tableArray[t][1],
                    form: tableArray[t][2],
                    formQuantity: tableArray[t][3],
                    sheets: tableArray[t][4],
                    hours: tableArray[t][5],
                    day: tableArray[t][6],
                    press: tableArray[t][7],
                    operator: tableArray[t][8]
                };

                if (tableArray[t][0] == "Silk") {
                    zeObj.id = sI;
                    globalVar.silkDataSet.push(zeObj);
                    sI++;

                };

                if (tableArray[t][0] == "Text") {
                    zeObj.id = tI;
                    globalVar.textDataSet.push(zeObj);
                    tI++;
                };

                if (tableArray[t][0] == "Digital") {
                    zeObj.id = dI;
                    globalVar.digDataSet.push(zeObj);
                    dI++;
                };

                globalVar.priorityNum++;

            };

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region CREATE ARRAY FROM OBJECT ---------------------------------------------------------------------------------------------------------------

        /**
        * Creates an array of arrays from the data set of objects setup to push into the press scheduling info table in the validation sheet
        * @param {Object} dataSet The object containing all the press scheduling properties from the E2Rs and the taskpane
        * @returns 
        */
        function createArrFromObj(dataSet) {

            let newTempArr = [];
            let newArray = [];

            for (let obj of dataSet) {

                newTempArr.push(obj.type);
                newTempArr.push(obj.priority);
                newTempArr.push(obj.form);
                newTempArr.push(obj.formQuantity);
                newTempArr.push(obj.sheets);
                newTempArr.push(obj.hours);
                newTempArr.push(obj.day);
                newTempArr.push(obj.press);
                // newTempArr.push(obj.operator);

                newArray.push(newTempArr);

                newTempArr = [];

            };

            return newArray;

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region PUSH TO ONE BIG ARRAY ------------------------------------------------------------------------------------------------------------------

        /**
         * Combines all table type arrays into one big array to write to an Excel table
         * @param {Array} smallArr The array containing all the object info from one of the table types
         * @param {Array} bigArr The array that will contain all of the table type info
         */
        function pushToBigArr(smallArr, bigArr) {

            smallArr.forEach((row) => {

                bigArr.push(row);

            });

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

export { clearPSInfo, pressSchedulingInfoTable, pressSchedulerHandler, updateDataFromTable };