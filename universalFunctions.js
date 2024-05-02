import globalVar from "./globalVar.js";

//====================================================================================================================================================
    //#region DEACTIVATE EVENTS ----------------------------------------------------------------------------------------------------------------------

        /**
         * Turns off all events in the workbook so that no other onChange handlers are fired while processing the current code
         */
        async function deactivateEvents() {
            await Excel.run(async (context) => {

                context.runtime.load("enableEvents");

                await context.sync();

                context.runtime.enableEvents = false;
                console.log("Events: OFF - Occured in registerOnActivateHandler");

            });
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region ACTIVATE EVENTS ------------------------------------------------------------------------------------------------------------------------

        /**
         * Turns on all events in the workbook so other onChange handlers will fire again moving forward
         */
        async function activateEvents() {
            await Excel.run(async (context) => {

                context.runtime.load("enableEvents");

                await context.sync();

                context.runtime.enableEvents = true;
                console.log("Events: ON - Occured in registerOnActivateHandler");

            });
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region CREATE DATA SET ------------------------------------------------------------------------------------------------------------------------

        /**
        * Writes the values of the E2R into properties of an object and then pushes the object to an array for the specific table type
        * @param {Array} E2RArr An array containing all the values of a specific E2R table you wish to make a dataSet out of
        * @param {String} tableId A string containing the title of the specific table you are making a data set for
        */
        function createDataSet(E2RArr, tableId) {

            let fId = 1;

            let dataSet = [];

            for (let f = 0; f < E2RArr.length; f++) {

                let newObj = {};

                //only catch the layout rows by checking if there is a form quantity present in the data
                if (E2RArr[f][2] !== "-" && E2RArr[f][2] !== "") {

                    newObj = {
                        id: fId,
                        priority: globalVar.priorityNum,
                        type: tableId,
                        day: E2RArr[f][5],
                        form: E2RArr[f][1],
                        formQuantity: E2RArr[f][2],
                        press: E2RArr[f][6],
                        operator: "-",
                        sheets: E2RArr[f][3],
                        hours: E2RArr[f][4]
                    };

                    dataSet.push(newObj);

                    fId++;

                    globalVar.priorityNum++;

                };

            };

            return dataSet;

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region CONDITIONAL FORMATTING -----------------------------------------------------------------------------------------------------------------

        /**
         * Use this function to change the formatting of a table, such as row colors, font styling, and more
         * @param {Object} worksheet The worksheet you are applying formatting to
         * @param {Range} rowRange The range of the row you are applying formatting to
         * @param {Array} rowInfo The values of the row you are applying formatting to
         * @param {Array} objOfCells An Array of individual cell ranges
         */
        function conditionalFormatting(worksheet, rowRange, rowInfo, objOfCells) {

            //the following formatting is specifically for certain items in the easy to read sheets
            if (worksheet.name == "SilkE2R" || worksheet.name == "TextE2R" || worksheet.name == "DIGE2R") {

                if (rowInfo[0].includes("RUSH")) {
                    objOfCells[0].format.font.bold = true;
                    objOfCells[0].format.font.color = "white";
                    objOfCells[0].format.fill.color = "#C00000";
                } else if (rowInfo[0].includes("REPRINT") || rowInfo[0].includes("RPT")) {
                    objOfCells[0].format.font.bold = true;
                    objOfCells[0].format.font.color = "red";
                    // objOfCells[0].format.fill.color = "white";
                    objOfCells[0].format.fill.clear();
                } else {
                    objOfCells[0].format.fill.clear()
                    objOfCells[0].format.font.color = "black"
                    objOfCells[0].format.font.bold = false;
                };


                if (rowInfo[0].startsWith("Layout") || rowInfo[0].startsWith("Form") || rowInfo[0].startsWith("Tube")) {
                    rowRange.format.font.bold = true;
                    rowRange.format.fill.color = "#FFF2CC";
                };

                //formQuantity starts at rowInfo[2], so we also start the loop at 2. This also translates into the objOfCells, allowing us to use this 
                //same number to get the range of the cell at the t position in said object
                for (var t = 2; t < 7; t++) {
                    if (rowInfo[t] == "-") {
                        objOfCells[t].format.horizontalAlignment = "Center";
                    };
                };
            };

            if (worksheet.name == "Master") {
                if (typeof (rowInfo) !== "number" && rowInfo !== "") {
                    rowRange.format.font.bold = true;
                    if (rowInfo == "MISSING") { //make missing forms red
                        rowRange.format.fill.color = "#C00000"
                        rowRange.format.font.color = "white"
                    } else if (rowInfo == "UA") { //make use art forms yellow
                        rowRange.format.fill.color = "#FFE699"
                        rowRange.format.font.color = "black"
                    } else if (rowInfo == "DIGITAL") { //make use art forms yellow
                        rowRange.format.fill.color = "#97d1a8"
                        rowRange.format.font.color = "black"
                    } else { //make all other forms with text instead of numbers grey
                        rowRange.format.fill.color = "#BFBFBF"; //#FFF2CC
                        rowRange.format.font.color = "black"

                    }
                } else {
                    rowRange.format.fill.clear()
                    rowRange.format.font.color = "black"
                };
            };

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region CREATE ROW INFO FUNCTION ---------------------------------------------------------------------------------------------------------------

        /**
         * Matches the table headers with the data and column index in the current row, assigning each as a property of each header within the empty
         * obj loaded in by the user
         * @param {Array} head An array of all the header column name values
         * @param {String} columnName The current column name value
         * @param {Array} rowValues An array of all the values of the current row in the current table
         * @param {Array} copyTable A copy of the entire table values (this way the original table doesn't get overwriten as we work in this function)
         * @param {Object} obj An empty object that will be loaded with all the current row data from the table matched to the column headers
         * @param {Number} rowIndex The index of the current row in the table
         * @param {Object} worksheet The worksheet object associated with the range
         */
        function createRowInfo(head, columnName, rowValues, copyTable, obj, rowIndex, worksheet) {

            let columnIndex = findColumnIndex(head, columnName);

            let value = rowValues[columnIndex];

            let cell = worksheet.getCell(rowIndex, columnIndex);

            const cellProps = cell.getCellProperties({
                address: true,
                format: {
                    fill: {
                        color: true
                    },
                    font: {
                        color: true,
                        bold: true,
                        italic: true
                    }
                },
                style: true
            });

            // copyTable[rowIndex][columnIndex] = value;

            obj[columnName] = {
                columnIndex,
                value,
                cellProps
            };

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region FIND COLUMN INDEX FUNCTION -------------------------------------------------------------------------------------------------------------

        /**
         * Returns the column index of the column name within the header array
         * @param {Array} header An array of all the header column name values
         * @param {String} columnName The current column name value
         * @returns Number
         */
        function findColumnIndex(header, columnName) {
            let i = 0;

            for (var column of header[0]) {
                if (column == columnName) {
                    return i;
                };

                i++;
            };
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region REFRESH PIVOT TABLES -------------------------------------------------------------------------------------------------------------------

        /**
        * Refreshes the Press Scheduling Pivot Tables
        */
        async function refreshPivotTable() {
            await Excel.run(async (context) => {
                const pressSchedulingSheet = context.workbook.worksheets.getItem("Press Scheduling").load("name");
                const dowSummaryPivotTable = pressSchedulingSheet.pivotTables.getItem("DOWSummaryPivot");
                const press1PivotTable = pressSchedulingSheet.pivotTables.getItem("Press1Pivot");
                const press2PivotTable = pressSchedulingSheet.pivotTables.getItem("Press2Pivot");
                const press3PivotTable = pressSchedulingSheet.pivotTables.getItem("Press3Pivot");
                const digitalPivotTable = pressSchedulingSheet.pivotTables.getItem("DigitalPivot");

                await context.sync();

                dowSummaryPivotTable.refresh();
                press1PivotTable.refresh();
                press2PivotTable.refresh();
                press3PivotTable.refresh();
                digitalPivotTable.refresh();

                console.log("Pivot table was refreshed!");

            });
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region TRYCATCH -------------------------------------------------------------------------------------------------------------------------------

        /**
         * Executes a callback function. If the callback function errors out, this function will log the error
         * @param {Function} callback The function the user is trying to execute
         */
        async function tryCatch(callback) {
            //console.log("Error callback type is: ");
            //console.log(typeof callback);
            //if (typeof callback === 'function') {
            try {
                await callback();
            } catch (err) {
                console.error(err);
                loadError(err.stack)

            };
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region Error Window ---------------------------------------------------------------------------------------------------------------------------

        /**
         * Generate an error window with a means to submit a ticket to DevOps.
         * @param {String} msg The error message
         */
        async function loadError(msg) {
            document.querySelector("#err-background").style.display="flex";
            const errorScreen= document.querySelector("#error-message");
            errorScreen.innerHTML = `<span>Automatically generated report<br>${"-".repeat(10)}<br>${msg}</span>`;
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================


export { deactivateEvents, activateEvents, createDataSet, conditionalFormatting, createRowInfo, refreshPivotTable, tryCatch, loadError };