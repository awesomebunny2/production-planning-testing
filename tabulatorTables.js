import globalVar from "./globalVar.js";
import { pressSchedulingInfoTable } from "./pressSchedulingInfo.js";

//====================================================================================================================================================
    //#region BUILD SELECT BOXES IN TABULATOR TABLE --------------------------------------------------------------------------------------------------

        /**
        * Builds a select box inside a specific cell in your tabulator table and populates it with your given arr of options
        * @param {Array} arr An array of the elements you wish to have as options in the select box
        * @param {Object} cell The cell in the tabulator table that you wish to have the select box in
        * @param {String} tableType A string that identifies if this belongs to the silk, text, or digital tabulator table
        * @returns String - HTML select element
        */
        function buildSelect(arr, cell, tableType) {
            let mySelect = $(`  
            <select class="select-box ${tableType}-box"
                formNum="${cell.getData().form}"
                rowId="${cell.getData().id}"
                colId="${cell.getColumn().getField()}"        
            ></select>
            `);


            arr.forEach((item) => {
                // Conditionally look to see what this row's value is
                // and select it I guess

                if (item == cell.getValue()) {
                    mySelect.append(`
                <option " selected>${item}</option>
                `);
                    return;
                }
                mySelect.append(`
                <option>${item}</option>
            `)
            });

            return mySelect;

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region BUILD TABULATOR TABLES -----------------------------------------------------------------------------------------------------------------

        /**
         * Build a tabulator table for a specific form based on a user-defined data set
         * @param {String} form A string that tells the function which form we are building (either silk, text, or digital)
         * @param {Variable} scheduleTable An empty variable to assign the newly built tabulator table data to
         * @param {Array} dataSet An array of objects that contains all the data for the tabulator table to be built from
         * @returns Object - Tabulator table data
         */
        function buildTabulatorTables(form, scheduleTable, dataSet) {
            // console.log(dataSet[0].priority)
            //initialize table
            scheduleTable = new Tabulator(`#${form}`, {
                data: dataSet, //assign data to table
                // groupBy:"type",
                layout: "fitColumns",
                // layout:"fitDataStretch",
                // resizableColumnFit:false,
                vertAlign: "middle",

                initialSort: [
                    { column: "priority", dir: "asc" }
                ],

                movableRows: true,
                // frozenRows:1,
                columns: [

                    { rowHandle: true, formatter: "handle", headerSort: false, frozen: true, width: 40, minWidth: 40 },

                    {
                        title: "", field: "form", headerSort: false, width: 110, formatter: (cell) => {
                            return `
                            <div class="form-title">
                                Form ${cell.getData().form}
                            </div>`
                        }
                    },


                    { title: "Priority", field: "priority", visible: false, sorter: "number", headerSort: false },

                    {
                        title: "DAY", field: "day", headerSort: false, widthGrow: 1, formatter: selectFormatter, formatterParams: { 
                            tableType: form 
                        }, cellClick: function (e, cell) {
                            // console.log(cell.getData().day);
                            // $(".select-box").addClass("select-arrow-active");         
                        }
                    },

                    {
                        title: "PRESS", field: "press", headerSort: false, widthGrow: 1, formatter: selectFormatter, formatterParams: { 
                            tableType: form 
                        }, cellClick: function (e, cell) {
                            // console.log(cell.getData().press);
                            // $(".select-box").addClass("select-arrow-active");         
                        }
                    },

                    {
                        //* Clears out the data in the row
                        formatter: "buttonCross", field: "clear", width: 30, hozAlign: "center", cellClick: function (e, cell) {

                            // Get the current row number
                            let cellData = cell.getRow().getData();
                            const rowNum = cellData.id;
                            const formNum = cellData.form;

                            let formType = cellData.type.toLowerCase();

                            if (formType == "digital") {
                                formType = "dig";
                            };

                            let zeNewScheduleData = scheduleTable.getData();

                            let tableForm = formType + "-form-box";
                            $(`.${tableForm}[rowid='${rowNum}']`).each(function (i) {
                                $(this).val("");
                            });

                            zeNewScheduleData[rowNum - 1].day = "";
                            zeNewScheduleData[rowNum - 1].press = "";

                            console.log(zeNewScheduleData);

                            const whichOne2 = zeNewScheduleData[0].type;

                            if (whichOne2 == "Silk") {
                                globalVar.silkDataSet = zeNewScheduleData;
                            } else if (whichOne2 == "Text") {
                                globalVar.textDataSet = zeNewScheduleData;
                            } else if (whichOne2 == "Digital") {
                                globalVar.digDataSet = zeNewScheduleData;
                            };

                            organizeData();

                            pressSchedulingInfoTable("Taskpane");
                        }
                    },
                ],
            });

            // USER MOVES THE ROW
            scheduleTable.on("rowMoved", (row) => {
                console.log("FIRED")

                globalVar.rowMoved = true;

                let newScheduleData = scheduleTable.getData();

                let pN = newScheduleData[0].priority

                // Handle if moved to 1st position
                if (row.getPosition() == 1) {
                    pN = newScheduleData[1].priority// Priority number
                }

                for (let index = 0; index < newScheduleData.length; index++) {
                    newScheduleData[index].priority = pN;
                    pN++;
                };

                scheduleTable.setData(newScheduleData);

                globalVar.scrollErr.scrollTop = globalVar.scrollHeight;

                const whichOne = newScheduleData[0].type

                if (whichOne == "Silk") {
                    globalVar.silkDataSet = newScheduleData;
                } else if (whichOne == "Text") {
                    globalVar.textDataSet = newScheduleData;
                } else if (whichOne == "Digital") {
                    globalVar.digDataSet = newScheduleData;
                };

                organizeData();

                pressSchedulingInfoTable("Taskpane");

                console.log("ROW MOVED!!!!!!!!!", row.getData().form);

            })

            return scheduleTable;

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region SELECT BOX FORMATTER BY ROW ------------------------------------------------------------------------------------------------------------

        //this "function" basically just builds select boxes for each cell in a row of data and returns an HTML property to use in the 
        //tabulator build function. Identified as a variable so it can be called as a tabulator formatter I think?
        /**
         * Builds select boxes for each cell in a row of data and returns an HTML property to use in the tabulator build function
         * @param {Object} cell The cell object of the tabulator table
         * @param {Object} param1 A string inside an object identifiying what table the data should belong to
         * @returns String - HTML property
         */
        const selectFormatter = (cell, { tableType }) => {

            // console.log(tableType);

            const header = cell.getField();

            let thisSelect;

            // https://www.oreilly.com/library/view/high-performance-javascript/9781449382308/ch04.html#if-else_versus_switch
            switch (header) {
                case "day":
                    // Loop the days
                    thisSelect = buildSelect(globalVar.daysOfWeek, cell, tableType);
                    break;
                case "operator":
                    // Loop the operator
                    thisSelect = buildSelect(globalVar.pressmen, cell, tableType);
                    break;
                case "press":
                    // Loop the presses
                    thisSelect = buildSelect(globalVar.presses, cell, tableType);
                    break;
            };

            let myDiv = $(`<div class="my-div ${tableType}-drop"></div>`)
            myDiv.append(thisSelect);
            return myDiv.prop(`outerHTML`);

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region DIVIDE TABULATOR TABLE DATA INTO WEEKDAYS AND CALL ORGANIZE, SORT, AND BUILD STATIC HTML TABLE FUNCTIONS FOR EACH ----------------------

        /**
        * Divides tabulator data up by the weekdays and then tosses it into functions to organize it for making static HTML tables
        */
        function organizeData() {

            globalVar.monday = [];
            globalVar.tuesday = [];
            globalVar.wednesday = [];
            globalVar.thursday = [];
            globalVar.friday = [];

            globalVar.operators = {};

            $("monday-static").empty();
            $("tuesday-static").empty();
            $("wednesday-static").empty();
            $("thursday-static").empty();
            $("friday-static").empty();

            globalVar.pressmen.forEach((man) => {
                globalVar.operators[man] = [];
            });

            globalVar.dotSelected = false;

            makeTablesForEachDOW(globalVar.silkDataSet);
            makeTablesForEachDOW(globalVar.textDataSet);
            makeTablesForEachDOW(globalVar.digDataSet);

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region MAKE TABLES FOR EACH DOW ---------------------------------------------------------------------------------------------------------------

        /**
        * Is feed either the Silk, Text, or Digital tabulator table data and organizes, sorts, and creates the static HTML tables
        * @param {Object} scheduleData An object containing all the data from the specific tabulator table to use in creating the 
        * static weekday tables
        */
        function makeTablesForEachDOW(scheduleData) {

            let tempObj = {};
            let tempForm = {};

            //if scheduleData is empty, remove all existing content from dow tables
            if (scheduleData.length === 0) {

                $("#monday-form").empty();
                $("#tuesday-form").empty();
                $("#wednesday-form").empty();
                $("#thursday-form").empty();
                $("#friday-form").empty();

            };

            scheduleData.forEach((row, rowIndex) => { //for each row of the data...

                tempForm.num = row.form;
                tempForm.quantity = row.sheets;
                tempForm.hours = row.hours;

                if (row.day == "Monday") {

                    //organizes data to be used to build static HTML tables
                    createWeekTables(globalVar.monday, row, tempObj, tempForm);

                    //sorts organized data so the press numbers are in order
                    sortArrOfObj(globalVar.monday, "press");

                    //builds the static HTML table based on the data from the previous 2 functions
                    buildDOWTable(globalVar.monday, "Monday");

                    //keeps the week title the same if the selected dot hasn't changed and something else is updated in the tabulator tables
                    if ($(`#monday-static`).hasClass("dot-selected")) {
                        $("#week-title").text("MONDAY");
                        globalVar.dotSelected = true;
                    }

                } else if (row.day == "Tuesday") {

                    //organizes data to be used to build static HTML tables
                    createWeekTables(globalVar.tuesday, row, tempObj, tempForm);

                    //sorts organized data so the press numbers are in order
                    sortArrOfObj(globalVar.tuesday, "press");

                    //builds the static HTML table based on the data from the previous 2 functions
                    buildDOWTable(globalVar.tuesday, "Tuesday");

                    //keeps the week title the same if the selected dot hasn't changed and something else is updated in the tabulator tables
                    if ($(`#tuesday-static`).hasClass("dot-selected")) {
                        $("#week-title").text("TUESDAY");
                        globalVar.dotSelected = true;
                    }

                } else if (row.day == "Wednesday") {

                    //organizes data to be used to build static HTML tables
                    createWeekTables(globalVar.wednesday, row, tempObj, tempForm);

                    //sorts organized data so the press numbers are in order
                    sortArrOfObj(globalVar.wednesday, "press");

                    //builds the static HTML table based on the data from the previous 2 functions
                    buildDOWTable(globalVar.wednesday, "Wednesday");

                    //keeps the week title the same if the selected dot hasn't changed and something else is updated in the tabulator tables
                    if ($(`#wednesday-static`).hasClass("dot-selected")) {
                        $("#week-title").text("WEDNESDAY");
                        globalVar.dotSelected = true;
                    }

                } else if (row.day == "Thursday") {

                    //organizes data to be used to build static HTML tables
                    createWeekTables(globalVar.thursday, row, tempObj, tempForm);

                    //sorts organized data so the press numbers are in order
                    sortArrOfObj(globalVar.thursday, "press");

                    //builds the static HTML table based on the data from the previous 2 functions
                    buildDOWTable(globalVar.thursday, "Thursday");

                    //keeps the week title the same if the selected dot hasn't changed and something else is updated in the tabulator tables
                    if ($(`#thursday-static`).hasClass("dot-selected")) {
                        $("#week-title").text("THURSDAY");
                        globalVar.dotSelected = true;
                    }

                } else if (row.day == "Friday") {

                    //organizes data to be used to build static HTML tables
                    createWeekTables(globalVar.friday, row, tempObj, tempForm);

                    //sorts organized data so the press numbers are in order
                    sortArrOfObj(globalVar.friday, "press");

                    //builds the static HTML table based on the data from the previous 2 functions
                    buildDOWTable(globalVar.friday, "Friday");

                    //keeps the week title the same if the selected dot hasn't changed and something else is updated in the tabulator tables
                    if ($(`#friday-static`).hasClass("dot-selected")) {
                        $("#week-title").text("FRIDAY");
                        globalVar.dotSelected = true;
                    }

                } else {
                    // console.log("Don't know what to tell ya...")
                };

            });

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region ORGANIZE WEEK TABLE DATA ---------------------------------------------------------------------------------------------------------------

        /**
         * Organizes the dow data into an array of objects that can be referenced later to build a HTML table
         * @param {Array} dowArr An array that starts out empty but fills with the data for each item that belongs it that specific day of the week
         * @param {Object} rowData An object of all the data in the current row
         * @param {Object} tempObj An object that starts out empty but gets filled with data before it is sorted into it's proper merged or 
         * unmerged data set
         * @param {Object} tempForm An object that starts out empty but gets filled with data about the different forms before being moved 
         * into it's proper data set
         */
        function createWeekTables(dowArr, rowData, tempObj, tempForm) {

            //creates a duplicate of original array to be used for changing stuff
            var tempObjCopy = JSON.parse(JSON.stringify(tempObj));
            var tempFormCopy = JSON.parse(JSON.stringify(tempForm));


            if (dowArr.length == 0) { //if dowArr is empty, toss in the first row of data into a new object

                tempObjCopy.press = String(rowData.press);
                // tempObjCopy.operator = Array(rowData.operator);
                tempObjCopy.forms = Array(tempFormCopy);

            } else {

                //make sure the operator and forms feilds in the temp files are empty
                // tempObjCopy.operator = [];
                tempObjCopy.forms = [];

                let pressExist = false;
                let zeIndex;

                for (let i = 0; i < dowArr.length; i++) {

                    //if the press of the current row already exists in the dowArr data set, 
                    //store the index of said press in the dowArr and set pressExist to true
                    if (dowArr[i].press == rowData.press) {
                        zeIndex = i;
                        pressExist = true;
                    };
                };

                if (pressExist) {

                    //if the press exists AND the operator is already assinged to this press, just push the form number to the forms object 
                    //inside the same press object
                    // if (dowArr[zeIndex].operator.includes(rowData.operator)) {
                    //   dowArr[zeIndex].forms.push(tempFormCopy);

                    // //otherwise do the same but also push the operator into the operator array as well 
                    // } else {
                    // dowArr[zeIndex].operator.push(rowData.operator);
                    dowArr[zeIndex].forms.push(tempFormCopy);
                    // };

                    //reset variables for next go
                    pressExist = false;
                    zeIndex = "";

                    //if press doesn;t yet exisit in the data, just push all the data into a new object
                } else {
                    tempObjCopy.press = String(rowData.press);
                    // tempObjCopy.operator = Array(rowData.operator);
                    tempObjCopy.forms = Array(tempFormCopy);
                };

            };

            //if a new press object was created, push it into dowArr
            if (tempObjCopy.press !== undefined) {
                dowArr.push(tempObjCopy);
            };

            //reset variables for next go
            tempObj = {};
            tempForm = {};
            tempObjCopy = {};
            tempFormCopy = {};

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region ORGANIZE OPERATOR TABLE DATA -----------------------------------------------------------------------------------------------------------

        /**
         * Organizes the dow data into an array of objects that can be referenced later to build a HTML table
         * @param {Array} opArr An array that starts out empty but fills with the data for each item that belongs to a specific operator
         * @param {Object} rowData An object of all the data in the current row
         * @param {Object} tempObj An object that starts out empty but gets filled with data before it is sorted into it's proper merged or 
         * unmerged data set
         * @param {Object} tempForm An object that starts out empty but gets filled with data about the different forms before being moved 
         * into it's proper data set
         */
        function createOperatorTables(opArr, rowData, tempObj, tempForm, tempPress) {

            //creates a duplicate of original array to be used for changing stuff
            var tempObjCopy = JSON.parse(JSON.stringify(tempObj));
            var tempFormCopy = JSON.parse(JSON.stringify(tempForm));
            var tempPressCopy = JSON.parse(JSON.stringify(tempPress));

            if (opArr == undefined) {
                return;
            };


            if (opArr.length == 0) { //if opArr is empty, toss in the first row of data into a new object

                tempPressCopy[rowData.press] = Array(tempFormCopy);

                tempObjCopy.operator = rowData.operator;
                tempObjCopy.day = rowData.day;
                tempObjCopy.pressRef = [];
                tempObjCopy.press = Array(tempPressCopy);
                tempObjCopy.pressRef.push(rowData.press);
                // tempObjCopy.forms = Array(tempFormCopy);

            } else {

                //make sure the operator and forms feilds in the temp files are empty
                tempObjCopy.operator = [];
                // tempObjCopy.forms = [];

                let dayExist = false;
                let zeIndex;

                for (let i = 0; i < opArr.length; i++) {

                    //if the day of the current row already exists in the opArr data set, 
                    //store the index of said day in the opArr and set dayExist to true
                    if (opArr[i].day == rowData.day) {
                        zeIndex = i;
                        dayExist = true;
                    };
                };

                // console.log("break");
                // console.log("down");

                if (dayExist) {

                    // for (let key of opArr[zeIndex].press) {
                    //   if (key == rowData.press) {
                    //     console.log("It worked");
                    //   };
                    // };
                    //if the day exists AND the press is already assigned to this day, just push the form number to the forms object 
                    //inside the same press object
                    if (opArr[zeIndex].pressRef.includes(rowData.press)) {
                        // opArr[zeIndex].forms.push(tempFormCopy);
                        // tempPressCopy[rowData.press] = Array(tempFormCopy);
                        opArr[zeIndex].press[0][rowData.press].push(tempFormCopy);

                        opArr.sort((firstItem, secondItem) => firstItem[press[0]] - secondItem[press[0]]);

                        //otherwise do the same but also push the day into the operator array as well 
                    } else {
                        tempPressCopy[rowData.press] = Array(tempFormCopy);
                        opArr[zeIndex].press.push(tempPressCopy);
                        opArr[zeIndex].pressRef.push(rowData.press);

                        opArr.sort((firstItem, secondItem) => firstItem[press[0]] - secondItem[press[0]]);


                        // opArr[zeIndex].press.push(rowData.press);
                        // opArr[zeIndex].forms.push(tempFormCopy);
                    };

                    //reset variables for next go
                    dayExist = false;
                    zeIndex = "";

                    //if day doesn't yet exisit in the data, just push all the data into a new object
                } else {
                    // tempObjCopy = rowData.day;
                    // tempObjCopy.press = Array(rowData.press);
                    // tempObjCopy.operator = rowData.operator;
                    // tempObjCopy.forms = Array(tempFormCopy);

                    tempPressCopy[rowData.press] = Array(tempFormCopy);

                    tempObjCopy.operator = rowData.operator;
                    tempObjCopy.day = rowData.day;
                    tempObjCopy.pressRef = [];
                    tempObjCopy.press = Array(tempPressCopy);
                    tempObjCopy.pressRef.push(rowData.press);
                };

            };

            //if a new day object was created, push it into opArr
            if (tempObjCopy.day !== undefined) {
                opArr.push(tempObjCopy);
            };

            //reset variables for next go
            tempObj = {};
            tempForm = {};
            tempObjCopy = {};
            tempFormCopy = {};

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region SORT DATA ALPHABETICALLY ---------------------------------------------------------------------------------------------------------------

        /**
         * Sorts an array of objects alphabetically by a specific propery
         * @param {Array} arrOfObj The array containing objects that you wish to sort
         * @param {String} propToSort The object property that you wish to sort by
         */
        function sortArrOfObj(arrOfObj, propToSort) {
            // arrOfObj.sort((firstItem, secondItem) => firstItem[propToSort] - secondItem[propToSort]);
            arrOfObj.sort(function (firstItem, secondItem) {
                if (firstItem[propToSort] < secondItem[propToSort]) { return -1; }
                if (firstItem[propToSort] > secondItem[propToSort]) { return 1; }
                return 0;
            });
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region BUILD DAY OF THE WEEK TABLES -----------------------------------------------------------------------------------------------------------

        /**
         * Takes the organized data set and builds a static HTML table out of it
         * @param {Array} dow An array of objects containing all the data for the specific day of the week to use to build the static HTML tables
         * @param {String} dowString The name of the day of the week table
         */
        function buildDOWTable(dow, dowString) {

            let totalHours = [];

            let upperDOW = dowString.toUpperCase();
            let lowerDOW = dowString.toLowerCase();

            //empty and add dow table div to HTML
            $(`#${lowerDOW}-form`).empty().append(`<table id="${lowerDOW}-table"></table>`);

            //set up the columns and column headers in the table
            $(`#${lowerDOW}-table`)
                .addClass("html-table")
                .append(`
                <tr class="row-headers">
                <td class="dow-headers">PRESS</td>
                <td class="dow-headers">FORM</td>
                <td class="dow-headers">SHEET QUANTITY</td>
                <td class="dow-headers">HOURS</td>
                </tr>
            `)


            dow.forEach((press, pressIndex) => { //for each press in the dow data set...

                press.forms.forEach((form, formIndex) => { //for each form in each press in the dow data set...

                    //create row spans for press and operators based on the number of forms in the press data set
                    $(`#${lowerDOW}-table`).append(`
                <tr>
                    ${formIndex === 0
                            ? `<td class="press" rowspan="${press.forms.length}">${press.press}</td>`
                            : ``}
                    <td class="form">${form.num}</td>
                    <td class="quantity">${form.quantity}</td>
                    <td class="hours">${form.hours}</td>
                </tr>
                `);

                    //add each form hours as a new value in the total hours array
                    totalHours.push(form.hours);

                });

                //once each form is accounted for for the current press, take all values that are in the 
                //totalHours array and add them up
                let totalHoursNum = totalHours.reduce((a, b) => a + b, 0);
                totalHoursNum = Math.round(100 * totalHoursNum) / 100;

                //create a total hours row at the bottom of the press table
                $(`#${lowerDOW}-table`).append(`
                <tr class="total-row">
                <td class="total-hours" colspan="3">Press ${press.press} Total Hours:</td>
                <td class="total-hours-num">${totalHoursNum}</td>
                </tr>
        
            `);

                //add space below press section after total hours row ONLY IF this is not the last record in the dow array
                if (dow.length >= (pressIndex + 2)) {//+2 because pressIndex is zero-indexed and 
                    //I need this number to be 1 more than the current index
                    $(`#${lowerDOW}-table`).append(`
                <tr class="spacer-row">
                    <td class="separator"></td>
                </tr>
                `);
                };

                //reset variables for next go
                totalHours = [];

            });

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

// ---------------------------------------------------------------------------------------------------------------------------------------------------
// ---------------------------------------------------------------------------------------------------------------------------------------------------
/*

██    ██ ███    ██ ██    ██ ███████ ███████ ██████      ███████ ██    ██ ███    ██  ██████ ████████ ██  ██████  ███    ██ ███████ 
██    ██ ████   ██ ██    ██ ██      ██      ██   ██     ██      ██    ██ ████   ██ ██         ██    ██ ██    ██ ████   ██ ██      
██    ██ ██ ██  ██ ██    ██ ███████ █████   ██   ██     █████   ██    ██ ██ ██  ██ ██         ██    ██ ██    ██ ██ ██  ██ ███████ 
██    ██ ██  ██ ██ ██    ██      ██ ██      ██   ██     ██      ██    ██ ██  ██ ██ ██         ██    ██ ██    ██ ██  ██ ██      ██ 
 ██████  ██   ████  ██████  ███████ ███████ ██████      ██       ██████  ██   ████  ██████    ██    ██  ██████  ██   ████ ███████ 

*/

    //#region ----------------------------------------------------------------------------------------------------------------------------------------

        //============================================================================================================================================
            //#region MAKE TABLES FOR EACH OPERATOR --------------------------------------------------------------------------------------------------

                /**
                * Is feed either the Silk, Text, or Digital tabulator table data and organizes, sorts, and creates the static HTML tables
                * @param {Object} scheduleData An object containing all the data from the specific tabulator table to use in creating the 
                * static weekday tables
                */
                function makeTablesForEachOperator(scheduleData) {

                    let tempObj = {};
                    let tempForm = {};
                    let tempPress = {};

                    globalVar.opTableCounter = 1;

                    //if scheduleData is empty, remove all existing content from dow tables
                    if (scheduleData.length === 0) {

                        $("#monday-form").empty();
                        $("#tuesday-form").empty();
                        $("#wednesday-form").empty();
                        $("#thursday-form").empty();
                        $("#friday-form").empty();

                    };

                    scheduleData.forEach((row, rowIndex) => { //for each row of the data...

                        tempForm.num = row.form;
                        tempForm.quantity = row.sheets;
                        tempForm.hours = row.hours;

                        createOperatorTables(globalVar.operators[row.operator], row, tempObj, tempForm, tempPress);

                        console.log(JSON.stringify(globalVar.operators));

                        //builds the static HTML table based on the data from the previous function
                        buildOperatorTable(globalVar.operators, row.operator);

                        //keeps the week title the same if the selected dot hasn't changed and something else is updated in the tabulator tables
                        if ($(`#monday-static`).hasClass("dot-selected")) {
                            $("#week-title").text("MONDAY");
                            globalVar.dotSelected = true;
                        }

                    });

                };

            //#endregion -----------------------------------------------------------------------------------------------------------------------------
        //============================================================================================================================================

        //============================================================================================================================================
            //#region BUILD OPERATORS TABLE ----------------------------------------------------------------------------------------------------------

                //? Since this function was never used, I made the JSDOC way after the fact and cannot really remember how it worked. So some of the 
                //? param descriptions may not be accurate. But who cares really? This function will never be used, but for some reason 
                //? I cannot see it removed...
                /**
                 * Builds a static HTML table based on the data from the tabulator table
                 * @param {Object} operatorData An object containing the data from the tabulator table?
                 * @param {String} operatorName The name of the operator that is having a table made for them?
                 * @returns 
                 */
                function buildOperatorTable(operatorData, operatorName) {

                    let upperOp = operatorName.toUpperCase();
                    let lowerOp = operatorName.toLowerCase();

                    if (operatorName == "-") {
                        return;
                    }

                    $(`#html-tables`).append(`
                    <div id="${lowerOp}-static" class="current-week">
                    <div id="${lowerOp}-form" class="fill-out">
                        <table id="${lowerOp}-table" class="html-table">
                        <tr class="row-headers">
                            <td class="dow-headers">DAY</td>
                            <td class="dow-headers">PRESS</td>
                            <td class="dow-headers">FORM</td>
                            <td class="dow-headers">SHEET QUANTITY</td>
                            <td class="dow-headers">HOURS</td>
                        </tr>
                        </table>
                    </div>
                    </div>
                `);

                    console.log(operatorData[operatorName]);

                    operatorData[operatorName].forEach((weekday, weekdayIndex) => {

                        weekday.press.forEach((lePress, lePressIndex) => {

                            for (const [key] of Object.entries(lePress)) {

                                lePress[key].forEach((form, formIndex) => {

                                    console.log(form);

                                    //create row spans for press and operators based on the number of forms in the press data set
                                    $(`#${lowerOp}-table`).append(`
                            <tr>

                            ${formIndex === 0
                                            ? `<td class="day" rowspan="${form.length}">${weekday.day}</td>
                                    <td class="press" rowspan="${form.length}">${key}</td>`
                                            : ``}
                                <td class="form">${form.num}</td>
                                <td class="quantity">${form.quantity}</td>
                                <td class="hours">${form.hours}</td>
                            </tr>
                            `);

                                });

                            };

                        });

                    });

                    if (globalVar.opTableCounter == 1) {
                        $(`#${lowerOp}-static`).addClass("show-table");
                    };

                    globalVar.opTableCounter++;

                };

            //#endregion -----------------------------------------------------------------------------------------------------------------------------
        //============================================================================================================================================

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------

// ---------------------------------------------------------------------------------------------------------------------------------------------------
// ---------------------------------------------------------------------------------------------------------------------------------------------------


export { buildTabulatorTables, organizeData };