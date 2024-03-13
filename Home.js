/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

//#region GLOBAL VARIABLES ---------------------------------------------------------------------------------------------------------------------------


//#region GLOBAL UNDEFINED VARIABLES ---------------------------------------------------------------------------------------------------------------

let silkLayouts = [];
let sheetHourData = {};
let breakoutData = {};
let productData = {};
let wasteData = {};
let linesData = [];
let formToCarry;
let changeEvent;
let eventResult;
let formsColumnIndex;
let result = "";

let silkTable = undefined;
let textTable = undefined;
let digTable = undefined;
let monday = [];
let tuesday = [];
let wednesday = [];
let thursday = [];
let friday = [];

let silkDataSet = [];
let textDataSet = [];
let digDataSet = [];


let operators = {};
// let currentPage;
let scrollHeight;

let listOfBreakoutTables = [];

let emptySheets = [];

let normalBreakoutsFormatting = {};




//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region GLOBAL DEFINED VARIABLES -----------------------------------------------------------------------------------------------------------------

const hiddenLinesData = ["MISSING", "IGNORE", "Shipping", "Empty", "PRINTED", "DIGITAL"];
let priorityNum = 1;
let singleSided = false;
let dotSelected = false;
let plannedSep = false;
let rushItem = false;
let opTableCounter = 1;

let rowMoved = false;

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region PRE-DEFINED VARIABLES (PROPBABLY MOVE THESE TO TABLES IN THE VALIDATION) -----------------------------------------------------------------

// var pressmen = [" ", "Steve", "Roberto", "Ryan", "Jamie", "Cody", "Terry", "Paul"];
// var presses = [" ", 1, 2, 3, 4, "Digital 1", "Digital 2"];

let pressmen = [" "];
let presses = [" "];
var daysOfWeek = [" ", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];



//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#endregion -----------------------------------------------------------------------------------------------------------------------------------------


// $("#handle").resizable({
//   also
// })

let scrollErr;


//#region ON READY -----------------------------------------------------------------------------------------------------------------------------------

Office.onReady((info) => {

    if (info.host === Office.HostType.Excel) {

        scrollErr = document.querySelector("#select-tables");

        document.getElementById("populate-forms").onclick = populateForms;
        // document.getElementById("remove-breakout-sheets").onclick = removeBreakoutSheetsTaskpane;
        // document.getElementById("clear-forms").onclick = clearForms;
        document.getElementById("breakout").onclick = breakout;
        // document.getElementById("create-press-table-btn").onclick = refreshPivotTable;
        // const scrollErr = document.querySelector("#select-tables");
        document.getElementById("clearall-butt").onclick = clearForms;

        scrollErr.addEventListener("scroll", (event) => {
            // console.log("Scroll changed.", event)
            scrollHeight = scrollErr.scrollTop;
        });

        // document.getElementById("clear-press-schedule-info").onclick = clearPSInfo;
    };


    Excel.run(async (context) => {

        // console.log("I ran!");

        // if (currentPage == "Home") {

        //   $("#select-tables").css("display", "none");
        //   $("#week-tables").css("display", "none");
        //   $("#footer").css("display", "none");

        //   $("#home-page").css("display", "flex");

        // };

        // if (currentPage == "Press Scheduling") {

        //   $("#home-page").css("display", "none");

        //   $("#select-tables").css("display", "block");
        //   $("#week-tables").css("display", "flex");
        //   $("#footer").css("display", "flex");

        // };

        //#region PLACEHOLDER DATA FOR PLAYGROUND ------------------------------------------------------------------------------------------------------

        // silkDataSet = [
        //     {id:1, type: "Silk", form:1, day:"Monday", press:"1", operator:"Todd"},
        //     {id:2, type: "Silk", form:2, day:"Tuesday", press:"1", operator:"Todd"},
        //     {id:3, type: "Silk", form:3, day:"Monday", press:"2", operator:"Matt"},
        //     {id:4, type: "Silk", form:4, day:"Tuesday", press:"2", operator:"Matt"},
        //     {id:5, type: "Silk", form:5, day:"Monday", press:"1", operator:"Todd"},
        //     {id:6, type: "Silk", form:6, day:"Tuesday", press:"1", operator:"Todd"},
        //     {id:7, type: "Silk", form:7, day:"Monday", press:"2", operator:"Matt"},
        //     {id:8, type: "Silk", form:8, day:"Tuesday", press:"2", operator:"Matt"}
        // ]

        // textDataSet = [
        //     {id:1, type: "Text", form:51, day:"Tuesday", press:"1", operator:"Todd"},
        //     {id:2, type: "Text", form:52, day:"Monday", press:"1", operator:"Todd"},
        //     {id:3, type: "Text", form:53, day:"Tuesday", press:"2", operator:"Matt"},
        //     {id:4, type: "Text", form:54, day:"Monday", press:"2", operator:"Matt"}
        // ]

        // digDataSet = [
        //     {id:1, type: "Text", form:600, day:"Tuesday", press:"1", operator:"Todd"},
        //     {id:2, type: "Text", form:601, day:"Monday", press:"1", operator:"Todd"},
        //     {id:3, type: "Text", form:602, day:"Tuesday", press:"2", operator:"Matt"},
        //     {id:4, type: "Text", form:603, day:"Monday", press:"2", operator:"Matt"}
        // ]

        //#endregion -----------------------------------------------------------------------------------------------------------------------------------

        const silkE2RSheet = context.workbook.worksheets.getItem("SilkE2R").load("name");
        const textE2RSheet = context.workbook.worksheets.getItem("TextE2R").load("name");
        const digE2RSheet = context.workbook.worksheets.getItem("DIGE2R").load("name");
        const validationSheet = context.workbook.worksheets.getItem("Validation").load("name");
        const pressSchedulingSheet = context.workbook.worksheets.getItem("Press Scheduling").load("name");
        const masterSheet = context.workbook.worksheets.getItem("Master").load("name");
        const allSheets = context.workbook.worksheets.load("items");


        const silkE2RTable = silkE2RSheet.tables.getItem("SilkE2R");
        const textE2RTable = textE2RSheet.tables.getItem("TextE2R");
        const digE2RTable = digE2RSheet.tables.getItem("DIGE2R");
        const pressSchedulingInfo = validationSheet.tables.getItem("PressSchedulingInfo");
        const dowSummaryPivotTable = pressSchedulingSheet.pivotTables.getItem("DOWSummaryPivot");
        const press1PivotTable = pressSchedulingSheet.pivotTables.getItem("Press1Pivot");
        const press2PivotTable = pressSchedulingSheet.pivotTables.getItem("Press2Pivot");
        const press3PivotTable = pressSchedulingSheet.pivotTables.getItem("Press3Pivot");
        const digitalPivotTable = pressSchedulingSheet.pivotTables.getItem("DigitalPivot");
        const masterTable = masterSheet.tables.getItem("Master");
        const customFormsTable = validationSheet.tables.getItem("CustomForms").load("name");
        const pressmenTable = validationSheet.tables.getItem("Pressmen").load("name");
        const pressesTable = validationSheet.tables.getItem("Presses").load("name");


        const silkE2RBodyRangeUpdate = silkE2RTable.getDataBodyRange().load("values");
        const textE2RBodyRangeUpdate = textE2RTable.getDataBodyRange().load("values");
        const digE2RBodyRangeUpdate = digE2RTable.getDataBodyRange().load("values");
        const pressSchedulingBodyRange = pressSchedulingInfo.getDataBodyRange().load("values");
        const masterBodyRange = masterTable.getDataBodyRange().load("values");
        const customFormsBodyRange = customFormsTable.getDataBodyRange().load("values");
        const pressmenBodyRange = pressmenTable.getDataBodyRange().load("values");
        const pressesBodyRange = pressesTable.getDataBodyRange().load("values");


        const masterTableHeader = masterTable.getHeaderRowRange().load("values");


        await context.sync();

        let silkE2RArr = silkE2RBodyRangeUpdate.values; //moves all values of the SilkE2R table to an array
        let textE2RArr = textE2RBodyRangeUpdate.values; //moves all values of the TextE2R table to an array
        let digE2RArr = digE2RBodyRangeUpdate.values; //moves all values of the TextE2R table to an array
        let pressSchedArr = pressSchedulingBodyRange.values;
        let pressmenArr = pressmenBodyRange.values;
        let pressesArr = pressesBodyRange.values;
        let masterArr = masterBodyRange.values;
        let masterArrCopy = JSON.parse(JSON.stringify(masterBodyRange.values));

        let missingExist = false;

        for (let i = 0; i < allSheets.items.length; i++) {
            // console.log(allSheets.items[i].name);
            if (allSheets.items[i].name == "MISSING") {
                missingExist = true;
            };
        };

        if (missingExist) {
            $("#breakout").text("Delete Breakout Sheets");
        } else {
            $("#breakout").text("Breakout");
        };


        //#region WRITE PRESSMEN AND PRESSES TABLE INFO TO ARRAYS --------------------------------------------------------------------------------------

        pressmen = [" "];
        presses = [" "];

        pushBodyRangeValuesToArray(pressmenArr, pressmen);
        pushBodyRangeValuesToArray(pressesArr, presses);

        // for (let pressman of pressmenArr) {
        //   window[pressman] = + [];
        // }

        // let prop1 = {
        //   beans: "pickled",
        //   rice: "white"
        // };

        // let prop2 = {
        //   beans: "fart",
        //   rice: "none"
        // };

        // // console.log(Steve);

        // for (const key of pressmen) {
        //   operators[key] = prop1
        // };

        // console.log(operators);

        pressmen.forEach((man) => {
            operators[man] = [];
        });

        //#endregion -----------------------------------------------------------------------------------------------------------------------------------


        let masterHeaderValues = masterTableHeader.values;

        dowSummaryPivotTable.refreshOnOpen = true;
        press1PivotTable.refreshOnOpen = true;
        press2PivotTable.refreshOnOpen = true;
        press3PivotTable.refreshOnOpen = true;
        digitalPivotTable.refreshOnOpen = true;

        silkDataSet = [];
        textDataSet = [];
        digDataSet = [];
        priorityNum = 1;

        silkDataSet = createDataSet(silkE2RArr, "Silk");
        textDataSet = createDataSet(textE2RArr, "Text");
        digDataSet = createDataSet(digE2RArr, "Digital");

        priorityNum = 1;

        updateDataFromTable(pressSchedArr);

        await context.sync();


        silkTable = buildTabulatorTables("silk-form", silkTable, silkDataSet);
        textTable = buildTabulatorTables("text-form", textTable, textDataSet);
        digTable = buildTabulatorTables("dig-form", digTable, digDataSet);

        // await context.sync();


        // silkTable.setData(silkDataSet);
        // textTable.setData(textDataSet);
        // digTable.setData(digDataSet);

        organizeData();

        pressSchedulingInfo.onChanged.add(pressSchedulerHandler);

        await context.sync();

        scrollErr.scrollTop = scrollHeight;


        silkE2RTable.onChanged.add(E2RHandler);
        textE2RTable.onChanged.add(E2RHandler);
        digE2RTable.onChanged.add(E2RHandler);





        var hasChildDiv = document.getElementById("silk-form").querySelector(".tabulator-header");

        let hasChildDiv2;

        let tabTableData = false;

        if (hasChildDiv !== null) {

            // console.log("tabulator-header exists!");

            let tabulatorTableElements = document.getElementsByClassName("tabulator-table");

            for (let eachTableElement of tabulatorTableElements) {

                hasChildDiv2 = eachTableElement.querySelector(".tabulator-row");

                // console.log(hasChildDiv2);

                if (hasChildDiv2 !== null) {
                    // console.log("The tabulator tables should have data in them!");
                    tabTableData = true;
                    break;
                }
            }

            if (hasChildDiv2 == null) {
                console.log("DOUBLE NOOO!");
                tabTableData = false;
            };

        } else {
            console.log('NO');
            tabTableData = false;
        };

        if (tabTableData == false) {

            //need to hide the select table and week tables here and show another div
            // $("#select-tables").css("animation", "hide-div 5s");
            // $("#week-tables").css("animation", "hide-div 5s");
            // $("#header").css("animation", "hide-div 5s");

            // $("#welcome-page").css("animation", "show-div 5s");

            setTimeout(swapToWelcome, 5);

        } else {

            // $("#select-tables").css("animation", "show-div 5s");
            // $("#week-tables").css("animation", "show-div 5s");
            // $("#header").css("animation", "show-div 5s");

            // $("#welcome-page").css("animation", "hide-div 5s");

            setTimeout(swapToShceduling, 5);

        }

        let emptyCell = false;

        for (let cell of masterArr[0]) {
            if (cell == "") {
                emptyCell = true;
            } else {
                emptyCell = false;
                break;
            }
        };


        function swapToShceduling() {
            $("#select-tables").css("display", "block");
            $("#week-tables").css("display", "flex");
            $("#header").css("display", "block");

            $("#welcome-page").css("display", "none");
        };

        function swapToWelcome() {
            $("#select-tables").css("display", "none");
            $("#week-tables").css("display", "none");
            $("#header").css("display", "none");

            $("#welcome-page").css("display", "flex");
        };


        // if (masterArr.length === 1 && emptyCell == true) {
        //   console.log("Master is empty, so no events were bound to the forms column just yet.")
        // } else {
        //   // activateEvents();
        //   // eventResult = masterTable.onChanged.add(handleChange);

        //   for (let rowIndex in masterArr) {

        //     let masterRowInfo = new Object();

        //     for (let headName of masterHeaderValues[0]) {
        //       createRowInfo(masterHeaderValues, headName, masterArr[rowIndex], masterArrCopy, masterRowInfo, rowIndex);
        //     };

        //     console.log(masterRowInfo.Forms.value);

        //   }
        // };





        //keeping these next 3 on change calls within the scope of the onReady so that the tabulator table variable I defined above are available to use

        //#region ON TABLE CHANGE, UPDATE HTML TABLES --------------------------------------------------------------------------------------------------


        //#region ON SILK TABLE CHANGE, UPDATE DATA AND REBUILD HTML TABLES --------------------------------------------------------------------------

        $("#silk-form").on("change", ".select-box", function () {

            deactivateEvents();

            const whichRow = $(this).attr("rowId");
            const whichColumn = $(this).attr("colId");
            const newData = $(this).find(":selected").text();

            const rowForm = $(this).attr("formNum");

            const tableData = silkTable.getData();

            //matches the index + 1 (for 0 index) to the rowId of the changed table and then replaces the data in the column with the newData
            const replaceData = tableData.map((td, index) => {
                // console.log(index + 1, whichRow)
                if ((index + 1) + "" === whichRow) {
                    // console.log("NEW DATA")
                    td[whichColumn] = newData
                }
                return td;
            });

            silkTable.setData(replaceData); //sets data in tabulator table to new data
            scrollErr.scrollTop = scrollHeight;

            silkDataSet = replaceData;

            organizeData();

            pressSchedulingInfoTable("Taskpane");

            updateE2RFromTaskpane("Silk", rowForm);


            console.log("A silk select box was changed in the Taskpane, so values in the Press Scheduling Info table and in the SilkE2R were updated");



        });

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------


        //#region ON TEXT TABLE CHANGE, UPDATE DATA AND REBUILD HTML TABLES --------------------------------------------------------------------------

        $("#text-form").on("change", ".select-box", function () {

            deactivateEvents();

            const whichRow = $(this).attr("rowId");
            const whichColumn = $(this).attr("colId");
            const newData = $(this).find(":selected").text();

            const rowForm = $(this).attr("formNum");

            const tableData = textTable.getData();

            const replaceData = tableData.map((td, index) => {
                // console.log(index + 1, whichRow)
                if ((index + 1) + "" === whichRow) {
                    console.log("NEW DATA")
                    td[whichColumn] = newData
                }
                return td;
            });

            textTable.setData(replaceData);
            // textTable.replaceData(replaceData);
            scrollErr.scrollTop = scrollHeight;

            textDataSet = replaceData;

            organizeData();

            pressSchedulingInfoTable("Taskpane");

            updateE2RFromTaskpane("Text", rowForm);

            console.log("A text select box was changed in the Taskpane, so values in the Press Scheduling Info table and in the TextE2R were updated");

        });

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------


        //#region ON DIGITAL TABLE CHANGE, UPDATE DATA AND REBUILD HTML TABLES -----------------------------------------------------------------------

        $("#dig-form").on("change", ".select-box", function () {

            deactivateEvents();

            console.log("Digital Select Box was Updated");

            const whichRow = $(this).attr("rowId");
            const whichColumn = $(this).attr("colId");
            const newData = $(this).find(":selected").text();

            const rowForm = $(this).attr("formNum");

            const tableData = digTable.getData();

            const replaceData = tableData.map((td, index) => {
                // console.log(index + 1, whichRow)
                if ((index + 1) + "" === whichRow) {
                    console.log("NEW DATA")
                    td[whichColumn] = newData
                }
                return td;
            });

            // digTable.setData(replaceData);

            digTable.setData(replaceData);

            scrollErr.scrollTop = scrollHeight;

            digDataSet = replaceData;

            organizeData();

            pressSchedulingInfoTable("Taskpane");

            updateE2RFromTaskpane("Digital", rowForm);

            console.log("A digital select box was changed in the Taskpane, so values in the Press Scheduling Info table and in the DIGE2R were updated");

        });

        //#endregion ---------------------------------------------------------------------------------------------------------------------------------


        eventResult = masterTable.onChanged.add(handleChange);


        //#endregion -----------------------------------------------------------------------------------------------------------------------------------

    });

});


async function updateE2RFromTaskpane(type, rowForm) {

    await Excel.run(async (context) => {

        const silkE2RSheet = context.workbook.worksheets.getItem("SilkE2R").load("name");
        const textE2RSheet = context.workbook.worksheets.getItem("TextE2R").load("name");
        const digE2RSheet = context.workbook.worksheets.getItem("DIGE2R").load("name");

        const silkE2RTable = silkE2RSheet.tables.getItem("SilkE2R");
        const textE2RTable = textE2RSheet.tables.getItem("TextE2R");
        const digE2RTable = digE2RSheet.tables.getItem("DIGE2R");

        const silkE2RBodyRangeUpdate = silkE2RTable.getDataBodyRange().load("values");
        const textE2RBodyRangeUpdate = textE2RTable.getDataBodyRange().load("values");
        const digE2RBodyRangeUpdate = digE2RTable.getDataBodyRange().load("values");

        await context.sync();

        let silkE2RArr = silkE2RBodyRangeUpdate.values;
        let textE2RArr = textE2RBodyRangeUpdate.values;
        let digE2RArr = digE2RBodyRangeUpdate.values;




        if (type == "Silk") {
            let newSilkArr = updateDayPressInE2RArr(silkDataSet, silkE2RArr, rowForm);
            silkE2RBodyRangeUpdate.values = newSilkArr;
        };

        if (type == "Text") {
            let newTextArr = updateDayPressInE2RArr(textDataSet, textE2RArr, rowForm);
            textE2RBodyRangeUpdate.values = newTextArr;
        };

        if (type == "Digital") {
            let newDigArr = updateDayPressInE2RArr(digDataSet, digE2RArr, rowForm);
            digE2RBodyRangeUpdate.values = newDigArr;
        };

    });
};


function updateDayPressInE2RArr(dataSet, arr, rowForm) {

    let dayUpdate;
    let pressUpdate;
    let found = false;

    dataSet.forEach((dataRow) => {
        if (dataRow.form == rowForm) {
            dayUpdate = dataRow.day;
            pressUpdate = dataRow.press;
            found = true;
        };
    });

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

//#endregion -----------------------------------------------------------------------------------------------------------------------------------------

//#region ON BUTTON CLICKS ---------------------------------------------------------------------------------------------------------------------------


//#region NAVIGATION ARROWS ------------------------------------------------------------------------------------------------------------------------


//#region PAGE ARROWS ----------------------------------------------------------------------------------------------------------------------------

//goes between main pages
$(".title-arrow").on("click", function () {

    const thisPage = $("#page-title").text().toLowerCase();

    // Are we going forward or backward
    const forwardOrBackward = $(this).attr("value");

    let pages = ["home", "press scheduling"];

    const next = nextPage(thisPage, forwardOrBackward, pages);

    let thisPageWithDash = replaceCharacter(thisPage, " ", "-");
    let nextWithDash = replaceCharacter(next, " ", "-");


    $("#page-title").text(next.toUpperCase());

    $(`.dots[page='${thisPageWithDash}']`).removeClass("dot-selected");
    $(`.dots[page='${nextWithDash}']`).addClass("dot-selected");

    if (thisPage == "press scheduling" && next == "home") {

        $("#select-tables").css("display", "none");
        $("#week-tables").css("display", "none");
        $("#footer").css("display", "none");

        $("#home-page").css("display", "flex");

        // currentPage = "Home";

    };

    if (thisPage == "home" && next == "press scheduling") {

        $("#home-page").css("display", "none");

        $("#select-tables").css("display", "block");
        $("#week-tables").css("display", "flex");
        $("#footer").css("display", "flex");

        // currentPage = "Press Scheduling"

    };

});

//#endregion -------------------------------------------------------------------------------------------------------------------------------------


//#region DAY OF WEEK ARROWS ---------------------------------------------------------------------------------------------------------------------


//go between weekday tables in the week-tables div
$(document).on("click", ".dow-arrow", function () {
    // Get this weekday
    const thisWeekday = $("#week-title").text().toLowerCase();

    // Are we going forward or backward
    const forwardOrBackward = $(this).attr("value");

    // Hide this week's table
    // $(`#${thisWeekday}-static`).css("display", "none");
    $(`#${thisWeekday}-static`).removeClass("show-table");

    // Days of the week
    let wd = ["monday", "tuesday", "wednesday", "thursday", "friday"]

    // Get the next or previous day in the week
    const nextDay = nextPage(thisWeekday, forwardOrBackward, wd)

    // Change the header text to match next day
    $("#week-title").text(nextDay.toUpperCase());

    // Update the dots
    $(`.dots[weekday='${thisWeekday}']`).removeClass("dot-selected");
    $(`.dots[weekday='${nextDay}']`).addClass("dot-selected");

    // Show the next day's table
    // $(`#${nextDay}-static`).css("display", "block");
    $(`#${nextDay}-static`).addClass("show-table");

});

//#endregion -------------------------------------------------------------------------------------------------------------------------------------


//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region NAVIGATION DOTS --------------------------------------------------------------------------------------------------------------------------

//when navigation dots are clicked, change the page/table to the proper dot's page/table
$(document).on("click", ".dots", function () {

    const element = $(this);
    let leWeekday = element.attr("weekday"); //returns a weekday value if nav-dots in week-tables were clicked, otherwise undefined
    let leTitle = element.attr("page"); //returns a page value if nav-dots in title-header were clicked, otherwise undefined

    let dotSet;

    //sets dotSet to either "Title" or "Weekday" depending on which set of nav-dots were clicked
    if (!leWeekday) {
        if (!leTitle) {
            console.log("The dot you click doesn't exist...");
        } else {
            dotSet = "Title";
        };
    } else {
        dotSet = "Weekday";
    };


    //For week-tables nav-dots. Sets the header text to the new table title, sets the dot-selected class to the proper dot, 
    //and shows the proper table
    if (dotSet == "Weekday") {

        //sets table title
        $("#week-title").text(leWeekday.toUpperCase());

        //removes the dot-selected class from all week-dots class items and adds it to the proper element
        $(".week-dots").removeClass("dot-selected");
        $(element).addClass("dot-selected");

        //removes show-table class from all current-week class items and sets it to the proper element
        $(".current-week").removeClass("show-table");
        $(`#${leWeekday}-static`).addClass("show-table");

    };

    //For page-title nav-dots. Sets the header text to the new page title, sets the dot-selected class to the proper dot, 
    //and shows the proper page elements
    if (dotSet == "Title") {

        //gives us a version the page title with spaces instead of dashes
        let leTitleSpace = replaceCharacter(leTitle, "-", " ");

        //sets page title
        $("#page-title").text(leTitleSpace.toUpperCase());

        //removes the dot-selected from all title-dots class items and adds it to the proper element
        $(".title-dots").removeClass("dot-selected");
        $(element).addClass("dot-selected");

        //if changing to the "press-scheduling" page, hide all home elements and show all press scheduling elements
        if (leTitle == "press-scheduling") {

            $("#home-page").css("display", "none");

            $("#select-tables").css("display", "block");
            $("#week-tables").css("display", "flex");
            $("#footer").css("display", "flex");

            // currentPage = "Press Scheduling";

            //if changing to the "home" page, hide all press scheduling elements and show all home elements
        } else if (leTitle == "home") {

            $("#select-tables").css("display", "none");
            $("#week-tables").css("display", "none");
            $("#footer").css("display", "none");

            $("#home-page").css("display", "flex");

            // currentPage = "Home";

        };

    };

});

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region DOUBLE-CLICK EVENTS ----------------------------------------------------------------------------------------------------------------------



//#region DOUBLE-CLICK HANDLE TO RESIZE WEEK-TABLE DIV TO TABLE HEIGHT ---------------------------------------------------------------------------


//adjusts the height of the week-tables div when the user double-clicks the handle to fit the height of the table, 
//handle element, and header element
$("#handle").on("dblclick", function () {
    console.log("Double clicked...");

    let h = 0;
    let hh = 0;
    let wth = 0;

    h = $(".show-table").height();
    hh = $("#handle").height();
    wth = $("#week-table-head").height();

    $("#week-tables").animate({
        height: `${h + hh + wth + 37}px`
    }, 250)
});

//#endregion -------------------------------------------------------------------------------------------------------------------------------------



// $("#header").dblclick(() => {
//$(".border-box").slideUp();
// })

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region OTHER BOUND EVENTS -----------------------------------------------------------------------------------------------------------------------


//#region MAKES WEEK-TABLES DIV RESIZABLE --------------------------------------------------------------------------------------------------------

//makes the week-tables div user resizable
$("#week-tables").resizable({
    handles: { "n": "#handle" },
    stop: function (event, ui) {
        // ui.element.width("");
    }
});

//#endregion -------------------------------------------------------------------------------------------------------------------------------------


//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


async function clearPSInfo() {

    await Excel.run(async (context) => {

        deactivateEvents();

        const validation = context.workbook.worksheets.getItem("Validation");
        const pressSchedulingInfo = validation.tables.getItem("PressSchedulingInfo");
        const pressSchedulingInfoRows = pressSchedulingInfo.rows.load("count");

        await context.sync();


        let rowCount = pressSchedulingInfoRows.count - 1;

        pressSchedulingInfoRows.deleteRowsAt(0, rowCount);

        await context.sync();

        pressSchedulingInfoRows.getItemAt(0).delete();

        await context.sync();

        activateEvents();

        location.reload();

    });

};





//#region ON POPULATE FORMS BUTTON CLICK -----------------------------------------------------------------------------------------------------------

async function populateForms() {

    deactivateEvents();

    $("#loading-background").css("display", "flex");

    $("#loading-head").text("Populating Forms");


    try {

        await Excel.run(async (context) => { //loads context so I can directly pull and write info to Excel within this async function


            //#region POPULATE E2R INFO ----------------------------------------------------------------------------------------------------------------


            //#region DEFINE FUNCTION VARIABLES ------------------------------------------------------------------------------------------------------

            //load in worksheets
            // const range = context.workbook.getSelectedRange(); //selected range; left over from inital example
            const validation = context.workbook.worksheets.getItem("Validation");
            const customFormsTable = validation.tables.getItem("CustomForms").load("name");
            const linesTable = validation.tables.getItem("Lines");
            const linesBodyRange = linesTable.getDataBodyRange().load("values");
            const linesHeaderRange = linesTable.getHeaderRowRange().load("values");
            const silkE2RSheet = context.workbook.worksheets.getItem("SilkE2R").load("name");
            const textE2RSheet = context.workbook.worksheets.getItem("TextE2R").load("name");
            const digE2RSheet = context.workbook.worksheets.getItem("DIGE2R").load("name");
            const masterSheet = context.workbook.worksheets.getItem("Master").load("name");

            //load in tables
            const sheetsPerHourTable = validation.tables.getItem("SheetsPerHour");
            const wasteTable = validation.tables.getItem("Waste");
            const productsTable = validation.tables.getItem("Products");
            const apparelTable = validation.tables.getItem("ApparelTable");
            // const pressSchedulingInfo = validation.tables.getItem("PressSchedulingInfo");
            const silkE2RTable = silkE2RSheet.tables.getItem("SilkE2R");
            const textE2RTable = textE2RSheet.tables.getItem("TextE2R");
            const digE2RTable = digE2RSheet.tables.getItem("DIGE2R");
            const masterTable = masterSheet.tables.getItem("Master");
            const defaultToIgnoreTable = validation.tables.getItem("DefaultToIgnore");


            //loads the data body range of the tables above
            const sheetHourBodyRange = sheetsPerHourTable.getDataBodyRange().load("values");
            const customFormsBodyRange = customFormsTable.getDataBodyRange().load("values");
            const wasteBodyRange = wasteTable.getDataBodyRange().load("values");
            const productsBodyRange = productsTable.getDataBodyRange().load("values");
            const apparelBodyRange = apparelTable.getDataBodyRange().load("values");
            const silkE2RBodyRange = silkE2RTable.getDataBodyRange().load("values");
            const textE2RBodyRange = textE2RTable.getDataBodyRange().load("values");
            const digE2RBodyRange = digE2RTable.getDataBodyRange().load("values");
            const masterBodyRange = masterTable.getDataBodyRange().load("values");
            const defaultToIgnoreBodyRange = defaultToIgnoreTable.getDataBodyRange().load("values");


            //loads the headers of the validation tables
            const sheetHourHeaderRange = sheetsPerHourTable.getHeaderRowRange().load("values");
            const wasteHeaderRange = wasteTable.getHeaderRowRange().load("values");
            const productsHeaderRange = productsTable.getHeaderRowRange().load("values");
            const masterHeaderRange = masterTable.getHeaderRowRange().load("values");

            //loads the row items for specific tables
            const silkE2RTableRows = silkE2RTable.rows.load("items");
            const textE2RTableRows = textE2RTable.rows.load("items");
            const digE2RTableRows = digE2RTable.rows.load("items");
            const masterTableRows = masterTable.rows.load("items");

            //#endregion -----------------------------------------------------------------------------------------------------------------------------


            await context.sync();


            //#region LOAD FUNCTION VARIABLES --------------------------------------------------------------------------------------------------------

            //load in table arrays from table data
            let sheetHourArr = sheetHourBodyRange.values;
            let wasteArr = wasteBodyRange.values;
            let productArr = productsBodyRange.values;
            let apparelArr = apparelBodyRange.values;
            let silkE2RArr = silkE2RBodyRange.values; //moves all values of the SilkE2R table to an array
            let textE2RArr = textE2RBodyRange.values; //moves all values of the TextE2R table to an array
            let digE2RArr = digE2RBodyRange.values;
            let masterArr = masterBodyRange.values;

            //load in header values from table data
            let sheetHourHeader = sheetHourHeaderRange.values;
            let wasteHeader = wasteHeaderRange.values;
            let productHeader = productsHeaderRange.values;
            let masterHeader = masterHeaderRange.values;

            //range of items from the table row data
            let masterRowItems = masterTableRows.items;

            //shallow copy of table array for doing calculations to
            let masterArrCopy = JSON.parse(JSON.stringify(masterBodyRange.values));

            let defaultToIgnoreArr = defaultToIgnoreBodyRange.values;


            //#endregion -----------------------------------------------------------------------------------------------------------------------------


            //#region CREATE DATA OBJECTS FOR EASIER REFERENCING IN LATER CODE -----------------------------------------------------------------------

            //#region CREATE COLUMN HEADER OBJECTS WITH ALL VALUES ASSIGNED TO THEM FROM SHEET HOURS TABLE -----------------------------------------

            var x = 0

            let tempArr = []; //temporary array for storing values to push into the data object (clears each time)

            let valueToPush;

            for (var item of sheetHourHeader[0]) { //for each header cell in thr sheet hour table...
                for (var y = 0; y < sheetHourArr.length; y++) { //y will never be larger than the number of rows in the sheet hour table

                    //the last maximum value should typically be everything after the last minimum number, or infinity. In the table, this is typically 
                    //expressed by a "+", "-", or just an empty cell. This translates that into the javascript equivilant for infinity
                    if ((sheetHourArr[y][x] == "-" || sheetHourArr[y][x] == "+" || sheetHourArr[y][x] == "") && sheetHourArr[y][x] !== 0) {
                        valueToPush = Infinity;
                    } else {
                        valueToPush = sheetHourArr[y][x]; //if not infinity, then just use the current value without translation
                    };

                    tempArr.push(valueToPush); //pushes the xth cell in row y to the tempArr, then repeats for the same cell in each row

                };

                sheetHourData[item] = tempArr; //assigns all the values for each column header to itself as an object in sheetHourData

                tempArr = []; //clears the tempArr so it can be clean as we loop through the parent for loop

                x = x + 1;

            };

            //#endregion ---------------------------------------------------------------------------------------------------------------------------

            //#region CREATE PRODUCT DATA OBJECT ---------------------------------------------------------------------------------------------------

            let tempProdObj = {}; //temporary object for storing values to push into the data object (clears each time)

            for (var u = 0; u < productArr.length; u++) {

                let leProduct = productArr[u];

                let g = 0;

                for (var title of productHeader[0]) {

                    tempProdObj[title] = leProduct[g];

                    g = g + 1;

                };

                productData[tempProdObj["Name"]] = tempProdObj;

                tempProdObj = {};

            };

            //#region (LEGACY) CODE FOR OLD BREAKOUT TABLE OBJECT --------------------------------------------------------------------------------

            // let v = 0;

            // let breakoutTempArr = [];

            // for (var value of breakoutHeader[0]) {
            //   for (var w = 0; w < breakoutArr.length; w++) {
            //     breakoutTempArr.push(breakoutArr[w][v])
            //   }

            //   breakoutData[value] = breakoutTempArr;

            //   breakoutTempArr = [];

            //   v = v + 1;

            // };

            //#endregion -------------------------------------------------------------------------------------------------------------------------

            //#endregion ---------------------------------------------------------------------------------------------------------------------------

            let tempWasteObj = {}; //temporary object for storing values to push into the data object (clears each time)

            for (var t = 0; t < wasteArr.length; t++) {

                let currentWaste = wasteArr[t];

                let p = 0;

                for (var heading of wasteHeader[0]) {

                    tempWasteObj[heading] = currentWaste[p];

                    p = p + 1;

                };

                wasteData[tempWasteObj["Type"]] = tempWasteObj;

                tempWasteObj = {};

            };

            //#endregion -----------------------------------------------------------------------------------------------------------------------------

            let isSilkEmpty = false;
            let isTextEmpty = false;
            let isDigEmpty = false;

            //#region POPULATE FORM COLUMN IN E2R TABLES ---------------------------------------------------------------------------------------------

            if (silkE2RArr.length > 0 && silkE2RArr[0][0] !== "" || silkE2RArr.length > 1) {
                let silkE2RUpdate = easyToReads(silkE2RArr, silkE2RTableRows, silkE2RSheet, sheetHourArr);
                silkE2RBodyRange.values = silkE2RUpdate; //push the updated values commited to silkE2RArr into the SilkE2R table in Excel
            } else {
                isSilkEmpty = true;
            };

            if (textE2RArr.length > 0 && textE2RArr[0][0] !== "" || textE2RArr.length > 1) {
                let textE2RUpdate = easyToReads(textE2RArr, textE2RTableRows, textE2RSheet, sheetHourArr);
                textE2RBodyRange.values = textE2RUpdate; //push the updated values commited to textE2RArr into the TextE2R table in Excel
            } else {
                isTextEmpty = true;
            };

            if (digE2RArr.length > 0 && digE2RArr[0][0] !== "" || digE2RArr.length > 1) {
                let digE2RUpdate = easyToReads(digE2RArr, digE2RTableRows, digE2RSheet, sheetHourArr);
                digE2RBodyRange.values = digE2RUpdate; //push the updated values commited to digE2RArr into the DIGE2R table in Excel
            } else {
                isDigEmpty = true;
            };

            //#endregion -----------------------------------------------------------------------------------------------------------------------------


            //#endregion -------------------------------------------------------------------------------------------------------------------------------


            //#region POPULATE MASTER INFO -------------------------------------------------------------------------------------------------------------


            //#region RE-DEFINE RANGE VARIABLES FOR MASTER TABLE AUTO-FILL ---------------------------------------------------------------------------

            const silkE2RBodyRangeUpdate = silkE2RTable.getDataBodyRange().load("values");
            const textE2RBodyRangeUpdate = textE2RTable.getDataBodyRange().load("values");
            const digE2RBodyRangeUpdate = digE2RTable.getDataBodyRange().load("values");
            // const pressSchedulingInfoBodyRange = pressSchedulingInfo.getDataBodyRange().load("values");

            const silkHeaderRange = silkE2RTable.getHeaderRowRange().load("values");
            const textHeaderRange = textE2RTable.getHeaderRowRange().load("values");
            const digHeaderRange = digE2RTable.getHeaderRowRange().load("values");
            // const pressSchedulingInfoHeaderRange = pressSchedulingInfo.getHeaderRowRange().load("values");


            const silkE2RTableRowsUpdate = silkE2RTable.rows.load("items");
            const textE2RTableRowsUpdate = textE2RTable.rows.load("items");
            const digE2RTableRowsUpdate = digE2RTable.rows.load("items");
            // const pressSchedulingInfoRows = pressSchedulingInfo.rows.load("items");

            //#endregion -----------------------------------------------------------------------------------------------------------------------------


            await context.sync()


            //#region RELOAD RANGE VARIABLES FOR MASTER AUTO-FILL ------------------------------------------------------------------------------------

            silkE2RArr = silkE2RBodyRangeUpdate.values; //moves all values of the SilkE2R table to an array
            textE2RArr = textE2RBodyRangeUpdate.values; //moves all values of the TextE2R table to an array
            digE2RArr = digE2RBodyRangeUpdate.values; //moves all values of the TextE2R table to an array
            // let pressSchedulingArr = pressSchedulingInfoBodyRange.values;

            silkDataSet = [];
            textDataSet = [];
            digDataSet = [];

            priorityNum = 1;

            silkDataSet = createDataSet(silkE2RArr, "Silk");
            textDataSet = createDataSet(textE2RArr, "Text");
            digDataSet = createDataSet(digE2RArr, "Digital");

            pressSchedulingInfoTable("Populate");

            // updateDataFromTable(pressSchedArr);





            //need a condition that if there is more than 1 row in the shceduling table or there is one row with info in it, then need to compare incoming E2R type and form to existing info in press scheduling table and find matchs, then update the table info with the new E2R info







            const silkHeader = silkHeaderRange.values;
            const textHeader = textHeaderRange.values;
            const digHeader = digHeaderRange.values;


            let silkRowItems = silkE2RTableRowsUpdate.items;
            let textRowItems = textE2RTableRowsUpdate.items;
            let digRowItems = digE2RTableRowsUpdate.items;


            let silkArrCopy = JSON.parse(JSON.stringify(silkE2RBodyRangeUpdate.values));
            let textArrCopy = JSON.parse(JSON.stringify(textE2RBodyRangeUpdate.values));
            let digArrCopy = JSON.parse(JSON.stringify(digE2RBodyRangeUpdate.values));
            //could be done with map or filter, possibly better?

            //#endregion -----------------------------------------------------------------------------------------------------------------------------


            //#region FILL MASTER FORMS --------------------------------------------------------------------------------------------------------------

            let z = 0;

            let masterRowInfo = new Object();

            let missingForms = [];

            for (var masterRow of masterArr) { //for each row in the master sheet...

                //#region FUNCTION VARIABLES ---------------------------------------------------------------------------------------------------------

                let wasItSilk = false;
                let wasItText = false;
                let wasItDig = false;

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region CREATE MASTER SHEET ROW INFO -----------------------------------------------------------------------------------------------

                // let rowValues = masterRowItems[z].values; //an array of all the row values at position z of the master table

                //the following matches the master table headers with the data and column index in row [z], assigning each as a property to each 
                //header within the masterRowInfo object.
                for (var name of masterHeader[0]) {
                    createRowInfo(masterHeader, name, masterRow, masterArrCopy, masterRowInfo, z, masterSheet);
                };

                let masterUJID = masterRowInfo["UJID"].value;

                let masterCode = masterRowInfo["Code"].value; //returns the client code of the the current row in the master table

                let masterProduct = masterRowInfo["Product"].value; //returns the product of the current row in the master table

                let masterVersion = masterRowInfo["Version No"].value; //returns the version ID of the current row in the master table

                let masterOptions = masterRowInfo["Options"].value;

                let masterWeeks = masterRowInfo["Wks"].value;

                let masterQuantity = masterRowInfo["Qty"].value;


                if (masterUJID == "" && masterProduct == "") {
                    console.log(`Looks like row ${z + 2} is missing the UJID, Product, and possibly even more important info. Please update this info and run the "Populate Forms" function again to fix this line`);
                    z = z + 1;
                    break;
                } else if (masterUJID == "" && masterProduct !== "") {
                    console.log(`Looks like row ${z + 2} is missing a UJID. Please update this info and run the "Populate Forms" function again to fix this line`);
                    z = z + 1;
                    break;
                } else if (masterUJID !== "" && masterProduct == "") {
                    console.log(`Looks like row ${z + 2} is missing the Product. Please update this info and run the "Populate Forms" function again to fix this line`);
                    z = z + 1;
                    break;
                };

                //swaps out the product for it's abbreviation (and all other possibly abbrevations if the product may have more than one)
                // let masterProductAbbr = [];

                // masterProductAbbr.push(productData[masterProduct]["Abbr"]);

                // if (masterProductAbbr[0].includes(",")) {
                //   masterProductAbbr = masterProductAbbr[0].split(",");
                // }

                let masterBreakout;

                let potentalFoldOnly = ["MENU", "NonProfit80#", "x.80#custom", "x.Flyer.10.5x17", "x.Flyer.8.5x10.5", "x.Menu.10.5x17"];

                let foldOnlyOverwrite = false;

                try {

                    for (let menuType of potentalFoldOnly) {
                        if (masterProduct == menuType && (masterWeeks == 0 || masterQuantity == 0)) {
                            masterBreakout = "Fold Only";
                            foldOnlyOverwrite = true;
                        };
                    };

                    if (!foldOnlyOverwrite) {
                        masterBreakout = productData[masterProduct]["Breakout"];
                    };
                } catch (e) {
                    console.log(e);
                };

                //#endregion -------------------------------------------------------------------------------------------------------------------------


                // apparelArr.forEach((item) => {
                //   if (masterProduct == item) {
                //     formsToCarry = "APPAREL";
                //   };
                // });

                //comment

                let isApperal = false;
                let ignoreMissing = false;

                for (let apparelItem of apparelArr) {
                    if (masterProduct == apparelItem) {
                        isApperal = true;
                    };
                };

                for (let ignoreItem of defaultToIgnoreArr) {
                    if (masterProduct == ignoreItem) {
                        ignoreMissing = true;
                    };
                };

                if (masterOptions.includes("ZSHELF")) {

                    formToCarry = "ZSHELF";

                } else if (masterVersion.includes("UA-")) {

                    formToCarry = "UA";

                } else if (isApperal == true) {

                    formToCarry = "APPAREL";

                } else if (ignoreMissing == true) {

                    formToCarry = "IGNORE";
                } else {

                    //#region TRY TO FIND FORM NUMBER IN SILKE2R ---------------------------------------------------------------------------------------

                    let a = 0;

                    if (isSilkEmpty == false) {

                        for (var silkRow of silkE2RArr) {

                            wasItSilk = carryForm(silkRow, silkArrCopy, silkRowItems, silkHeader, masterUJID, a, silkE2RSheet);

                            if (wasItSilk) {
                                break;
                            };

                            a = a + 1; //repeat until we have gone through all the rows in the SilkE2R table

                        };

                    };

                    //#endregion -----------------------------------------------------------------------------------------------------------------------

                    //#region TRY TO FIND FORM NUMBER IN TEXTE2R ---------------------------------------------------------------------------------------

                    if (!wasItSilk) { //if the product was not in the SilkE2R, then we move on to check if it is in the TextE2R

                        if (isTextEmpty == false) {

                            let b = 0;

                            for (var textRow of textE2RArr) {

                                wasItText = carryForm(textRow, textArrCopy, textRowItems, textHeader, masterUJID, b, textE2RSheet);

                                if (wasItText) {
                                    break;
                                };

                                b = b + 1; //repeat until we have gone through all the rows in the SilkE2R table

                            };

                        };


                        //  let textCode;
                        //  let textProduct;

                        //  // let silkRow = silkE2RArr[0][z];
                        //  let b = 0;

                        //  for (var textRow of textE2RArr) { //for each row in textE2R...

                        //    if (!textRow[0].startsWith("Layout")) { //skips rows that start new form numbers and are not actual products

                        //      let textRowValues = textRowItems[b].values; //an array of all the row values at position [b] of the textE2R table

                        //      let spaceSplitText = textRow[0].split(" "); //splits the metrix info apart by spaces and makes it into an array of items

                        //      //takes the string out from the array between the first and second space, then removes the ( character from said text. 
                        //      //Join combines this new blank record with the code to take it from an array to a value
                        //      textCode = Number((spaceSplitText.slice(1, 2))[0].split("(").join("")); 

                        //      //returns (as value, not array becuase of [0] at end) the 2nd (0 indexed) item in the array, and stops before the 3rd
                        //      textProduct = spaceSplitText.slice(2, 3)[0];

                        //      let textRowInfo = new Object();

                        //      //returns true if the code and product for this row in the master table are in the textE2R
                        //      wasItText = compareMasterToE2R(masterProductAbbr, textProduct, textCode, masterCode);

                        //      // console.log("");

                        //      if (wasItText) { //if it was in the textE2R table, then we create an object for the text row and carry over the form number

                        //        //the following matches the TextE2R table headers with the data and column index in row [b], assigning each as a property 
                        //to each header within the textRowInfo object.
                        //        for (var textName of textHeader[0]) {
                        //          createRowInfo(textHeader, textName, textRowValues, textArrCopy, textRowInfo, b);
                        //        };

                        //        formToCarry = textRowInfo["Form"].value; //gets the form number value from the row

                        //        // console.log(formToCarry);

                        //        break;

                        //      };         

                        //    };

                        //      b = b + 1;

                        //  };

                    };

                    //#endregion -----------------------------------------------------------------------------------------------------------------------

                    //#region TRY TO FIND FORM NUMBER IN DIGE2R ----------------------------------------------------------------------------------------

                    if (!wasItSilk && !wasItText) { //if the product was not found it either the text or silkE2R, then we move onto the DIGE2R

                        if (isDigEmpty == false) {

                            let c = 0;

                            for (var digRow of digE2RArr) {

                                wasItDig = carryForm(digRow, digArrCopy, digRowItems, digHeader, masterUJID, c, digE2RSheet);

                                if (wasItDig) {
                                    break;
                                };

                                c = c + 1;

                            };

                        };

                    };


                    if (isSilkEmpty && isTextEmpty && isDigEmpty) {

                        console.log("All E2Rs are empty, so this must be a Planned Separately Form");
                        plannedSep = true;

                    } else if (!wasItSilk && !wasItText && !wasItDig) {

                        console.log("Look's like I'm a missing product...");

                        missingForms.push({
                            row: z + 2,
                            code: masterRowInfo["Code"].value,
                            company: masterRowInfo["Company"].value,
                            total: masterRowInfo["Total"].value,
                            orderStatus: masterRowInfo["Order Status"].value
                        });

                    };

                    //#endregion -----------------------------------------------------------------------------------------------------------------------

                };

                //#region APPLY CORRECT FORM NUMBER TO MASTER ----------------------------------------------------------------------------------------

                let masterFormColumnIndex = masterRowInfo["Forms"].columnIndex;

                if (!plannedSep) {

                    if (formToCarry) {
                        masterArr[z][masterFormColumnIndex] = formToCarry;
                    } else {
                        masterArr[z][masterFormColumnIndex] = "MISSING";
                    };

                };

                //#endregion -------------------------------------------------------------------------------------------------------------------------


                //#region APPLY CORRECT BREAKOUT TYPE TO MASTER --------------------------------------------------------------------------------------

                let masterTypeColumnIndex = masterRowInfo["Type"].columnIndex;

                masterArr[z][masterTypeColumnIndex] = masterBreakout;

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                //#region APPLY DATA VALIDATION TO NON-NUMBERED FORMS --------------------------------------------------------------------------------

                let leRange = "A" + (z + 2);

                let zeRange = masterSheet.getRange(leRange);

                zeRange.dataValidation.clear();

                zeRange.format.fill.clear();
                zeRange.format.font.bold = false;

                let typeRange = masterSheet.getRangeByIndexes(z + 2, masterTypeColumnIndex, 1, 1);

                typeRange.dataValidation.clear();

                if (!Number(masterArr[z][masterFormColumnIndex])) {
                    // console.log(masterArr[z][masterFormColumnIndex] + "is not a number!");
                    let dv = {
                        list: {
                            inCellDropdown: true,
                            source: customFormsBodyRange
                        }
                    };

                    zeRange.dataValidation.rule = dv;

                    conditionalFormatting(masterSheet, zeRange, masterArr[z][masterFormColumnIndex], null);


                    await context.sync();

                };

                // console.log(masterArr[z][masterFormColumnIndex] + "is not a number!");
                let dvType = {
                    list: {
                        inCellDropdown: true,
                        source: linesBodyRange
                    }
                };

                typeRange.dataValidation.rule = dvType;

                //#endregion -------------------------------------------------------------------------------------------------------------------------

                // console.log(z);

                z = z + 1;

            };

            masterBodyRange.values = masterArr; //write masterArr to the master table in Excel

            console.log(`There are ${missingForms.length} missing forms, listed here:`, missingForms);

            //#endregion -----------------------------------------------------------------------------------------------------------------------------


            //#endregion -------------------------------------------------------------------------------------------------------------------------------


            activateEvents();

            // eventResult = masterTable.onChanged.add(handleChange);

            // $("#populate-forms").css("display", "none");
            // $("#clear-forms").css("display", "flex");
            // $("#breakout").css("display", "flex");

            location.reload();

        });


    } catch (err) {
        console.error(err);
        // showMessage(error, "show");
    };

    $("#loading-background").css("display", "none");

    $("#loading-background").css("display", "none");

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------



async function handleChange(event) {
    await Excel.run(async (context) => {



        if (event.details == undefined) {
            console.log("Event is undefined");
            return;
        } else {
            console.log("The Event is: ");
            console.log(event);
        };

        const validation = context.workbook.worksheets.getItem("Validation");
        const customFormsTable = validation.tables.getItem("CustomForms").load("name");
        const linesTable = validation.tables.getItem("Lines").load("name");
        const customFormsBodyRange = customFormsTable.getDataBodyRange().load("values");
        const linesBodyRange = linesTable.getDataBodyRange().load("values");

        let address = event.address;
        let changedWorksheet = context.workbook.worksheets.getItem(event.worksheetId).load("name");
        let changedRange = changedWorksheet.getRange(address);

        changedRange.dataValidation.clear();


        //turns out that the column index for both the changedWorksheet & the changedTable are identical, so I am just sticking with the worksheet one
        let changedColumn = changedWorksheet.getRange(address).load("columnIndex");
        let changedRow = changedWorksheet.getRange(address).load("rowIndex");

        let changedTable = context.workbook.tables.getItem(event.tableId).load("name");
        let changedTableBodyRange = changedTable.getDataBodyRange().load("values");
        let changedTableHeader = changedTable.getHeaderRowRange().load("values");
        let changedTableColumns = changedTable.columns
        changedTableColumns.load("items/name");
        let changedTableRows = changedTable.rows;
        changedTableRows.load("items");


        await context.sync();

        let changedTableArr = changedTableBodyRange.values;
        let changedHeadersValues = changedTableHeader.values;
        let tableRowItems = changedTableRows.items;
        let changedTableRowIndex = changedRow.rowIndex - 1;

        let changedRowValues = tableRowItems[changedTableRowIndex].values

        let changedTableArrCopy = JSON.parse(JSON.stringify(changedTableBodyRange.values));

        let changedRowInfo = new Object();

        for (var headName of changedHeadersValues[0]) {
            createRowInfo(changedHeadersValues, headName, changedRowValues[0], changedTableArrCopy, changedRowInfo, changedRow.rowIndex, changedWorksheet);
        }

        formsColumnIndex = changedRowInfo.Forms.columnIndex;
        // console.log(changedRow.rowIndex + 1); //add one since it is zero-indexed
        // console.log(changedRowInfo.Forms.value);


        //if the changed value is in the Forms column, trigger a re-evaluation of the data validation for the changed cell
        if (changedRowInfo.Forms.columnIndex == changedColumn.columnIndex) {

            changedRange.dataValidation.clear();

            if (!Number(event.details.valueAfter) && event.details.valueAfter !== "") {
                // console.log(masterArr[z][masterFormColumnIndex] + "is not a number!");
                let dv = {
                    list: {
                        inCellDropdown: true,
                        source: customFormsBodyRange
                    }
                };

                changedRange.dataValidation.rule = dv;

                conditionalFormatting(changedWorksheet, changedRange, changedRowInfo.Forms.value, null);


                await context.sync();

            };

            if (Number(event.details.valueAfter) || event.details.valueAfter == "") {

                changedRange.dataValidation.clear();

                conditionalFormatting(changedWorksheet, changedRange, changedRowInfo.Forms.value, null);


                console.log("I am a number now!");

            };


        };

        if (changedRowInfo.Type.columnIndex == changedColumn.columnIndex) {

            changedRange.dataValidation.clear();

            let dv = {
                list: {
                    inCellDropdown: true,
                    source: linesBodyRange
                }
            };

            changedRange.dataValidation.rule = dv;

            await context.sync();

        };

        // console.log("Address of event: " + event.address);
        // console.log("Changed column is: " + changedColumn.columnIndex);
        // console.log("Changed row is: " + changedRow.rowIndex);





        // remove();   

    }).catch(errorHandlerFunction);
}

// async function remove() {
//   await Excel.run(eventResult.context, async (context) => {
//     eventResult.remove();
//     await context.sync();

//     eventResult = null;
//     console.log("Event handler successfully removed.");
//   });
// }





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






async function breakout() {


    deactivateEvents();

    $("#loading-background").css("display", "flex");

    if ($("#breakout").text() == "Delete Breakout Sheets") {
        $("#loading-head").text("Removing Breakout Tables");
    } else {
        $("#loading-head").text("Creating Breakout Tables");
    };



    await Excel.run(async (context) => {

        const allSheets = context.workbook.worksheets.load("items/name");
        const validation = context.workbook.worksheets.getItem("Validation");
        const linesTable = validation.tables.getItem("Lines");
        // const defaultToIgnoreTable = validation.tables.getItem("DefaultToIgnore");
        const linesBodyRange = linesTable.getDataBodyRange().load("values");
        // const defaultToIgnoreBodyRange = defaultToIgnoreTable.getDataBodyRange().load("values");
        const linesHeaderRange = linesTable.getHeaderRowRange().load("values");
        const masterSheet = context.workbook.worksheets.getItem("Master").load("name");
        const masterTable = masterSheet.tables.getItem("Master");
        const masterBodyRange = masterTable.getDataBodyRange().load("values");
        const masterHeaderRange = masterTable.getHeaderRowRange().load("values");
        const masterTableRows = masterTable.rows.load("items");

        await context.sync();

        let linesArr = linesBodyRange.values;
        let linesHeader = linesHeaderRange.values;
        // let defaultToIgnoreArr = defaultToIgnoreBodyRange.values;
        let masterArr = masterBodyRange.values;
        let masterHeader = masterHeaderRange.values;
        let masterRowItems = masterTableRows.items;
        let masterArrCopy = JSON.parse(JSON.stringify(masterBodyRange.values));

        // await context.sync();

        console.log("Doing it...")
        const doIt = await removeBreakoutSheets(linesArr, allSheets);
        console.log(doIt);

        await context.sync();

        if ($("#breakout").text() == "Delete Breakout Sheets") {
            $("#breakout").text("Breakout");

            console.log("exiting...");

            return;
        };


        //////////////////////////////////////////////// STOPPED HERE TRYING TO FIGURE OUT ROW COLOR STORING ///////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        // for (let row of masterRowItems) {
        //   let rowRange = row.getRange();

        //   let rangeFill = rowRange.format.fill;
        //   rangeFill.load(["color"])

        //   await context.sync();
        //   console.log("range.format.fill.color", rangeFill.color)
        // };



        // var allTablesPrior = context.workbook.tables;
        // allTablesPrior.load("items/name");

        await context.sync();

        // allTablesPrior.items.forEach((tablePrior) => {
        //   console.log(tablePrior.name);
        // });


        // allTablesPrior.items

        // let tempLinesObj = {}; //temporary object for storing values to push into the data object (clears each time)

        linesData = [];

        for (var p = 0; p < linesArr.length; p++) {

            let currentLines = linesArr[p][0];

            linesData.push(currentLines);

        };

        // let lineArray = [];


        let filteredData = {};

        linesData.forEach((line) => {

            filteredData[line] = [];

        });

        // console.log(filteredData);



        let z = 1;
        let masterRowInfo = new Object();

        let empty = [];
        let missing = [];
        let ignore = [];
        let shipping = [];
        let printed = [];
        let digitalBreakout = [];

        let emptyFormatting = [];
        let missingFormatting = [];
        let ignoreFormatting = [];
        let shippingFormatting = [];
        let printedFormatting = [];
        let digitalFormatting = [];

        let allNormalTables = [];

        let normalFormatting = [];

        let masterUJIDColumnIndex;

        let missingData = {};

        hiddenLinesData.forEach((item) => {

            missingData[item] = [];

        });

        let overwriteMissing = false;


        let formsObj = {};

        let tableName;

        for (let zeLine of linesData) {
            normalBreakoutsFormatting[zeLine] = [];
        };


        // linesData.forEach(async (line) => {
        // for (let line of linesData) {
        //   //#region CREATE TABLE NAMES FROM LINE DATA --------------------------------------------------------------------------------------------------

        //     let lineItemSplit = line.split(" ");

        //     let firstWord = lineItemSplit[0].charAt(0).toUpperCase() + lineItemSplit[0].substr(1).toLowerCase();

        //     if (firstWord.includes("-")) {
        //       firstWord = firstWord.replace("-", "");
        //     };

        //     lineItemSplit.shift();

        //     tableName = firstWord;

        //     for (var word of lineItemSplit) {

        //       if (word.includes("-")) {
        //         word = word.replace("-", "");
        //       };

        //       let cheese = word.charAt(0).toUpperCase() + word.substr(1).toLowerCase(); //+ lineItemSplit[word].slice(1);

        //       tableName = tableName + cheese;

        //     };

        //     normalBreakoutsFormatting[tableName] = [];

        //     listOfBreakoutTables.push(tableName);

        //   };


        // filteredData["MISSING"] = [];

        for (let masterRow of masterArr) {

            formsObj = {};

            //the following matches the master table headers with the data and column index in row [z], assigning each as a property to each 
            //header within the masterRowInfo object.
            for (let name of masterHeader[0]) {
                createRowInfo(masterHeader, name, masterRow, masterArrCopy, masterRowInfo, z, masterSheet);
            };

            z++;

            await context.sync();
            // const cellProperties = propertiesToGet.value[0][0];

            // console.log(
            //   `Address: ${cellProperties.address}\nStyle: ${cellProperties.style}\nFill Color: ${cellProperties.format.fill.color}\nFont Color: ${cellProperties.format.font.color}`)

            let masterForms = masterRowInfo["Forms"].value;
            let masterType = masterRowInfo["Type"].value;
            masterUJIDColumnIndex = masterRowInfo["UJID"].columnIndex;

            let masterExtras = masterRowInfo["Extras"].value;
            let masterCutoff = masterRowInfo["Cutoff"].value;

            // if (masterForms == "MISSING") {
            //   console.log("Breakpoint here");
            // };

            for (let head of masterHeader[0]) {
                let formsAddress = masterRowInfo[head].cellProps.value[0][0].address
                let formsFill = masterRowInfo[head].cellProps.value[0][0].format.fill.color
                let formsFontColor = masterRowInfo[head].cellProps.value[0][0].format.font.color
                let formsFontBold = masterRowInfo[head].cellProps.value[0][0].format.font.bold
                let formsFontItalic = masterRowInfo[head].cellProps.value[0][0].format.font.italic

                formsObj[head] = {
                    formsAddress,
                    formsFill,
                    formsFontColor,
                    formsFontBold,
                    formsFontItalic
                };
            };


            // normalBreakoutsFormatting[masterType] = [];





            //need to load in values from Lines table in validation and loop through them for each row of master and if matches move row info over into that type's array

            // if (masterExtras > 0 && masterExtras !== "" || masterCutoff > 0 && masterCutoff !== "" ) {

            //   shipping.push(masterRow);
            //   missingData["Shipping"].push(masterRow);

            // }

            try {

                //if form number is between 101 & 599, then it is a digital form and should be duplicated into it's own digital breakout (without taking it away from any other breakout it might fall into, which is why it is outside of the if/else area below)
                if (masterForms > 100 && masterForms < 600) {

                    digitalBreakout.push(masterRow);
                    digitalFormatting.push(formsObj);

                    missingData["DIGITAL"].push(masterRow);

                };

                if (masterType == "Shipping") { //if type is shipping, push to just shipping array

                    shipping.push(masterRow);
                    shippingFormatting.push(formsObj);

                    missingData["Shipping"].push(masterRow);

                } else if (masterType == "Ignore") {

                    ignore.push(masterRow);
                    ignoreFormatting.push(formsObj);

                    missingData["IGNORE"].push(masterRow);

                    //#region PUSH EXTRAS & CUTOFFS TO SHIPPING BREAKOUT -------------------------------------------------------------------------------------

                    //if includes extras or cutoffs, push to both type array and shipping array
                    // } else if (masterExtras > 0 && masterExtras !== "" || masterCutoff > 0 && masterCutoff !== "") { 

                    //   if (masterForms == "MISSING") {

                    //     missing.push(masterRow);
                    //     missingFormatting.push(formsObj);

                    //     missingData["MISSING"].push(masterRow);

                    //     shipping.push(masterRow);
                    //     shippingFormatting.push(formsObj);

                    //     missingData["Shipping"].push(masterRow);
                    //     filteredData[masterType].push(masterRow);
                    //     normalBreakoutsFormatting[masterType].push(formsObj);


                    //   } else if (masterForms == "IGNORE") {

                    //     // console.log("I need to be ignored");
                    //     ignore.push(masterRow);
                    //     ignoreFormatting.push(formsObj);

                    //     missingData["IGNORE"].push(masterRow);

                    //   } else if (masterForms == "PRINTED") {

                    //     printed.push(masterRow);
                    //     printedFormatting.push(formsObj);

                    //     shipping.push(masterRow);
                    //     shippingFormatting.push(formsObj);

                    //     missingData["PRINTED"].push(masterRow);
                    //     missingData["Shipping"].push(masterRow);

                    //   } else {

                    //     shipping.push(masterRow);
                    //     shippingFormatting.push(formsObj);

                    //     missingData["Shipping"].push(masterRow);
                    //     filteredData[masterType].push(masterRow);
                    //     normalBreakoutsFormatting[masterType].push(formsObj);

                    //   };

                    //#endregion -------------------------------------------------------------------------------------------------------------------------------

                } else if (masterForms == "MISSING") { //also pushes these items to the normal type breakouts as well

                    missing.push(masterRow);
                    missingFormatting.push(formsObj);

                    missingData["MISSING"].push(masterRow);
                    filteredData[masterType].push(masterRow);
                    normalBreakoutsFormatting[masterType].push(formsObj);


                } else if (masterForms == "IGNORE") {

                    // console.log("I need to be ignored");
                    ignore.push(masterRow);
                    ignoreFormatting.push(formsObj);

                    missingData["IGNORE"].push(masterRow);


                } else if (masterForms == "PRINTED") {

                    printed.push(masterRow);
                    printedFormatting.push(formsObj);

                    missingData["PRINTED"].push(masterRow);


                } else { //if neither shipping type nor has extras or cutoffs, treat normally

                    filteredData[masterType].push(masterRow);
                    normalBreakoutsFormatting[masterType].push(formsObj);

                    // normalFormatting.push(formsObj);

                };

            } catch (e) {
                console.log(e);
                console.log("Could not find master type: ", masterType);
                empty.push(masterRow);
                missingData["Empty"].push(masterRow);
            };

            //need to alert user of any items that may appear in the empty array. This means that the type column is either empty or has an invlaid record

        };

        // let columnsToHide = ["Type", "AS", "Email", "Version No", "Artwork Variable", "Options", "UPS", "UJID"];
        let columnsToHide = [];

        //hellowse



        let missingTable = addSheetAndTable("MISSING", allSheets, missingData, masterHeader, "Missing");

        let printedTable = addSheetAndTable("PRINTED", allSheets, missingData, masterHeader, "Printed");

        let ignoreTable = addSheetAndTable("IGNORE", allSheets, missingData, masterHeader, "Ignore");

        let digitalTable = addSheetAndTable("DIGITAL", allSheets, missingData, masterHeader, "Digital");


        // let shippingTable = addSheetAndTable("Shipping", allSheets, missingData, masterHeader, "Shipping");


        await context.sync();

        let missingRowCount = missingTable.table.rows.getCount();
        let printedRowCount = printedTable.table.rows.getCount();
        let ignoreRowCount = ignoreTable.table.rows.getCount();
        let digitalRowCount = digitalTable.table.rows.getCount();

        // let shippingRowCount = shippingTable.table.rows.getCount();

        await context.sync();

        let theMissingFormat = styleCells(missingTable.table, missingFormatting, missingRowCount.value, masterHeader, "MISSING");
        let thePrintedFormat = styleCells(printedTable.table, printedFormatting, printedRowCount.value, masterHeader, "PRINTED");
        let theIgnoreFormat = styleCells(ignoreTable.table, ignoreFormatting, ignoreRowCount.value, masterHeader, "IGNORE");
        let theDigitalFormat = styleCells(digitalTable.table, digitalFormatting, digitalRowCount.value, masterHeader, "DIGITAL");
        // let theShippingFormat = styleCells(shippingTable.table, shippingFormatting, shippingRowCount.value, masterHeader, "Shipping");



        hideColumns(missingTable.table, columnsToHide);
        hideColumns(printedTable.table, columnsToHide);
        hideColumns(ignoreTable.table, columnsToHide);
        hideColumns(digitalTable.table, columnsToHide);
        // hideColumns(shippingTable.table, columnsToHide);

        printSettings(missingTable.sheet);
        printSettings(printedTable.sheet);
        printSettings(ignoreTable.sheet);
        printSettings(digitalTable.sheet);
        // printSettings(shippingTable.sheet);


        await context.sync();

        // linesData.forEach(async (line) => {
        for (let line of linesData) {

            //#region CREATE TABLE NAMES FROM LINE DATA --------------------------------------------------------------------------------------------------

            let lineItemSplit = line.split(" ");

            let firstWord = lineItemSplit[0].charAt(0).toUpperCase() + lineItemSplit[0].substr(1).toLowerCase();

            if (firstWord.includes("-")) {
                firstWord = firstWord.replace("-", "");
            };

            lineItemSplit.shift();

            tableName = firstWord;

            for (var word of lineItemSplit) {

                if (word.includes("-")) {
                    word = word.replace("-", "");
                };

                let cheese = word.charAt(0).toUpperCase() + word.substr(1).toLowerCase(); //+ lineItemSplit[word].slice(1);

                tableName = tableName + cheese;

            };

            // normalBreakoutsFormatting[tableName] = [];

            listOfBreakoutTables.push(tableName);

            //#endregion ---------------------------------------------------------------------------------------------------------------------------------

            let tempTable = [];

            if (tableName !== "Shipping" && tableName !== "Ignore") {

                // dynamicFormatting[tableName] = {

                // }
                // printedFormatting.push(formsObj);


                let thisTable = addSheetAndTable(line, allSheets, filteredData, masterHeader, tableName);

                let normalRowCount = thisTable.table.rows.getCount();


                await context.sync();


                let theNormalFormat = styleCells(thisTable.table, normalBreakoutsFormatting[line], normalRowCount.value, masterHeader, tableName);

                hideColumns(thisTable.table, columnsToHide);

                printSettings(thisTable.sheet);

            };

            await context.sync();
        };


        console.log("FILTERED DATA: ", filteredData);
        console.log("MISSING FORMS: ", missing);
        console.log("IGNORED ITEMS: ", ignore);
        console.log("EMPTY TYPE DATA: ", empty);
        console.log("DIGITAL TYPE DATA: ", digitalBreakout);

        let emptyUJIDs = [];

        if (empty.length > 0) {
            //show a warning box and do not generate breakout sheets

            if (empty.length > 1) {

                for (let u = 0; u < empty.length; u++) {

                    emptyUJIDs.push(empty[u][masterUJIDColumnIndex]);

                }

            } else {
                emptyUJIDs = empty[0][masterUJIDColumnIndex];
            };

            document.getElementById("empty-ujid").innerHTML = emptyUJIDs;

            emptyWarning();

            return;

        };

        let shippingTable = addSheetAndTable("Shipping", allSheets, missingData, masterHeader, "Shipping");


        await context.sync();

        let shippingRowCount = shippingTable.table.rows.getCount();

        await context.sync();

        let theShippingFormat = styleCells(shippingTable.table, shippingFormatting, shippingRowCount.value, masterHeader, "Shipping");


        hideColumns(shippingTable.table, columnsToHide);

        printSettings(shippingTable.sheet);


        await context.sync();




        // let newTempObj = {};
        // let newFormattedCells = [];

        // let allFormats = [theMissingFormat, thePrintedFormat, theIgnoreFormat, theShippingFormat];

        // console.log(allNormalTables);

        // allFormats.forEach((arr) => {

        //   if (arr.length !== 0) {
        //     for (var y = 0; y < arr.length; y++) {
        //       newTempObj = {
        //         address: arr[y].cell.address,
        //         fill: arr[y].fill,
        //         fontColor: arr[y].fontColor,
        //         fontBold: arr[y].fontBold,
        //         fontItalic: arr[y].fontItalic,

        //       };
        //       newFormattedCells.push(newTempObj);
        //     };

        //     console.log(`${arr[0].sheet} sheet's formatting was set to the following:`);
        //     console.log(newFormattedCells);
        //   };

        // });

        if (emptySheets) {
            console.log("The following sheets are empty:");
            console.log(emptySheets);
        };




        console.log(`Breakout sheet's hidden columns should now be hidden`);





        // var allTablesAfter = context.workbook.tables;
        // allTablesAfter.load("items/name");

        // await context.sync();

        // allTablesAfter.items.forEach((tableAfter) => {
        //   console.log(tableAfter.name);
        // });

        // console.log(allTablesAfter.items.name);

        // listOfBreakoutTables.forEach(async(table) => {

        //   let theTable = context.workbook.tables.getItem(table);
        //   let theTableRange = theTable.columns.getItem("Type").getRange().load("columnHidden");

        //   await context.sync();

        //   theTableRange.columnHidden = true;

        //   console.log(`Type Columns should be hidden on ${table}`);

        // });




        activateEvents();


        if ($("#breakout").text() == "Breakout") {
            $("#breakout").text("Delete Breakout Sheets");
        } else {
            $("#breakout").text("Breakout");
        };


    });



    $("#loading-background").css("display", "none");

};



function hideColumns(table, hideColArr) {
    for (let column of hideColArr) {
        let thisColumn = table.columns.getItem(column).getRange();
        thisColumn.columnHidden = true;
    };

    let addressColumn = table.columns.getItem("Address").getRange();
    addressColumn.format.columnWidth = 185;
    addressColumn.format.wrapText = true;

    let companyColumn = table.columns.getItem("Company").getRange();
    companyColumn.format.columnWidth = 185;
    companyColumn.format.wrapText = true;

    let notesColumn = table.columns.getItem("Notes").getRange();
    notesColumn.format.columnWidth = 185;
    notesColumn.format.wrapText = true;
};


function styleCells(table, formattingArr, rowCount, headerValues, sheetName) {

    let tempObj = {};

    let allFormattedCells = [];

    headerValues = headerValues[0];

    if (rowCount == 0) {
        emptySheets.push(sheetName);
    };

    for (let u = 0; u < rowCount; u++) {

        try {

            let arrRow = formattingArr[u];

            let thisRow = table.rows.getItemAt(u).getRange();

            for (let v = 0; v < headerValues.length; v++) {

                try {
                    let cell = thisRow.getCell(0, v).load("address");
                    // let headerValues[0][v] = headerValues[0][v];
                    cell.format.fill.color = arrRow[headerValues[v]].formsFill;
                    cell.format.font.color = arrRow[headerValues[v]].formsFontColor;
                    cell.format.font.bold = arrRow[headerValues[v]].formsFontBold;
                    cell.format.font.italic = arrRow[headerValues[v]].formsFontItalic;


                    tempObj = {
                        sheet: sheetName,
                        index: u,
                        cell: cell,
                        fill: arrRow[headerValues[v]].formsFill,
                        fontColor: arrRow[headerValues[v]].formsFontColor,
                        fontBold: arrRow[headerValues[v]].formsFontBold,
                        fontItalic: arrRow[headerValues[v]].formsFontItalic
                    };

                    allFormattedCells.push(tempObj);

                } catch (e) {
                    // console.log(`The value of v at error: ${v}`);
                    console.log(e);
                };

            };

        } catch (e) {
            // console.log(`The value of u at error: ${u}`);
            // console.log(`The value of v at error: ${v}`);
            console.log(e);
        };


    };

    return allFormattedCells;

};



function printSettings(sheet) {
    sheet.pageLayout.rightMargin = 0;
    sheet.pageLayout.leftMargin = 0;
    sheet.pageLayout.topMargin = 0;
    sheet.pageLayout.bottomMargin = 0;
    sheet.pageLayout.headerMargin = 0;
    sheet.pageLayout.footerMargin = 0;

    sheet.pageLayout.centerHorizontally = true;
    sheet.pageLayout.centerVertically = true;

    let pageLayoutZoomOptions = {
        'horizontalFitToPages': 1,
        'verticalFitToPages': 0,
    };

    sheet.pageLayout.zoom = pageLayoutZoomOptions;

    // Set the first row as the title row for every page.
    // sheet.pageLayout.setPrintTitleRows("$1:$2");
    sheet.pageLayout.setPrintTitleRows("$2:$2");


    // Limit the area to be printed to the range "A1:D100".
    // sheet.pageLayout.setPrintArea("A1:D100");

    sheet.pageLayout.orientation = Excel.PageOrientation.landscape;
};





function addSheetAndTable(line, allSheets, filteredData, masterHeader, tableName) {

    // Excel.run(async (context) => {



    // deactivateEvents();

    let table = "";
    let tableColumnLetter = "";

    let sheet = allSheets.add(line);
    sheet.load("name, position");
    let tableRowLength = filteredData[line].length;
    let tableHeaderLength = masterHeader[0].length;

    let alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    tableColumnLetter = printToLetter(tableHeaderLength, alphabet)

    let tableTitleString = "A1:" + tableColumnLetter + "1";
    let tableTitleRange = sheet.getRange(tableTitleString);

    tableTitleRange.merge(true);

    let mergedTitleRange = sheet.getRange("A1");

    // let titleRowRange = mergedTitleRange.getEntireRow();

    mergedTitleRange.values = [[line]];
    mergedTitleRange.format.autofitColumns();
    mergedTitleRange.format.horizontalAlignment = "center";
    mergedTitleRange.format.verticalAlignment = "center";
    mergedTitleRange.format.font.size = 22;
    mergedTitleRange.format.font.bold = true;
    mergedTitleRange.format.rowHeight = 62;
    mergedTitleRange.format.fill.color = "#F3EAF7";

    // titleRowRange.format.rowHeight = 20;







    // let tableRangeString = "A1:" + tableColumnLetter + tableRowLength;
    let tableRangeString = "A2:" + tableColumnLetter + "2";

    // let tableRange = sheet.getRange(tableRangeString);

    table = sheet.tables.add(tableRangeString, true /*hasHeaders*/);
    // let table = sheet.tables.add("A1:D1", true /*hasHeaders*/);

    // table.name = line;

    // table.name = "Sample" + q;
    table.name = tableName;

    // await context.sync();


    // let theTable = context.workbook.tables.getItem(tableName);
    // let theTableRange = theTable.columns.getItem("Type").getRange().load("columnHidden");

    // await context.sync();

    // theTableRange.columnHidden = true;

    // console.log("Type Columns should be hidden on Breakout Sheets");


    // q = q + 1;

    table.getHeaderRowRange().values = [masterHeader[0]];
    // table.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    let sheetAndTable = {
        sheet: "",
        table: ""
    };

    let tryThis = filteredData[line];

    if (tryThis == "") {
        sheet.activate();
        sheetAndTable.sheet = sheet;
        sheetAndTable.table = table;
        return sheetAndTable;
    };

    table.rows.add(null /*add rows to the end of the table*/, tryThis);

    // table.rows.add(null /*add rows to the end of the table*/, [
    //   [24, "Postcard Line 4 Podium", "TomH", "StephanieL", 33961, "ColossalPC", "Jet's Pizza - MI-142", "6275 28th St SE", "brrrrrnt74@gmail.com", 10, 5000, 0, 0, 0, 5000, "", "PRINTED", "", "MA", "", "", "", "33961-84636-1"]
    // ]);

    // console.log("bout to activate sheet");





    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    };

    sheet.activate();

    sheetAndTable.sheet = sheet;
    sheetAndTable.table = table;

    return sheetAndTable;


    // })


};


function printToLetter(number, alphabet) {

    let charIndex = number % alphabet.length;
    let quotient = number / alphabet.length;

    if (charIndex - 1 == -1) {

        charIndex = alphabet.length;
        quotient--;

    };

    result = alphabet.charAt(charIndex - 1); // + result;

    if (quotient >= 1) {

        printToLetter(parseInt(quotient));

    } else {

        // console.log(result);
        return result;

    };

};


function emptyWarning() {
    deactivateEvents();

    console.log("There are empty types preset. Exiting Breakout function...");

    $("#empty-background").css("display", "flex");

    $("#set-to-ignore").on("click", () => {
        console.log("setting to ignore...");
        setBlanksToAllOther();
        $("#empty-background").css("display", "none");
    });

    $("#handle-manually").on("click", () => {
        console.log("letting user handle...");
        $("#empty-background").css("display", "none");
    });
}


async function setBlanksToAllOther() {

    await Excel.run(async (context) => {

        const validation = context.workbook.worksheets.getItem("Validation");
        const linesTable = validation.tables.getItem("Lines");
        const linesBodyRange = linesTable.getDataBodyRange().load("values");
        const linesHeaderRange = linesTable.getHeaderRowRange().load("values");
        const masterSheet = context.workbook.worksheets.getItem("Master").load("name");
        const masterTable = masterSheet.tables.getItem("Master");
        const masterBodyRange = masterTable.getDataBodyRange().load("values");
        const masterHeaderRange = masterTable.getHeaderRowRange().load("values");
        const masterTableRows = masterTable.rows.load("items");

        await context.sync();

        let linesArr = linesBodyRange.values;
        let linesHeader = linesHeaderRange.values;
        let masterArr = masterBodyRange.values;
        let masterHeader = masterHeaderRange.values;
        let masterRowItems = masterTableRows.items;
        let masterArrCopy = JSON.parse(JSON.stringify(masterBodyRange.values));

        let z = 0;
        let masterRowInfo = new Object();

        for (let masterRow of masterArr) {

            //the following matches the master table headers with the data and column index in row [z], assigning each as a property to each 
            //header within the masterRowInfo object.
            for (let name of masterHeader[0]) {
                createRowInfo(masterHeader, name, masterRow, masterArrCopy, masterRowInfo, z, masterSheet);
            };

            let masterTypeColumnIndex = masterRowInfo["Type"].columnIndex;

            if (masterArr[z][masterTypeColumnIndex] == "") {
                masterArr[z][masterTypeColumnIndex] = "All Other";
            };

            z = z + 1;

        };

        await context.sync()

        masterBodyRange.values = masterArr;

        await context.sync();

        breakout();

    });

};


async function clearForms() {

    deactivateEvents();

    await Excel.run(async (context) => {

        const allSheets = context.workbook.worksheets.load("items/name");
        const validation = context.workbook.worksheets.getItem("Validation");
        const linesTable = validation.tables.getItem("Lines");
        const linesBodyRange = linesTable.getDataBodyRange().load("values");

        await context.sync();

        let linesArr = linesBodyRange.values;

        $("#reset-background").css("display", "flex");

        $("#reset-forms").on("click", () => {
            resetForms();
            clearPSInfo();
            removeBreakoutSheets(linesArr, allSheets);
        });

        $("#no").on("click", () => {
            $("#reset-background").css("display", "none");
        });

    });

};



//#region ON CLEAR FORMS BUTTON CLICK --------------------------------------------------------------------------------------------------------------

async function resetForms() {

    await Excel.run(async (context) => {


        deactivateEvents();

        // const validation = context.workbook.worksheets.getItem("Validation");
        // const linesTable = validation.tables.getItem("Lines");
        // const linesBodyRange = linesTable.getDataBodyRange().load("values");
        // const linesHeaderRange = linesTable.getHeaderRowRange().load("values");

        // await context.sync();

        // let linesArr = linesBodyRange.values;
        // let linesHeader = linesHeaderRange.values;

        // deleteSheets(linesArr);



        // for (var p = 0; p < linesArr.length; p++) {

        //   let currentLines = linesArr[p][0];

        //   linesData.push(currentLines);

        // };

        // linesData.forEach(async (line) => {
        //   let deleteSheet = context.workbook.worksheets.getItem(line);
        //   deleteSheet.delete();
        //   await context.sync();
        // });





        reset();
        console.log("All forms were reset");
        $("#reset-background").css("display", "none");
        $("#populate-forms").css("display", "flex");
        $("#clear-forms").css("display", "none");
        $("#breakout").css("display", "none");

    });

};

async function removeBreakoutSheets(arrOfSheets, allSheets) {

    // $("#loading-head").text("Removing Breakout Tables");

    return new Promise((resolve, reject) => {

        try {
            // deactivateEvents();

            linesData = [];

            for (var p = 0; p < arrOfSheets.length; p++) {

                let currentLines = arrOfSheets[p][0];

                linesData.push(currentLines);

            };

            let arrOfSheetNames = [];

            let removeThese = ["Validation", "SilkE2R", "TextE2R", "DIGE2R", "Master", "Press Scheduling", "MISSING", "PRINTED", "IGNORE", "DIGITAL"];

            for (let sheet of allSheets.items) {
                arrOfSheetNames.push(sheet.name);
            };

            for (let name of removeThese) {
                let indexOfName = arrOfSheetNames.indexOf(name);
                if (indexOfName !== undefined) {
                    arrOfSheetNames.splice(indexOfName, 1);
                };
            };

            deleteSheets("MISSING");
            deleteSheets("PRINTED");
            deleteSheets("IGNORE");
            deleteSheets("DIGITAL");


            for (let line of linesData) {
                for (let item of arrOfSheetNames) {
                    if (item == line) {
                        deleteSheets(line);
                        arrOfSheetNames.splice(arrOfSheetNames.indexOf(item), 1);
                        break;
                    };
                };
            };

            // linesData.forEach(async (line) => {
            //   allSheets.items.forEach((sheet) => {
            //     if (sheet.name == line) {
            //       deleteSheets(line);
            //     } else if (sheet.name == "MISSING") {
            //       deleteSheets("MISSING");
            //     } else if (sheet.name == "PRINTED") {
            //       deleteSheets("PRINTED");
            //     } else if (sheet.name == "IGNORE") {
            //       deleteSheets("IGNORE");
            //     };
            //   });
            // });
            resolve("Done.");
        } catch (e) {
            reject(e);
        };

        activateEvents();

    });

};


async function deleteSheets(worksheetName) {

    await Excel.run(async (context) => {

        let deleteSheet = context.workbook.worksheets.getItem(worksheetName);
        deleteSheet.delete();

    });

};



async function removeBreakoutSheetsTaskpane() {

    await Excel.run(async (context) => {

        const allSheets = context.workbook.worksheets.load("items/name");
        const validation = context.workbook.worksheets.getItem("Validation");
        const linesTable = validation.tables.getItem("Lines");
        const linesBodyRange = linesTable.getDataBodyRange().load("values");

        await context.sync();

        let linesArr = linesBodyRange.values;

        removeBreakoutSheets(linesArr, allSheets);

    });

};





async function deactivateEvents() {
    await Excel.run(async (context) => {

        context.runtime.load("enableEvents");

        await context.sync();

        context.runtime.enableEvents = false;
        console.log("Events: OFF - Occured in registerOnActivateHandler");

    });

};

async function activateEvents() {
    await Excel.run(async (context) => {

        context.runtime.load("enableEvents");

        await context.sync();

        context.runtime.enableEvents = true;
        console.log("Events: ON - Occured in registerOnActivateHandler");

    });

};

async function reset() {

    try {

        await Excel.run(async (context) => {

            //#region DEFINE FUNCTION VARIABLES --------------------------------------------------------------------------------------------------------

            //load in worksheets
            const silkE2RSheet = context.workbook.worksheets.getItem("SilkE2R").load("name");
            const textE2RSheet = context.workbook.worksheets.getItem("TextE2R").load("name");
            const digE2RSheet = context.workbook.worksheets.getItem("DIGE2R").load("name");
            const masterSheet = context.workbook.worksheets.getItem("Master").load("name");


            //load in tables
            const silkE2RTable = silkE2RSheet.tables.getItem("SilkE2R");
            const textE2RTable = textE2RSheet.tables.getItem("TextE2R");
            const digE2RTable = digE2RSheet.tables.getItem("DIGE2R");
            const masterTable = masterSheet.tables.getItem("Master");


            //loads the data body range of the tables above
            const silkE2RBodyRange = silkE2RTable.getDataBodyRange().load("values");
            const textE2RBodyRange = textE2RTable.getDataBodyRange().load("values");
            const digE2RBodyRange = digE2RTable.getDataBodyRange().load("values");
            const masterBodyRange = masterTable.getDataBodyRange().load("values");


            //loads the row items for specific tables
            const silkE2RTableRows = silkE2RTable.rows.load("items");
            const textE2RTableRows = textE2RTable.rows.load("items");
            const digE2RTableRows = digE2RTable.rows.load("items");
            const masterTableRows = masterTable.rows.load("items");


            //#endregion -------------------------------------------------------------------------------------------------------------------------------

            await context.sync();

            //#region LOAD FUNCTION VARIABLES ----------------------------------------------------------------------------------------------------------

            let silkE2RArr = silkE2RBodyRange.values; //moves all values of the SilkE2R table to an array
            let textE2RArr = textE2RBodyRange.values; //moves all values of the TextE2R table to an array
            let digE2RArr = digE2RBodyRange.values; //moves all values of the DIGE2R table to an array
            let masterArr = masterBodyRange.values;

            //#endregion -------------------------------------------------------------------------------------------------------------------------------

            let silkE2RClear = clearE2R(silkE2RArr, silkE2RTableRows, silkE2RSheet);

            silkE2RBodyRange.values = silkE2RClear;

            let textE2RClear = clearE2R(textE2RArr, textE2RTableRows, textE2RSheet);

            textE2RBodyRange.values = textE2RClear;

            let digE2RClear = clearE2R(digE2RArr, digE2RTableRows, digE2RSheet);

            digE2RBodyRange.values = digE2RClear;

            let masterClear = clearE2R(masterArr, masterTableRows, masterSheet);

            masterBodyRange.values = masterClear;


        });

        location.reload();

    } catch (err) {
        console.error(err);
        // showMessage(error, "show");
    };

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#endregion -----------------------------------------------------------------------------------------------------------------------------------------










//#region FUNCTIONS ----------------------------------------------------------------------------------------------------------------------------------


//#region PUSH BODY RANGE VALUES TO SINGLE ARRAY ---------------------------------------------------------------------------------------------------

/**
* Pushes the values from an array of arrays (typically the values from a body range) to a single array
* @param {Array} bodyRangeValues An array of arrays containing the table values from the body range
* @param {Array} arrayToPushTo A single array you wish to push the body range values to
*/
function pushBodyRangeValuesToArray(bodyRangeValues, arrayToPushTo) {

    for (let row of bodyRangeValues) {
        arrayToPushTo.push(row[0]);
    };

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region CREATE DATA SET --------------------------------------------------------------------------------------------------------------------------

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
                priority: priorityNum,
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

            priorityNum++;

        };

    };

    return dataSet;

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region PUSH TO ONE BIG ARRAY --------------------------------------------------------------------------------------------------------------------

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

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region CREATE ARRAY FROM OBJECT -----------------------------------------------------------------------------------------------------------------

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

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region REPLACE CHARACTER(S) ---------------------------------------------------------------------------------------------------------------------

/**
 * Replaces a character or set of characters in a string with another user defined character or set of characters and returns the altered string
 * @param {String} string The string that you wish to replace a certain character in
 * @param {String} replaceThis The character or set of characters that you wish to replace
 * @param {String} withThis The character or set of characters that you wish to substitute
 * @returns String
 */
function replaceCharacter(string, replaceThis, withThis) {
    if (string.includes(replaceThis)) {
        string = string.replace(replaceThis, withThis);
    };
    return string;
};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region NEXT PAGE FUNCTION FOR NAVIGATION ARROW CLICKS -------------------------------------------------------------------------------------------

/**
 * Returns the next or previous page or day in the week
 * @param {String} currentPage The currently shown page or day of the week table
 * @param {String} prevOrNext Do we move "forward" or "back" in the pages or days of the week table
 * @param {Array} arrOfPages Array of all the page nbames or days of the week
 */
function nextPage(currentPage, prevOrNext, arrOfPages) {

    // Get the index of the current weekday
    const thisIndex = arrOfPages.indexOf(currentPage);

    // If it's forward or backward...
    switch (prevOrNext) {

        case "forward": // Show the next weekday
            if (thisIndex + 1 === arrOfPages.length) { // Current is Friday
                return arrOfPages[0] // return Monday
            } else {
                return arrOfPages[thisIndex + 1]
            }
            break;

        case "back":  // Show the previous weekday

            if (thisIndex == 0) { // Current is Monday
                return arrOfPages[arrOfPages.length - 1] // return Friday
            } else {
                return arrOfPages[thisIndex - 1]
            }

            // Show the previous weekday
            break;
    };

    throw new Error("Could not figure out the next or previous page");

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region BUILD SELECT BOXES IN TABULATOR TABLE ----------------------------------------------------------------------------------------------------

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

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region BUILD TABULATOR TABLES -------------------------------------------------------------------------------------------------------------------

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
                    return `<div class="form-title">
            Form ${cell.getData().form}
          </div>`
                }
            },


            { title: "Priority", field: "priority", visible: false, sorter: "number", headerSort: false },

            {
                title: "DAY", field: "day", headerSort: false, widthGrow: 1, formatter: selectFormatter, formatterParams: { tableType: form }, cellClick: function (e, cell) {
                    // console.log(cell.getData().day);
                    // $(".select-box").addClass("select-arrow-active");         
                }
            },

            {
                title: "PRESS", field: "press", headerSort: false, widthGrow: 1, formatter: selectFormatter, formatterParams: { tableType: form }, cellClick: function (e, cell) {
                    // console.log(cell.getData().press);
                    // $(".select-box").addClass("select-arrow-active");         
                }
            },

            {
                formatter: "buttonCross", field: "clear", width: 30, hozAlign: "center", cellClick: function (e, cell) {

                    // Get the current row number
                    let cellData = cell.getRow().getData();
                    const rowNum = cellData.id;
                    const formNum = cellData.form;

                    // $(`select[rowId=${rowNum}]`)
                    // .silk-form-box[rowid='1']

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
                        silkDataSet = zeNewScheduleData;
                    } else if (whichOne2 == "Text") {
                        textDataSet = zeNewScheduleData;
                    } else if (whichOne2 == "Digital") {
                        digDataSet = zeNewScheduleData;
                    };

                    organizeData();

                    pressSchedulingInfoTable("Taskpane");




                    // console.log(cellData);
                    // cellData.day = "";

                }
            },


            // {title:"OPERATOR", field:"operator", headerSort:false, widthGrow: 1, formatter:selectFormatter, formatterParams:{tableType:form}, cellClick:function(e, cell){
            //   // console.log(cell.getData().operator);
            //   // $(".select-box").addClass("select-arrow-active");         
            // }},
        ],

        // rowMoved: (row) => {
        //   console.log("ROW MOVED!!!!!", row)
        // }

    });

    // // DATA CHANGED
    // scheduleTable.on("dataProcessed", (data) => {

    //   console.log(scrollHeight);

    // //   // deactivateEvents();


    // //   

    // //   // Update the schedule data table in excel to match data object above

    // });

    // USER MOVES THE ROW
    scheduleTable.on("rowMoved", (row) => {
        console.log("FIRED")
        // deactivateEvents();

        rowMoved = true;

        let newScheduleData = scheduleTable.getData();

        let pN = newScheduleData[0].priority


        // console.log("DATA BEFORE PRIORITY", newScheduleData)

        // Handle if moved to 1st position
        if (row.getPosition() == 1) {
            pN = newScheduleData[1].priority// Priority number
        }

        for (let index = 0; index < newScheduleData.length; index++) {

            newScheduleData[index].priority = pN;

            pN++;

        };

        // console.log("DATA AFTER PRIORITY", newScheduleData);


        scheduleTable.setData(newScheduleData);

        scrollErr.scrollTop = scrollHeight;
        // .then(() => {

        // console.log("New data has been set. Now I need to update the Excel table to match the new data...")
        const whichOne = newScheduleData[0].type
        // console.log("whichOne:", whichOne);


        if (whichOne == "Silk") {
            silkDataSet = newScheduleData;
        } else if (whichOne == "Text") {
            textDataSet = newScheduleData;
        } else if (whichOne == "Digital") {
            digDataSet = newScheduleData;
        }

        // activateEvents();


        organizeData();

        pressSchedulingInfoTable("Taskpane");

        // })

        console.log("ROW MOVED!!!!!!!!!", row.getData().form);

        // pressSchedulingInfoTable("Taskpane");

        // scrollErr.scrollTop = scrollHeight;

        // activateEvents();

    })

    return scheduleTable;

    // activateEvents();

};



//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region SELECT BOX FORMATTER BY ROW --------------------------------------------------------------------------------------------------------------

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
            thisSelect = buildSelect(daysOfWeek, cell, tableType);
            break;
        case "operator":
            // Loop the operator
            thisSelect = buildSelect(pressmen, cell, tableType);
            break;
        case "press":
            // Loop the presses
            thisSelect = buildSelect(presses, cell, tableType);
            break;
    };

    let myDiv = $(`<div class="my-div ${tableType}-drop"></div>`)
    myDiv.append(thisSelect);
    return myDiv.prop(`outerHTML`);

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region DIVIDE TABULATOR TABLE DATA INTO WEEKDAYS AND CALL ORGANIZE, SORT, AND BUILD STATIC HTML TABLE FUNCTIONS FOR EACH ------------------------

/**
* Divides tabulator data up by the weekdays and then tosses it into functions to organize it for making static HTML tables
*/
function organizeData() {
    // const organizeData = (scheduleTable) => {

    monday = [];
    tuesday = [];
    wednesday = [];
    thursday = [];
    friday = [];

    operators = {};

    // $("#monday-form").empty();//.append(`<table id="pressman"></table>`);

    $("monday-static").empty();
    $("tuesday-static").empty();
    $("wednesday-static").empty();
    $("thursday-static").empty();
    $("friday-static").empty();

    pressmen.forEach((man) => {
        operators[man] = [];
    });

    dotSelected = false;


    // let silkData = silkTable.getData();
    // let textData = textTable.getData();
    // let digData = digTable.getData();






    makeTablesForEachDOW(silkDataSet);
    makeTablesForEachDOW(textDataSet);
    makeTablesForEachDOW(digDataSet);

    // makeTablesForEachOperator(silkDataSet);
    // makeTablesForEachOperator(textDataSet);
    // makeTablesForEachOperator(digDataSet);





    // console.log(`Monday is: `);
    // console.log(monday);
    // console.log(`Tuesday is: `);
    // console.log(tuesday);
    // console.log(`Wednesday is: `);
    // console.log(wednesday);
    // console.log(`Thursday is: `);
    // console.log(thursday);
    // console.log(`Friday is: `);
    // console.log(friday);

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region MAKE TABLES FOR EACH DOW -----------------------------------------------------------------------------------------------------------------

/**
* Is feed either the Silk, Text, or Digital tabulator table data and organizes, sorts, and creates the static HTML tables
* @param {Object} scheduleData An object containing all the data from the specific tabulator table to use in creating the static weekday tables
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
            createWeekTables(monday, row, tempObj, tempForm);

            //sorts organized data so the press numbers are in order
            sortArrOfObj(monday, "press");

            //builds the static HTML table based on the data from the previous 2 functions
            buildDOWTable(monday, "Monday");

            //keeps the week title the same if the selected dot hasn't changed and something else is updated in the tabulator tables
            if ($(`#monday-static`).hasClass("dot-selected")) {
                $("#week-title").text("MONDAY");
                dotSelected = true;
            }

        } else if (row.day == "Tuesday") {

            //organizes data to be used to build static HTML tables
            createWeekTables(tuesday, row, tempObj, tempForm);

            //sorts organized data so the press numbers are in order
            sortArrOfObj(tuesday, "press");

            //builds the static HTML table based on the data from the previous 2 functions
            buildDOWTable(tuesday, "Tuesday");

            //keeps the week title the same if the selected dot hasn't changed and something else is updated in the tabulator tables
            if ($(`#tuesday-static`).hasClass("dot-selected")) {
                $("#week-title").text("TUESDAY");
                dotSelected = true;
            }

        } else if (row.day == "Wednesday") {

            //organizes data to be used to build static HTML tables
            createWeekTables(wednesday, row, tempObj, tempForm);

            //sorts organized data so the press numbers are in order
            sortArrOfObj(wednesday, "press");

            //builds the static HTML table based on the data from the previous 2 functions
            buildDOWTable(wednesday, "Wednesday");

            //keeps the week title the same if the selected dot hasn't changed and something else is updated in the tabulator tables
            if ($(`#wednesday-static`).hasClass("dot-selected")) {
                $("#week-title").text("WEDNESDAY");
                dotSelected = true;
            }

        } else if (row.day == "Thursday") {

            //organizes data to be used to build static HTML tables
            createWeekTables(thursday, row, tempObj, tempForm);

            //sorts organized data so the press numbers are in order
            sortArrOfObj(thursday, "press");

            //builds the static HTML table based on the data from the previous 2 functions
            buildDOWTable(thursday, "Thursday");

            //keeps the week title the same if the selected dot hasn't changed and something else is updated in the tabulator tables
            if ($(`#thursday-static`).hasClass("dot-selected")) {
                $("#week-title").text("THURSDAY");
                dotSelected = true;
            }

        } else if (row.day == "Friday") {

            //organizes data to be used to build static HTML tables
            createWeekTables(friday, row, tempObj, tempForm);

            //sorts organized data so the press numbers are in order
            sortArrOfObj(friday, "press");

            //builds the static HTML table based on the data from the previous 2 functions
            buildDOWTable(friday, "Friday");

            //keeps the week title the same if the selected dot hasn't changed and something else is updated in the tabulator tables
            if ($(`#friday-static`).hasClass("dot-selected")) {
                $("#week-title").text("FRIDAY");
                dotSelected = true;
            }

        } else {
            // console.log("Don't know what to tell ya...")
        };

    });

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region ORGANIZE WEEK TABLE DATA -----------------------------------------------------------------------------------------------------------------

/**
 * Organizes the dow data into an array of objects that can be referenced later to build a HTML table
 * @param {Array} dowArr An array that starts out empty but fills with the data for each item that belongs it that specific day of the week
 * @param {Object} rowData An object of all the data in the current row
 * @param {Object} tempObj An object that starts out empty but gets filled with data before it is sorted into it's proper merged or unmerged data set
 * @param {Object} tempForm An object that starts out empty but gets filled with data about the different forms before being moved into it's proper data set
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

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region MAKE TABLES FOR EACH OPERATOR ------------------------------------------------------------------------------------------------------------

/**
* Is feed either the Silk, Text, or Digital tabulator table data and organizes, sorts, and creates the static HTML tables
* @param {Object} scheduleData An object containing all the data from the specific tabulator table to use in creating the static weekday tables
*/
function makeTablesForEachOperator(scheduleData) {

    let tempObj = {};
    let tempForm = {};
    let tempPress = {};

    opTableCounter = 1;

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

        createOperatorTables(operators[row.operator], row, tempObj, tempForm, tempPress);

        console.log(JSON.stringify(operators));

        //builds the static HTML table based on the data from the previous function
        buildOperatorTable(operators, row.operator);

        //keeps the week title the same if the selected dot hasn't changed and something else is updated in the tabulator tables
        if ($(`#monday-static`).hasClass("dot-selected")) {
            $("#week-title").text("MONDAY");
            dotSelected = true;
        }

    });

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region ORGANIZE OPERATOR TABLE DATA -------------------------------------------------------------------------------------------------------------

/**
 * Organizes the dow data into an array of objects that can be referenced later to build a HTML table
 * @param {Array} opArr An array that starts out empty but fills with the data for each item that belongs to a specific operator
 * @param {Object} rowData An object of all the data in the current row
 * @param {Object} tempObj An object that starts out empty but gets filled with data before it is sorted into it's proper merged or unmerged data set
 * @param {Object} tempForm An object that starts out empty but gets filled with data about the different forms before being moved into it's proper data set
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

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------

//#region SORT DATA ALPHABETICALLY -----------------------------------------------------------------------------------------------------------------

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

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region BUILD DAY OF THE WEEK TABLES -------------------------------------------------------------------------------------------------------------

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

        // let operatorString;

        // //if there are multipe operator assigned to a single press, 
        // //set array to string and replce commas with " & "
        // if (press.operator.length > 1) {
        //   operatorString = press.operator.toString();
        //   operatorString = operatorString.replaceAll(",", " & ");

        // //otherwise, just use the only instance in the array
        // } else {
        //   operatorString = press.operator[0];
        // }

        press.forms.forEach((form, formIndex) => { //for each form in each press in the dow data set...

            // ${formIndex === 0 
            //   ? `<td class="press" rowspan="${press.forms.length}">${press.press}</td>
            //       <td class="operators" rowspan=${press.forms.length}">${operatorString}</td>`
            //   : ``}

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

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------

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

    if (opTableCounter == 1) {
        $(`#${lowerOp}-static`).addClass("show-table");
    };

    opTableCounter++;

}



//#region PUSH DATA TO PRESS SCHEDULING TABLE IN VALIDATION ----------------------------------------------------------------------------

async function pressSchedulingInfoTable(trigger) {

    await Excel.run(async (context) => {

        const validation = context.workbook.worksheets.getItem("Validation");
        const pressSchedulingInfo = validation.tables.getItem("PressSchedulingInfo");
        const pressSchedulingInfoBodyRange = pressSchedulingInfo.getDataBodyRange().load("values");
        const pressSchedulingInfoHeaderRange = pressSchedulingInfo.getHeaderRowRange().load("values");
        const pressSchedulingInfoRows = pressSchedulingInfo.rows.load("count");


        await context.sync();


        let pressSchedulingArr = pressSchedulingInfoBodyRange.values;


        let silkSchedulingInfo = [];
        let textSchedulingInfo = [];
        let digSchedulingInfo = [];

        let oneBigArr = [];

        let updateDataWithArrInfo = false;


        //make this into it's own function, but it will need to be an async excel context function so I can load in the press scheduling info 
        //table values and write to them. Also need to add functions that write tabulator data set info when updated to this val table and 
        //another function when populate forms is clicked and there is data in the val table to macth the type and form in the val table to 
        //the E2R data and only update the E2R related info while keeping the tabualtor info intact and to the right form

        //#region ON POPULATE FORMS BUTTON PRESS -------------------------------------------------------------------------------------------------------

        if (trigger == "Populate") {

            let emptyCell = false;
            let silkEmpty = true;
            let textEmpty = true;
            let digitalEmpty = true;

            //sets emptyCell to true if the press scheduling info table's first row in completely blank (meaning that the table is empty)

            // let row = ["", "", "", "", ""]

            // var isIt = row.some(function (item) {
            //   return item != "";
            // });

            // True
            // console.log("Are any of them not blank?", isIt);

            for (let cell of pressSchedulingArr[0]) {
                if (cell == "") {
                    emptyCell = true;
                } else {
                    emptyCell = false;
                    break;
                }
            };

            // for (let cell of pressSchedulingArr[0]) {
            //   if (cell == "Silk") {
            //     silkEmpty = false;
            //   };

            //   if (cell == "Text") {
            //     textEmpty = false;
            //   };

            //   if (cell == "Digital") {
            //     digitalEmpty = false;
            //   };
            // };

            //empties these variables just in case
            silkSchedulingInfo = [];
            textSchedulingInfo = [];
            digSchedulingInfo = [];

            //#region ON TABLE EMPTY -------------------------------------------------------------------------------------------------------------------

            //if the press scheduling table is only 1 row and said row is empty, simply add the E2R data to it
            if (pressSchedulingArr.length === 1 && emptyCell == true) {

                console.log("Press Scheduling Info table is empty!");

                silkSchedulingInfo = createArrFromObj(silkDataSet);
                textSchedulingInfo = createArrFromObj(textDataSet);
                digSchedulingInfo = createArrFromObj(digDataSet);

                //empties this array everytime this loop goes through
                oneBigArr = [];

                pushToBigArr(silkSchedulingInfo, oneBigArr);
                pushToBigArr(textSchedulingInfo, oneBigArr);
                pushToBigArr(digSchedulingInfo, oneBigArr);


                // let firstEmptyRow = pressSchedulingInfo.rows.getItemAt(0);
                // firstEmptyRow.delete();

                await context.sync();

                pressSchedulingInfo.rows.add(
                    null,
                    oneBigArr,
                    true,
                );

            } else {

                //#endregion -------------------------------------------------------------------------------------------------------------------------------


                //#region IF TABLE IS NOT EMPTY ------------------------------------------------------------------------------------------------------------

                //if val table is not empty and populate forms button is pressed, need to match type and form in val table for each row to the new 
                //dataSet. Upon match, update just the E2R info for that specific line in the val table array (leaving the tabulator table info intact). 
                //Then, write table array to Excel and let the last function to update the dataSets and taskpane play out.

                console.log("populate was pressed and the table is not empty!!!");

                // if (silkEmpty) {

                //   silkSchedulingInfo = createArrFromObj(silkDataSet);

                //   //empties this array everytime this loop goes through
                //   oneBigArr = [];

                //   pushToBigArr(silkSchedulingInfo, oneBigArr);


                //   // let firstEmptyRow = pressSchedulingInfo.rows.getItemAt(0);
                //   // firstEmptyRow.delete();

                //   await context.sync();

                //   pressSchedulingInfo.rows.add(
                //     null,
                //     oneBigArr,
                //     true,
                //   );

                // } else {

                let i = 0;
                let cheeseSoup = [];
                let snailGloves = [];

                let pressSchedArrUpdated = [];
                let newPressSchedArrUpdated = [];
                // for (let i = 0; i < pressSchedulingArr.length; i++) {

                let silkInfo = matchRewrite(pressSchedulingArr, i, silkDataSet, "Silk");

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

                // }


                // for (let row of pressSchedulingArr) {
                //   if (row[0] !== "Silk") {
                //     cheeseSoup.push(row);
                //   };
                // };

                // pressSchedArrUpdated.push(silkArr[0]);
                // pressSchedArrUpdated.push(cheeseSoup[0]);

                let textInfo = matchRewrite(pressSchedArrUpdated, i, textDataSet, "Text");

                let textArr = textInfo.arr;

                i = textInfo.index;


                silkArr.forEach((row) => {
                    newPressSchedArrUpdated.push(row);
                });

                textArr.forEach((row) => {
                    newPressSchedArrUpdated.push(row);
                });

                pressSchedulingArr.forEach((row) => {

                    if (row[0] !== "Silk" && row[0] !== "Text") {
                        newPressSchedArrUpdated.push(row);
                    };

                });

                // for (let po of pressSchedArrUpdated) {
                //   if (po[0] !== "Silk" && po[0] !== "Text") {
                //     snailGloves.push(po);
                //   };
                // };

                // newPressSchedArrUpdated.push(silkArr);
                // newPressSchedArrUpdated.push(textArr);
                // newPressSchedArrUpdated.push(snailGloves);

                let digInfo = matchRewrite(newPressSchedArrUpdated, i, digDataSet, "Digital");

                let digArr = digInfo.arr;

                i = digInfo.index;


                let bigBoi = [];

                pushToBigArr(silkArr, bigBoi);
                pushToBigArr(textArr, bigBoi);
                pushToBigArr(digArr, bigBoi);



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


            };

            //#endregion -------------------------------------------------------------------------------------------------------------------------------

        };

        //#endregion -----------------------------------------------------------------------------------------------------------------------------------


        //#region ON TASKPANE CHANGE -------------------------------------------------------------------------------------------------------------------

        if (trigger == "Taskpane") {

            //replace table array with new dataSet info and then push to excel table


            // console.log("Going to update the val table based on the taskpane dataSet");

            let rowCount = pressSchedulingInfoRows.count - 1;

            pressSchedulingInfoRows.deleteRowsAt(0, rowCount);

            // silkDataSet = [];
            // textDataSet = [];
            // digDataSet = [];



            silkSchedulingInfo = createArrFromObj(silkDataSet);
            textSchedulingInfo = createArrFromObj(textDataSet);
            digSchedulingInfo = createArrFromObj(digDataSet);

            oneBigArr = [];

            pushToBigArr(silkSchedulingInfo, oneBigArr);
            pushToBigArr(textSchedulingInfo, oneBigArr);
            pushToBigArr(digSchedulingInfo, oneBigArr);


            // let firstEmptyRow = pressSchedulingInfo.rows.getItemAt(0);
            // firstEmptyRow.delete();

            await context.sync();

            pressSchedulingInfoRows.getItemAt(0).delete();

            // console.log("farts");


            await context.sync();

            pressSchedulingInfo.rows.add(
                null,
                oneBigArr,
                true,
            );


        };

        //#endregion -----------------------------------------------------------------------------------------------------------------------------------


        // if (trigger == "Table") {

        //   //might not need this trigger, since all I need to do is replace dataSet with table array info

        //   console.log("Press Scheduling Table in the Validation was updated!");

        //   updateDataWithArrInfo = true;

        // }

        //if val table changes, replace dataSet info with table array info

        //if table is not empty, replace dataSet info with table array info

        if (trigger !== "Taskpane") {

            const pressSchedUpdateBodyRange = pressSchedulingInfo.getDataBodyRange().load("values");

            await context.sync();

            let pressSchedUpdate = pressSchedUpdateBodyRange.values;


            //this should only fire when the val table changes and after the content is matched when values exist in table when populate button is pressed

            updateDataFromTable(pressSchedUpdate);



            silkTable = buildTabulatorTables("silk-form", silkTable, silkDataSet);
            textTable = buildTabulatorTables("text-form", textTable, textDataSet);
            digTable = buildTabulatorTables("dig-form", digTable, digDataSet);


            organizeData();

        };


        refreshPivotTable();


    });


    activateEvents();

    scrollErr.scrollTop = scrollHeight;

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------

//#endregion ---------------------------------------------------------------------------------------------------------------------------




async function E2RHandler(event) {

    deactivateEvents();

    await Excel.run(async (context) => {

        //#region HANDLE REMOTE CHANGES ----------------------------------------------------------------------------------------------------------------

        // console.log("Source of the E2RHandler event: " + event.source);

        if (event.source == "Remote") {
            console.log("Content was changed by a remote user, exiting E2RHandler Event");
            return;
        };

        //#endregion -----------------------------------------------------------------------------------------------------------------------------------

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

        await context.sync();

        let pressScheduleArr = pressSchedulingBodyRange.values;
        let silkE2RArr = silkE2RBodyRangeUpdate.values;
        let textE2RArr = textE2RBodyRangeUpdate.values;
        let digE2RArr = digE2RBodyRangeUpdate.values;

        let tableContent = bodyRange.values;
        let head = headerRange.values;


        let tableRowIndex = changedAddress.rowIndex - 2;
        let changedRowValues = changedTableRows.items[tableRowIndex].values

        let changedRowForm = changedRowValues[0][1];
        let changedRowDay = changedRowValues[0][5];
        let changedRowPress = changedRowValues[0][6];



        if (changedWorksheet.name == "SilkE2R") {

            for (let rowIndex in pressScheduleArr) {
                if ((pressScheduleArr[rowIndex][0] == "Silk") && (pressScheduleArr[rowIndex][2] == changedRowForm)) {
                    pressScheduleArr[rowIndex][6] = changedRowDay;
                    pressScheduleArr[rowIndex][7] = changedRowPress;
                    break;
                };
            };
            // silkDataSet = [];
            // silkDataSet = createDataSet(silkE2RArr, "Silk");

        };

        if (changedWorksheet.name == "TextE2R") {

            for (let rowIndex in pressScheduleArr) {
                if ((pressScheduleArr[rowIndex][0] == "Text") && (pressScheduleArr[rowIndex][2] == changedRowForm)) {
                    pressScheduleArr[rowIndex][6] = changedRowDay;
                    pressScheduleArr[rowIndex][7] = changedRowPress;
                    break;
                };
            };
            // textDataSet = [];
            // textDataSet = createDataSet(textE2RArr, "Text");

        };

        if (changedWorksheet.name == "DIGE2R") {

            for (let rowIndex in pressScheduleArr) {
                if ((pressScheduleArr[rowIndex][0] == "Digital") && (pressScheduleArr[rowIndex][2] == changedRowForm)) {
                    pressScheduleArr[rowIndex][6] = changedRowDay;
                    pressScheduleArr[rowIndex][7] = changedRowPress;
                    break;
                };
            };
            // digDataSet = [];
            // digDataSet = createDataSet(digE2RArr, "Digital");

        };

        pressSchedulingBodyRange.values = pressScheduleArr;

        await context.sync();

        updateDataFromTable(pressScheduleArr);

        // pressSchedulingInfoTable("Taskpane");

        silkTable = buildTabulatorTables("silk-form", silkTable, silkDataSet);
        textTable = buildTabulatorTables("text-form", textTable, textDataSet);
        digTable = buildTabulatorTables("dig-form", digTable, digDataSet);

        organizeData();

        scrollErr.scrollTop = scrollHeight;

        console.log("E2RHandler was fired, which updated both the Press Scheduling Info table and the Taskpane");

        // console.log("pressSchedulerHandler scrollHeight:", scrollHeight);

        refreshPivotTable();

        //need to get the column indexs for the forms column, form quantity, sheets, and hours columns in the changed sheet
        //then need to get row range of the row that had content updated in the E2R
        //need to grab the values for that row in all of the columns listed above and assign them variables
        //need to take the form number variable and loop through the press scheduling table in the validation until the form number variable matchs the form number in the table
        //update the info in the matching columns for that row in the press scheduling table
        //do all of this WITHOUT deativating events so that the event to update the taskpane info will also fire


        //then I just need to figure out how to make it go the other way around. Maybe what I need to do is deactivate the event that updates the E2R from the press scheduler before I run the E2R function. Then I reactivate it after the function finishes. Same thing the other way around, just deactivate the pressSchedulerHandler event before running the the one that pushes the stuff to the E2R. Or something.


    });

    activateEvents();

};


async function pressSchedulerHandler(event) {

    deactivateEvents();

    await Excel.run(async (context) => {

        // if (rowMoved) {
        //   console.log("Row was moved so not running pressSchedulerHandler");
        //   // rowMoved = false;
        //   return;
        // };

        console.log("val table was changed!");

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


        await context.sync();


        let pressSchedArr = pressSchedulingBodyRange.values;

        let tableContent = bodyRange.values;
        let head = headerRange.values;

        //changedAddress.rowIndex is from a worksheet level, which is 0 indexed. In the worksheet, the first row (0) is the title row, then the next 
        //row (1) is the header. The content doesn't start until (2). However, we want the row index according to the table, which would have the first 
        //row start at 0. Since the title and header will always be the way they are, we can simply subtract 2 from the worksheet row index to get the 
        //table row index
        let tableRowIndex = changedAddress.rowIndex - 2;


        let changedRowValues = changedTableRows.items[tableRowIndex].values

        //just to make it easier on myself for now, I am assuming that the position of the columns will not change, so I am referencing the order in which
        //the content should be. However, if I ever want to go back and make it more dynamic with column placement, then I will need to load in the headers
        //and stuff like I am doing in the art queue and assign each feild in an object to each header and then create an array of object for every row in
        //the table. That's a lot of effort for now (even though i have done it before), so I am just not going to do that for now and assume that the
        //columns will never move

        let changedRowType = changedRowValues[0][0];
        let changedRowForm = changedRowValues[0][2];
        let changedRowDay = changedRowValues[0][6];
        let changedRowPress = changedRowValues[0][7];

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

        const E2RTable = context.workbook.tables.getItem(E2RTableName);
        const E2RBodyRange = E2RTable.getDataBodyRange().load("values");

        const E2RRows = E2RTable.rows;
        E2RRows.load("items");

        await context.sync();

        const E2RValues = E2RBodyRange.values;

        for (let rowIndex in E2RValues) {

            let E2RForm = E2RValues[rowIndex][1];

            if (changedRowForm == E2RForm) {
                E2RValues[rowIndex][5] = changedRowDay;
                E2RValues[rowIndex][6] = changedRowPress;
                break;
            };

        };

        E2RBodyRange.values = E2RValues;

        await context.sync();


        updateDataFromTable(pressSchedArr);

        await context.sync();

        silkTable = buildTabulatorTables("silk-form", silkTable, silkDataSet);
        textTable = buildTabulatorTables("text-form", textTable, textDataSet);
        digTable = buildTabulatorTables("dig-form", digTable, digDataSet);

        organizeData();

        scrollErr.scrollTop = scrollHeight;

        console.log("pressSchedulerHandler was fired, which updated the E2R associated with the changed value and the Taskpane");

        // console.log("pressSchedulerHandler scrollHeight:", scrollHeight);

        refreshPivotTable();

    });

    activateEvents();


};







function matchRewrite(tableArray, tableRowIndex, dataSet, tableType) {

    let newArr = [];
    let leTempArr = [];

    for (let rowNum in dataSet) {

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
        } else {

            if (tableArray[tableRowIndex][0] == tableType) {

                leTempArr.push(dataSet[rowNum].type);
                leTempArr.push(dataSet[rowNum].priority);
                leTempArr.push(dataSet[rowNum].form);
                leTempArr.push(dataSet[rowNum].formQuantity);
                leTempArr.push(dataSet[rowNum].sheets);
                leTempArr.push(dataSet[rowNum].hours);
                leTempArr.push(tableArray[tableRowIndex][6]);
                leTempArr.push(tableArray[tableRowIndex][7]);

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
                console.log(`The val table's type (${tableType}) does not match the new data's type (${tableArray[tableRowIndex][0]}).`);
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
            };
        };

        leTempArr = [];

    };

    return {
        arr: newArr,
        index: tableRowIndex
    };

};



//#region REPLACE OBJECT INFO IN DATA SETS WITH TABLE ARRAY INFO -------------------------------------------------------------------------------

/**
 * Replaces the object info in each of the different data set types with the info from the press scheduling info table
 * @param {Array} tableArray The array of the values of the Press Scheduling Info table
 */
function updateDataFromTable(tableArray) {

    // if (updateDataWithArrInfo == true) {

    // let tId = 1;

    //empty data current existing in data seta

    silkDataSet = []; let sI = 1;
    textDataSet = []; let tI = 1;
    digDataSet = []; let dI = 1;

    let emptyCell = false;

    for (let cell of tableArray[0]) {
        if (cell == "") {
            emptyCell = true;
        } else {
            emptyCell = false;
            break;
        }
    };

    if (tableArray.length === 1 && emptyCell == true) {

        console.log("Press Scheduling Info table has been emptied!");

        silkDataSet = [];
        textDataSet = [];
        digDataSet = [];

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
            silkDataSet.push(zeObj);
            sI++;

        };

        if (tableArray[t][0] == "Text") {
            zeObj.id = tI;
            textDataSet.push(zeObj);
            tI++;
        };

        if (tableArray[t][0] == "Digital") {
            zeObj.id = dI;
            digDataSet.push(zeObj);
            dI++;
        };

        priorityNum++;



        // tId++;

    };


    // console.log("silkDataSet", silkDataSet);
    // console.log("textDataSet", textDataSet);
    // console.log("digDataSet", digDataSet);

    // };

};

//#endregion -------------------------------------------------------------------------------------------------------------------------------------







//#region COMPARE MASTER TO E2R FUNCTION -----------------------------------------------------------------------------------------------------------

/**
 * Compares the Client Code and Product Abbrevations from the current row in the Master Sheet to the current row in the current E2R table.
 * @param {Array} masterProdAbbr An array of all the product abbreveations listed for the current line's product in the Master sheet
 * @param {String} compareToProd The product (already abbreviated) from the current line in the E2R table to compare to the Master Product
 * @param {String} compareToCode The client code from the current line in the E2R table to compare to the Master Client Code
 * @param {String} masterCode The client code from the current line in the Master sheet 
 * @returns Boolean
 */
function compareMasterToE2R(masterProdAbbr, compareToProd, compareToCode, masterCode) {

    for (var k = 0; k < masterProdAbbr.length; k++) {
        if (compareToProd == masterProdAbbr[k] && compareToCode == masterCode) {
            return true;
        };
    }
    return false;

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region OBJECTIFYING ROW INFO --------------------------------------------------------------------------------------------------------------------

//#region CREATE ROW INFO FUNCTION ---------------------------------------------------------------------------------------------------------------

/**
 * Matches the table headers with the data and column index in the current row, assigning each as a property of each header within the empty obj 
 * loaded in by the user
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

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region POPULATE EASY TO READS -------------------------------------------------------------------------------------------------------------------

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

        // if (row[0] == "") {
        //   firstNumber = "";
        // };

        //if the row starts with the word "Layout", we know that a new form is starting. Or "Form" if it's looking at Digital
        if (row[0].startsWith("Layout") || row[0].startsWith("Form") || row[0].startsWith("Tube")) {

            firstNumber = row[0].match(/[0-9]+/); //the first number in this cell will always be the form number, so let's grab that

            if (firstNumber) { //if a form number exists, write the value to the form variable for use later

                form = firstNumber[0];

                let fQ;

                //#region ASSIGNING WASTE VALUES -------------------------------------------------------------------------------------------------------

                let waste = 0;

                //if the worksheet's name is TextE2R, we need to offset the forms by 50, so we turn form into a number and add 50 to the result
                //Also defining the waste amount for non-digital text and silk products
                if (worksheet.name == "TextE2R") {

                    form = Number(form) + 50; //augment form by 50
                    waste = wasteData["Text"]["Waste"]; //text waste

                } else if (worksheet.name == "SilkE2R") {

                    waste = wasteData["Silk"]["Waste"]; //silk waste

                } else if (worksheet.name == "DIGE2R") {

                    if (Number(form) > 100 && Number(form) < 150) { //digital text

                        waste = wasteData["Digital Text"]["Waste"];

                    } else if (Number(form) > 150 && Number(form) < 200) { //digital silk

                        waste = wasteData["Digital Silk"]["Waste"];

                    } else if (Number(form) > 200 && Number(form) < 250) { //digital husky

                        waste = wasteData["Digital Text"]["Waste"];

                    } else if (Number(form) > 300 && Number(form) < 350) { //envelope windows

                        waste = wasteData["Digital Text"]["Waste"];

                    } else if (Number(form) > 350 && Number(form) < 400) { //envelope no windows

                        waste = wasteData["Digital Text"]["Waste"];

                    } else if (Number(form) > 400 && Number(form) < 500) { //wide format tubes

                        waste = wasteData["Wide Format"]["Waste"];

                    } else if (Number(form) > 500 && Number(form) < 550) { //digital variable silk

                        waste = wasteData["Digital Silk"]["Waste"];

                    } else if (Number(form) > 550 && Number(form) < 600) { //digital variable text

                        waste = wasteData["Digital Text"]["Waste"];

                    };

                };

                //#endregion ---------------------------------------------------------------------------------------------------------------------------

                singleSided = false;

                const numSidedArr = ["1-sided", "1-Sided", "1-SIDED"]; //with number
                const oneSidedArr = ["One-Sided", "One-sided", "one-sided", "ONE-SIDED", "one-Sided"]; //without number

                if (!row[0].startsWith("Tube")) { //only do the following for all non-tube forms

                    //#region FORM QUANTITY AND 1-SIDED WASTE AUGMENTATION -------------------------------------------------------------------------------

                    let formQuantity = 0;

                    if (worksheet.name == "DIGE2R") {

                        //#region DIGITAL FORM QUANTITY AND 1-SIDED WASTE AUGMENT ------------------------------------------------------------------------

                        let startIndexOfFormNumber = row[0].indexOf(firstNumber); //index of the position where the form number starts

                        //adds the form number start index to the length of the number itself
                        let indexAfterForm = startIndexOfFormNumber + (firstNumber[0].length);

                        //isolates all the text after the form number to the end of the string
                        let textAfterForm = row[0].substring(indexAfterForm).trim();

                        //#region ACCOUNT FOR MULTIPLE "1-SIDED" SPELLINGS -----------------------------------------------------------------------------

                        // singleSided = false;

                        // const numSidedArr = ["1-sided", "1-Sided", "1-SIDED"]; //with number
                        // const oneSidedArr = ["One-Sided", "One-sided", "one-sided", "ONE-SIDED", "one-Sided"]; //without number

                        //#region TRY DIFFERENT SPELLINGS INCLUDING NUMBER ---------------------------------------------------------------------------

                        for (var nSItem of numSidedArr) {

                            // let theItem = numSidedArr[nSItem];

                            if (textAfterForm.startsWith(nSItem)) { //modifying waste if the form is a 1-sided form and also adjusts 

                                waste = Math.ceil(waste / 2);

                                textAfterForm = textAfterForm.replace(nSItem, "").trim(); //removes "1-sided" from the textAfterForm string

                                //the quantity at the end of the string will be the only number left after the form number & 1-sided is removed, 
                                //which we did above
                                formQuantity = (textAfterForm.match(/[0-9]+/))[0];

                                fQ = Number(formQuantity);

                                singleSided = true;

                            };

                        };

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#region TRY DIFFERENT SPELLINGS WITHOUT NUMBER -----------------------------------------------------------------------------

                        for (var oSItem of oneSidedArr) { //just adjusts waste for single sided & finds form quantity 

                            // let zeItem = oneSidedArr[oSItem];

                            if (textAfterForm.startsWith(oSItem)) {

                                waste = Math.ceil(waste / 2);

                                //no need to remove "One-Sided" from string since there is no number present other than the quantity now
                                formQuantity = (textAfterForm.match(/[0-9]+/))[0];

                                fQ = Number(formQuantity);

                                singleSided = true;

                            };

                        };

                        //#endregion -----------------------------------------------------------------------------------------------------------------

                        //#endregion -------------------------------------------------------------------------------------------------------------------                    

                        //#endregion -----------------------------------------------------------------------------------------------------------------------

                        //#region FORM QUANTITY FOR DOUBLE SIDED DIGITAL -----------------------------------------------------------------------------------

                        if (singleSided == false) {

                            formQuantity = (textAfterForm.match(/[0-9]+/))[0];
                            fQ = Number(formQuantity);

                            singleSided = false;

                            // fQ = fQ * 2;

                        };

                        //#endregion -----------------------------------------------------------------------------------------------------------------------

                    } else {

                        //#region TEXT AND SILK FORM QUANTITY AND 1-SIDED WASTE AUGMENT --------------------------------------------------------------------

                        //finds the word before "House Stock". The ./s+ accounts for all spaces before House Stock and a single character (the comma)
                        const houseStockRegex = /\w+(?=.\s+House Stock)/;
                        var wordBeforeHouseStock = row[0].match(houseStockRegex);

                        if (wordBeforeHouseStock == "Sheetwise") { //all double sided forms should use the word "Sheetwise" before "House Stock"

                            //#region DOUBLE-SIDED FORMS (STANDARD) ----------------------------------------------------------------------------------------

                            //slice takes the characters between a start index and end index. The start has a +3 so that way the result does not 
                            //include the 3 characters we are looking for. The end has a -2 so that it does not include ", " characters after qty.
                            formQuantity = row[0].slice((row[0].indexOf("), ") + 3), (row[0].indexOf("Sheetwise") - 2));

                            fQ = Number(formQuantity.replace(/,/g, "")); //removes all commas and converts from string to number

                            singleSided = false;

                            // fQ = fQ * 2;

                            //#endregion -------------------------------------------------------------------------------------------------------------------

                        } else if (wordBeforeHouseStock == "Perfected") { //means it is using the new press and should cut waste by 200 for both text & silk and half the hours

                            //#region PERFECTED FORMS ------------------------------------------------------------------------------------------------------

                            //slice takes the characters between a start index and end index. The start has a +3 so that way the result does not 
                            //include the 3 characters we are looking for. The end has a -2 so that it does not include ", " characters after qty.
                            formQuantity = row[0].slice((row[0].indexOf("), ") + 3), (row[0].indexOf("Perfected") - 2));

                            fQ = Number(formQuantity.replace(/,/g, "")); //removes all commas and converts from string to number

                            singleSided = false;

                            waste = waste - 200;

                            //hours will be halved down below outside of this area

                            // fQ = fQ * 2;

                            //#endregion -------------------------------------------------------------------------------------------------------------------

                        } else {

                            //#region SINGLE SIDED FORMS (ABNORMAL) ----------------------------------------------------------------------------------------

                            //try for any other variable other than Sheetwise that should appear. This one also tells us if it's a 1-sided form
                            //If this fails, catch the error
                            try {

                                //slice takes the characters between a start index and end index. The start has a +3 so that way the result does not 
                                //include the 3 characters we are looking for. The end has a -2 so that it does not include ", " characters after qty.
                                for (let numSidedItem of numSidedArr) {
                                    if (row[0].includes(numSidedItem)) {
                                        formQuantity = row[0].slice((row[0].indexOf("), ") + 3), (row[0].indexOf(numSidedItem) - 2));
                                    };
                                };

                                for (let oneSidedItem of oneSidedArr) {
                                    if (row[0].includes(oneSidedItem)) {
                                        formQuantity = row[0].slice((row[0].indexOf("), ") + 3), (row[0].indexOf(oneSidedItem) - 2));
                                    };
                                };


                                //if this try succeeds, then we know this is a single sided form and the waste needs to be cut in half
                                waste = Math.ceil(waste / 2); //Always round up for paper purposes

                                fQ = Number(formQuantity.replace(/,/g, "")); //removes all commas and converts from string to number

                                singleSided = true;

                            } catch (err) {
                                console.error(err);
                                // showMessage(error, "show");
                            };

                            //#endregion -------------------------------------------------------------------------------------------------------------------

                        };

                        //#endregion -----------------------------------------------------------------------------------------------------------------------

                    };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------------


                    let sheets = fQ + waste; //form quantity plus waste

                    let sheetsAdj4Side;

                    if (singleSided == false) {

                        sheetsAdj4Side = (fQ * 2) + waste; //double the fQ is 2-sided. To be used for the hours calcuation only

                        //sheets variable is left alone since we want the value that is output to the sheet to remain unaffected by sides

                    } else {

                        sheetsAdj4Side = sheets; //don't double fQ since it is single sided. 

                        //waste variable defaults to a 2-sided calculation, and eariler we adjusted it for 1-sided within the handleOneSided function

                    };

                    let hours;
                    let roundHours;

                    for (var z = 0; z < sheetHourArr.length; z++) {
                        if (sheets >= sheetHourData["Sheets (Min)"][z] && sheets <= sheetHourData["Sheets (Max)"][z]) {
                            let divideBy = sheetHourData["Prints Per Hour"][z];
                            hours = (sheetsAdj4Side) / divideBy;

                            //if word before House Stock is "Perfected", then we need to cut the hours in half
                            if (wordBeforeHouseStock == "Perfected") {
                                hours = hours / 2;
                            };

                            roundHours = Math.round((hours + Number.EPSILON) * 100) / 100;
                        }
                    }

                    // let tempSilkObj = { //temp obj that changes each time this runs. However, the values are stored in the silkLayouts array, further down
                    //   layout: row[0],
                    //   form: form,
                    //   formQuantity: formQuantity,
                    //   sheets: sheets,
                    //   hours: roundHours
                    // }

                    // silkLayouts.push(tempSilkObj);


                    //set values for E2Rs, and default values for days and presses in E2Rs
                    let dayVal;
                    let pressVal;

                    if (row[5] == "") {
                        dayVal = "-";
                    } else {
                        dayVal = row[5];
                    };

                    if (worksheet.name == "DIGE2R") {
                        pressVal = "Digital";
                    } else if (row[6] == "") {
                        pressVal = "1";
                    } else {
                        pressVal = row[6];
                    };



                    row[1] = form;
                    row[2] = formQuantity;
                    row[3] = sheets;
                    row[4] = roundHours;
                    row[5] = dayVal;
                    row[6] = pressVal;

                    ////////////////////////////////////////////////////// DO DATA VALIDATION HERE ///////////////////////////////////////////////////////
                    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


                    try {

                        var daysForDataVal = JSON.parse(JSON.stringify(daysOfWeek));

                        var pressesForDataVal = JSON.parse(JSON.stringify(presses));

                        for (let entry in daysForDataVal) {
                            if (daysForDataVal[entry] == " ") {
                                daysForDataVal[entry] = "-";
                            };
                        };

                        let daysString = daysForDataVal.join();

                        for (let press in pressesForDataVal) {
                            if (pressesForDataVal[press] == " ") {
                                pressesForDataVal[press] = "-";
                            };
                        };

                        let pressString = pressesForDataVal.join();

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

                        dayRange.format.horizontalAlignment = "Center"; //makes everything centered (so "-" will be centered upon initally showing up)

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

                        pressRange.format.horizontalAlignment = "Center"; //makes everything centered (so "-" will be centered upon initally showing up)


                    } catch (e) {
                        console.log(e);
                    };





                } else { //only print form for tubes, just like non form lines
                    row[1] = form;
                    row[2] = "-";
                    row[3] = "-";
                    row[4] = "-";
                    row[5] = "-";
                    row[6] = "-";
                };

            } else {
                console.log(`This layout does not contain a form number...\n${row[0]}`);
            };

        } else if (row[0] == "") {
            row[1] = "";
            row[2] = "";
            row[3] = "";
            row[4] = "";
        } else if (form) { //if a form number exists (which almost always will at this point), add it to the second cell in the row

            if (row[0].includes("RUSH")) {
                console.log("This item includes a RUSH tag:");
                console.log(row[0]);
                rushItem = true;
            };

            row[1] = form;
            row[2] = "-";
            row[3] = "-";
            row[4] = "-";
            row[5] = "-";
            row[6] = "-";
        } else {
            console.log(`The following record does not belong to a form: \n${row[0]}`)
        }
        // console.log(`Row[${i}] of tableArray is:\n${row}`);

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
        }

        //conditional formatting...
        conditionalFormatting(worksheet, rowRange, row, certainAddresses);

        i = i + 1; //increments i, which is specifically for the console.log above that is probably commented out at the moment...

    };

    // };

    return tableArray;

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region CLEAR POPULATED INFO ---------------------------------------------------------------------------------------------------------------------

/**
 * Clears the info from the first 2 columns in the Master Table and from the 2nd, 3rd, 4th, & 5th columns in the E2R Tables
 * @param {Array} tableArray The array of the table to clear values from
 * @param {Object} tableRows An object of all the rows in the tab;e
 * @param {String} worksheet The worksheet
 * @returns Array
 */
function clearE2R(tableArray, tableRows, worksheet) {

    let i = 0;

    if (worksheet.name == "Master") {

        for (var row of tableArray) {

            //just in case we need to clear any formatting, but for now we want to leave all formatting as is
            let cheeseRange = tableRows.getItemAt(i).getRange();

            //just the address of the forms column to be used to remove conditional formatting on clear
            let formsAddress = worksheet.getCell(i, formsColumnIndex);

            row[0] = "";
            row[1] = "";

            formsAddress.format.fill.clear();
            formsAddress.format.font.color = "black";
            formsAddress.format.font.bold = false;

            i = i + 1;

        };

    } else {

        for (var row of tableArray) {

            if (tableArray.length < 2 && row[0] == "") {
                return;
            };

            let cheeseRange = tableRows.getItemAt(i).getRange();


            row[1] = "";
            row[2] = "";
            row[3] = "";
            row[4] = "";
            row[5] = "";
            row[6] = "";


            cheeseRange.format.font.bold = false;
            cheeseRange.format.fill.clear();

            i = i + 1;

        };

    };

    return tableArray;

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------

//#region CARRY FORMS ------------------------------------------------------------------------------------------------------------------------------

  /**
   * Compares the product and client code of the current line in the master table to the current line in the current E2R table. If it finds a match, 
   * it pushes the form value to the global formToCarry object. Function also returns true if match is found.
   * @param {Array} row An array of all the values in the row
   * @param {Array} tableArrayCopy A shallow copy of the tableArray to be used for manipulating stuff without changing the original array
   * @param {Array} tableRowItems An array of all the items in the current row
   * @param {Range} tableHeader The header range of the table
   * @param {String} masterProduct The product from the master table to compare to the table row product
   * @param {Number} masterCode The client code from the master table to compare to the table row client code
   * @param {Object} worksheet The worksheet object of the E2R
   */ Boolean
function carryForm(row, tableArrayCopy, tableRowItems, tableHeader, masterUJID, rowIndex, worksheet) {

    formToCarry = "";

    //skips rows that start new form numbers and are not actual products
    if (!row[0].startsWith("Layout") && !row[0].startsWith("Form") && !row[0].startsWith("Tube")) {

        let rowValues = tableRowItems[rowIndex].values; //an array of all the row values at position [a] of the silkE2R table

        rowValues = rowValues[0];

        let spaceSplit = row[0].split(" "); //splits the metrix info apart by spaces and makes it into an array of items

        let rowCode;

        if (row[0] == "") {
            return;
        }

        if (row[0].match(/^\d/)) {

            rowCode = spaceSplit[0];

        } else {

            //takes the string out from the array between the first and second space, then removes the ( character from said text. 
            //Join combines this new blank record with the code to take it from an array to a value
            rowCode = spaceSplit.slice(1, 2)[0].split("(").join("");
            // .replace("(", "")
        };


        //returns (as value, not array becuase of [0] at end) the 2nd (0 indexed) item in the array, and stops before the 3rd
        // let rowProduct = spaceSplit.slice(2, 3)[0]; 

        if (rowCode.includes("_")) {
            let codeSplit = rowCode.split("_");
            rowCode = codeSplit[0];
        };

        let rowInfo = new Object();

        let doesItMatch = false;

        if (masterUJID == rowCode) {
            doesItMatch = true;
        };

        //returns true if the code and product for this row in the master table are in the SilkE2R
        // let doesItMatch = compareMasterToE2R(masterProduct, rowProduct, rowCode, masterCode);

        // console.log("");

        if (doesItMatch) { //if it was in the silkE2R table, then we create an object for the silk row and carry over the form number

            //the following matches the SilkE2R table headers with the data and column index in row [a], assigning each as a property 
            //to each header within the rowInfo object.
            for (var rowName of tableHeader[0]) {
                createRowInfo(tableHeader, rowName, rowValues, tableArrayCopy, rowInfo, rowIndex, worksheet);
            };

            formToCarry = rowInfo["Form"].value; //gets the form number value from the row

            // console.log(formToCarry);

            return true;

        };

    };

};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region CONDITIONAL FORMATTING -------------------------------------------------------------------------------------------------------------------

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
            objOfCells[0].format.font.color = "white";
            objOfCells[0].format.fill.color = "#C65911";
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

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region TRYCATCH ---------------------------------------------------------------------------------------------------------------------------------

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
        showMessage(err, "show");

    };
};

//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#region LEGACY CODE I CAN'T BRING MYSELF TO DELETE YET -------------------------------------------------------------------------------------------


//#region (LEGACY) PRODUCT LIST GLOBAL VARIABLE --------------------------------------------------------------------------------------------------

// let productList = [
//   {name: "100# Gloss Postcard", abbr: "PC", breakout: null},
//   {name: "2 sided box topper", abbr: "2SBT", breakout: null},
//   {name: "2 Sided Flyer", abbr: "2SBT", breakout: null},
//   {name: "80LBFLYER", abbr: "80LBFL", breakout: null},
//   {name: "80LBFLYER", abbr: "80#FL", breakout: null},
//   {name: "Artwork Only", abbr: "CUSTOMTEXT", breakout: null},
//   {name: "Bella Canvas 3001C", abbr: "Bella Canvas 3001C", breakout: null},
//   {name: "Bella Canvas 3001CVC", abbr: "Bella Canvas 3001CVC", breakout: null},
//   {name: "Bella Canvas 3501", abbr: "Bella Canvas 3501", breakout: null},
//   {name: "Bella Canvas 3501CVC", abbr: "Bella Canvas 3501CVC", breakout: null},
//   {name: "BirthdayPC", abbr: "BirthdayPC", breakout: null},
//   {name: "BirthdayPC", abbr: "BIRTH", breakout: null},
//   {name: "Box Topper April", abbr: "MoBT", breakout: null},
//   {name: "Box Topper August", abbr: "MoBT", breakout: null},
//   {name: "Box Topper December", abbr: "MoBT", breakout: null},
//   {name: "Box Topper February", abbr: "MoBT", breakout: null},
//   {name: "Box Topper January", abbr: "MoBT", breakout: null},
//   {name: "Box Topper July", abbr: "MoBT", breakout: null},
//   {name: "Box Topper June", abbr: "MoBT", breakout: null},
//   {name: "Box Topper March", abbr: "MoBT", breakout: null},
//   {name: "Box Topper May", abbr: "MoBT", breakout: null},
//   {name: "Box Topper November", abbr: "MoBT", breakout: null},
//   {name: "Box Topper October", abbr: "MoBT", breakout: null},
//   {name: "Box Topper September", abbr: "MoBT", breakout: null},
//   {name: "Brochure 10.5x17", abbr: "MENU", breakout: null},
//   {name: "Brochure Small", abbr: "80LBFL", breakout: null},
//   {name: "Brochure Small", abbr: "80#FL", breakout: null},
//   {name: "Brochure Small", abbr: "BrochSm", breakout: null},
//   {name: "BrochureXL", abbr: "XL", breakout: null},
//   {name: "BTFP2side", abbr: "BTFP2", breakout: null},
//   {name: "BTFP2side", abbr: "BTFP2side", breakout: null},
//   {name: "BusinessCard", abbr: "CUSTOM110NEENAH", breakout: null},
//   {name: "CALL TRACKING", abbr: "CALLTRACKING", breakout: null},
//   {name: "ColossalPC", abbr: "COLPC", breakout: null},
//   {name: "COUPON BOOK", abbr: "COUP", breakout: null},
//   {name: "Custom 20# Bond", abbr: "CUSTOM20", breakout: null},
//   {name: "Custom Insert", abbr: "CUSTOMTEXT", breakout: null},
//   {name: "Custom100", abbr: "CUSTOM100", breakout: null},
//   {name: "CUSTOM80", abbr: "CUSTOMTEXT", breakout: null},
//   {name: "CustomEnv", abbr: "CUSTOMENV", breakout: null},
//   {name: "CustomEnv", abbr: "ENV", breakout: null},
//   {name: "CustomSilk", abbr: "CUSTOMSILK", breakout: null},
//   {name: "DataProcessing", abbr: "NULL", breakout: null},
//   {name: "DataProcessing", abbr: "DPS", breakout: null},
//   {name: "DDBanner4x8", abbr: "DDBANNER", breakout: null},
//   {name: "DDBizCardsCircle", abbr: "DDBIZ", breakout: null},
//   {name: "DDBrochBiFold", abbr: "DDBROCH", breakout: null},
//   {name: "DDCrewNeck", abbr: "DDCrewNeck", breakout: null},
//   {name: "DDFullZipHoodie", abbr: "DDFullZipHoodie", breakout: null},
//   {name: "DDHoodie", abbr: "DDHoodie", breakout: null},
//   {name: "DDLongSleeve", abbr: "DDLongSleeve", breakout: null},
//   {name: "DDPoloShirt", abbr: "DDPoloShirt", breakout: null},
//   {name: "DDPostcard6x4", abbr: "DDPC", breakout: null},
//   {name: "DDShortSleeve", abbr: "DDShortSleeve", breakout: null},
//   {name: "District DT6100", abbr: "District DT6100", breakout: null},
//   {name: "District DT6104", abbr: "District DT6104", breakout: null},
//   {name: "District DT8102", abbr: "District DT8102", breakout: null},
//   {name: "DOOR HANGER 100LB", abbr: "100#DH", breakout: null},
//   {name: "DOOR HANGER 100LB", abbr: "100LBDH", breakout: null},
//   {name: "DOOR HANGER 100LB", abbr: "DH", breakout: null},
//   {name: "DOOR HANGER 80LB", abbr: "80#DH", breakout: null},
//   {name: "DOOR HANGER 80LB", abbr: "80LBDH", breakout: null},
//   {name: "DOOR HANGER 80LB", abbr: "NULL", breakout: null},
//   {name: "DOOR HANGER 80LB", abbr: "DH", breakout: null},
//   {name: "EDDM 7x7 PC", abbr: "CUSTOM100", breakout: null},
//   {name: "EDDM 7x7 PC", abbr: "EDDMPC", breakout: null},
//   {name: "EDDM Folded Magnet", abbr: "FMAG", breakout: null},
//   {name: "EDDM Folded Magnet", abbr: "FMAGEDDM", breakout: null},
//   {name: "EDDM Mag", abbr: "MAG", breakout: null},
//   {name: "EDDM MENU", abbr: "MENU", breakout: null},
//   {name: "EDDM POSTCARD", abbr: "PC", breakout: null},
//   {name: "EDDM Scratch Off", abbr: "SO", breakout: null},
//   {name: "EDDM XL MENU", abbr: "XL", breakout: null},
//   {name: "EDDM80#FLYER", abbr: "80#FL", breakout: null},
//   {name: "EDDM80#FLYER", abbr: "80LBFL", breakout: null},
//   {name: "EDDM80#FLYER", abbr: "EDDM80FLY", breakout: null},
//   {name: "EDDMBroch", abbr: "MENU", breakout: null},
//   {name: "EDDMBrochXL", abbr: "XL", breakout: null},
//   {name: "EDDMColossal", abbr: "COLPC", breakout: null},
//   {name: "EDDMJumboPC", abbr: "JUMBO", breakout: null},
//   {name: "EDDMJumboSO", abbr: "JUMBO", breakout: null},
//   {name: "EDDMPeelAGift", abbr: "PPC", breakout: null},
//   {name: "EDDMPizzaPeelCard", abbr: "PPC", breakout: null},
//   {name: "Env #10 8.5x11 S1", abbr: "LET", breakout: null},
//   {name: "Env #10 8.5x11 S1", abbr: "LET_1S", breakout: null},
//   {name: "Env #10 8.5x11 S2", abbr: "LET", breakout: null},
//   {name: "Env #10 8.5x11 S2", abbr: "LET_2S", breakout: null},
//   {name: "Env #10 8.5x11 V1", abbr: "LET", breakout: null},
//   {name: "Env #10 8.5x11 V1", abbr: "LET_1S", breakout: null},
//   {name: "Env #10 8.5x11 V2", abbr: "LET", breakout: null},
//   {name: "Env #10 8.5x11 V2", abbr: "LET_2S", breakout: null},
//   {name: "Env #10 8.5x14 S1", abbr: "LEG", breakout: null},
//   {name: "Env #10 8.5x14 S1", abbr: "LEG_1S", breakout: null},
//   {name: "Env #10 8.5x14 S2", abbr: "LEG", breakout: null},
//   {name: "Env #10 8.5x14 S2", abbr: "LEG_2S", breakout: null},
//   {name: "Env #10 8.5x14 V1", abbr: "LEG", breakout: null},
//   {name: "Env #10 8.5x14 V1", abbr: "LEG_1S", breakout: null},
//   {name: "Env #10 8.5x14 V2", abbr: "LEG", breakout: null},
//   {name: "Env #10 8.5x14 V2", abbr: "LEG_2S", breakout: null},
//   {name: "Expedite", abbr: "Expedite", breakout: null},
//   {name: "FakeT", abbr: "FakeT", breakout: null},
//   {name: "Flyer 8.5X11", abbr: "8_5x11FL", breakout: null},
//   {name: "Folded Magnet", abbr: "FMAG", breakout: null},
//   {name: "Gildan 2000B", abbr: "Gildan 2000B", breakout: null},
//   {name: "Gildan 6400L", abbr: "Gildan 6400L", breakout: null},
//   {name: "Gildan 8800", abbr: "Gildan 8800", breakout: null},
//   {name: "Gildan G5000", abbr: "Gildan G5000", breakout: null},
//   {name: "Gildan G540", abbr: "Gildan G540", breakout: null},
//   {name: "Gildan G640", abbr: "Gildan G640", breakout: null},
//   {name: "Guide", abbr: "Guide", breakout: null},
//   {name: "Hanes F260", abbr: "Hanes F260", breakout: null},
//   {name: "Hanes T-Shirt", abbr: "Hanes T-Shirt", breakout: null},
//   {name: "Jumbo Scratch", abbr: "JUMBOSO", breakout: null},
//   {name: "JUMBOPC", abbr: "JUMBO", breakout: null},
//   {name: "Long Postcard", abbr: "LPO", breakout: null},
//   {name: "MAGNET", abbr: "MAG", breakout: null},
//   {name: "Mail List Costs", abbr: "Mail List Costs", breakout: null},
//   {name: "MENU", abbr: "MENU", breakout: null},
//   {name: "Menu- Flat", abbr: "MENU", breakout: null},
//   {name: "Menu Small", abbr: "80LBFL", breakout: null},
//   {name: "Menu Small", abbr: "80#FL", breakout: null},
//   {name: "Menu Small", abbr: "MenuSm", breakout: null},
//   {name: "MENU XXL", abbr: "XXL", breakout: null},
//   {name: "MenuXL", abbr: "XL", breakout: null},
//   {name: "MenuXL-Flat", abbr: "XL", breakout: null},
//   {name: "MGP-LONG POSTCARD", abbr: "LPO", breakout: null},
//   {name: "MGP-MAGNET", abbr: "MAG", breakout: null},
//   {name: "MGP-PLASTIC MED 20m", abbr: "MPL", breakout: null},
//   {name: "MGP-POSTCARD", abbr: "PC", breakout: null},
//   {name: "MGP-SCRATCHOFF", abbr: "SO", breakout: null},
//   {name: "MockPostcard", abbr: "MockPostcard", breakout: null},
//   {name: "MockPostcard", abbr: "NULL", breakout: null},
//   {name: "MPBTAC", abbr: "MPBTAC", breakout: null},
//   {name: "MPBTNO", abbr: "MPBTNO", breakout: null},
//   {name: "MPC", abbr: "MPC", breakout: null},
//   {name: "MPCC", abbr: "MPCC", breakout: null},
//   {name: "MPCMS", abbr: "MPCMS", breakout: null},
//   {name: "MPCVDLAM", abbr: "CVDLAM", breakout: null},
//   {name: "MPEXTWC2436AC", abbr: "MPWC2436AC", breakout: null},
//   {name: "MPEXTWC2436NO", abbr: "MPWC2436NO", breakout: null},
//   {name: "MPEXTWC3040AC", abbr: "MPWC3040AC", breakout: null},
//   {name: "MPEXTWC3040NO", abbr: "MPWC3040NO", breakout: null},
//   {name: "MPINTWC2436AC", abbr: "MPINTWC2436AC", breakout: null},
//   {name: "MPINTWC2436NO", abbr: "MPINTWC2436NO", breakout: null},
//   {name: "MPINTWC3040AC", abbr: "MPINTWC3040AC", breakout: null},
//   {name: "MPINTWC3040NO", abbr: "MPINTWC3040NO", breakout: null},
//   {name: "MPIS", abbr: "MPIS", breakout: null},
//   {name: "MPND", abbr: "MPND", breakout: null},
//   {name: "MPNutGuide", abbr: "MPNutGuide", breakout: null},
//   {name: "MPP", abbr: "MPP", breakout: null},
//   {name: "MPPICMENU", abbr: "PICMENU", breakout: null},
//   {name: "MPPlastic PC Med - S", abbr: "MPPlastic PC Med - S", breakout: null},
//   {name: "MPPlastic PC Med - S", abbr: "NULL", breakout: null},
//   {name: "MPPO2254", abbr: "MPPO2254", breakout: null},
//   {name: "MPPO2436", abbr: "MPPO2436", breakout: null},
//   {name: "MPPO3040", abbr: "MPPO3040", breakout: null},
//   {name: "MPPODS2436", abbr: "MPPODS2436", breakout: null},
//   {name: "MPPS", abbr: "MPPS", breakout: null},
//   {name: "MPSCFS", abbr: "MPSCFS", breakout: null},
//   {name: "NEW MOVERS PLASTIC", abbr: "NEW MOVERS PLASTIC", breakout: null},
//   {name: "NEW MOVERS PLASTIC", abbr: "SPL_NM", breakout: null},
//   {name: "NEW MOVERS POSTCARD", abbr: "NEW MOVERS POSTCARD", breakout: null},
//   {name: "NEW MOVERS POSTCARD", abbr: "NM", breakout: null},
//   {name: "Next Level 3600", abbr: "Next Level 3600", breakout: null},
//   {name: "Next Level N6210", abbr: "Next Level N6210", breakout: null},
//   {name: "NON PROFIT MAILER", abbr: "NPM", breakout: null},
//   {name: "NonProfit100#", abbr: "CUSTOM100", breakout: null},
//   {name: "NonProfit80#", abbr: "CUSTOMTEXT", breakout: null},
//   {name: "ONLINE ORDERING", abbr: "ONLINE ORDERING", breakout: null},
//   {name: "ONLINE ORDERING", abbr: "NULL", breakout: null},
//   {name: "Payment Plan", abbr: "Payment Plan", breakout: null},
//   {name: "Peel A Gift", abbr: "PPC", breakout: null},
//   {name: "PizzaPeelCard", abbr: "PPC", breakout: null},
//   {name: "Plastic PC Lg - S", abbr: "LPL", breakout: null},
//   {name: "Plastic PC Lg - V", abbr: "LPL", breakout: null},
//   {name: "Plastic PC Med - S", abbr: "MPL", breakout: null},
//   {name: "Plastic PC Med - V", abbr: "MPL", breakout: null},
//   {name: "Plastic PC Sm - S", abbr: "SPL", breakout: null},
//   {name: "Plastic PC Sm - V", abbr: "SPL", breakout: null},
//   {name: "Port Authority K8000", abbr: "Port Authority K8000", breakout: null},
//   {name: "Port Company PC850H", abbr: "Port Company PC850H", breakout: null},
//   {name: "POSTCARD", abbr: "PC", breakout: null},
//   {name: "Postcard Magnet", abbr: "MAG", breakout: null},
//   {name: "PosterHangBar", abbr: "PosterHangBar", breakout: null},
//   {name: "REPRINT100#", abbr: "REPRINT100", breakout: null},
//   {name: "REPRINT80#", abbr: "REPRINT80", breakout: null},
//   {name: "SCRATCHOFF", abbr: "SO", breakout: null},
//   {name: "Unique Codes", abbr: "Unique Codes", breakout: null},
//   {name: "WC2030", abbr: "WC2030", breakout: null},
//   {name: "WC2430", abbr: "WC2430", breakout: null},
//   {name: "WC2436", abbr: "WC2436", breakout: null},
//   {name: "WC3040", abbr: "WC3040", breakout: null},
//   {name: "WCCust", abbr: "WCCUSTOM", breakout: null},
//   {name: "WCPoster2Side2430", abbr: "2xWC2430", breakout: null},
//   {name: "WCPoster2Side2436", abbr: "2xWC2436", breakout: null},
//   {name: "x.100#custom", abbr: "CUSTOM100", breakout: null},
//   {name: "x.80#custom", abbr: "CUSTOM80", breakout: null},
//   {name: "x.Flyer.10.5x17", abbr: "MENU", breakout: null},
//   {name: "x.Flyer.8.5x10.5", abbr: "80#FL", breakout: null},
//   {name: "x.JUMBOPC", abbr: "JUMBO", breakout: null},
//   {name: "x.Magnet", abbr: "MAG", breakout: null},
//   {name: "x.Menu.10.5x17", abbr: "MENU", breakout: null},
//   {name: "x.Menu.8.5x10.5", abbr: "MenuSm", breakout: null},
//   {name: "x.Postcard.5.5x10.5", abbr: "PC", breakout: null},
//   {name: "x.Postcard.5.5x8.5", abbr: "CUSTOM100", breakout: null},
//   {name: "x.Postcard.8.5x10.5", abbr: "JUMBO", breakout: null},
//   {name: "x.Reprint100#", abbr: "x.Reprint100", breakout: null},
//   {name: "x.Reprint100#", abbr: "NULL", breakout: null},
//   {name: "x.Reprint80#", abbr: "x.Reprint80", breakout: null},
//   {name: "x.Reprint80#", abbr: "NULL", breakout: null},
//   {name: "x.ScratchOff", abbr: "SO"}
// ];

//#endregion -------------------------------------------------------------------------------------------------------------------------------------


//#region (LEGACY) FINDS ALL ABBREVIATIONS FOR THE GIVEN PRODUCT IN THE PRODUCT LIST -------------------------------------------------------------

// function productAbbr (product) {

//   const result = productList.filter((item) => {
//       if (item.name === product) {
//         return item;
//       }
//   });

//   // console.log(result);

//   if (result.length > 0) {
//       // Handle multiple results
//       return result
//   } else {
//       // Handle no results
//       const err = new Error("No products found.")
//       return err
//   }

// }

//#endregion -------------------------------------------------------------------------------------------------------------------------------------

//#region CAMEL-CASE A STRING FROM TABLE DATA AND TRY TO CREATE DYNAMIC VARIABLES (CAMEL-CASE WORKS, DYNAMIC VARIABLES DO NOT) -------------------

// for (var lineItem of linesData) {

//   let lineItemSplit = lineItem.split(" ");

//   let firstWord = lineItemSplit[0].toLowerCase();

//   if (firstWord.includes("-")) {
//     firstWord = firstWord.replace("-", "");
//   };

//   lineItemSplit.shift();

//   let tempString = firstWord;

//   for (var word of lineItemSplit) {

//     if (word.includes("-")) {
//       word = word.replace("-", "");
//     };

//     let cheese = word.charAt(0).toUpperCase() + word.substr(1).toLowerCase(); //+ lineItemSplit[word].slice(1);

//     tempString = tempString + cheese;

//     // console.log(tempString);

//   }

//   lineArray.push(tempString);

//   console.log(lineArray);

//   // if (masterType == lineItem) {
//   //   array.push(masterRow)
//   // }
// };


// arr1 = [[postcardLine4Podium], [scratchoffLine]]

// for (var m = 0; m < lineArray.length; m++) {
//   eval('let ' + lineArray[m] + '= ' + m + ';');
// };

// console.log(postcardLine4Podium);

//#endregion -------------------------------------------------------------------------------------------------------------------------------------


//#endregion ---------------------------------------------------------------------------------------------------------------------------------------


//#endregion -----------------------------------------------------------------------------------------------------------------------------------------


/*scrollErr.scrollTop = scrollHeight;*/