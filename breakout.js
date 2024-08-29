import globalVar from "./globalVar.js";
import { deactivateEvents, activateEvents, createRowInfo, loadError } from "./universalFunctions.js";

var showTitle = false; 

//====================================================================================================================================================
    //#region BREAKOUT FUNCTION ----------------------------------------------------------------------------------------------------------------------
            
        /**
         * Creates all the breakout sheets and tables, as well as formatting the cells inside each breakout table
         */
        async function breakout() {
            

            deactivateEvents(); //turn off events
            //========================================================================================================================================
                //#region SHOW LOADING SCREEN --------------------------------------------------------------------------------------------------------

                    $("#loading-background").css("display", "flex");

                    if ($("#breakout").text() == "Delete Breakout Sheets") {
                        $("#loading-head").text("Removing Breakout Tables");
                    } else {
                        $("#loading-head").text("Creating Breakout Tables");
                    };

                //#endregion -------------------------------------------------------------------------------------------------------------------------
            //========================================================================================================================================

            await Excel.run(async (context) => {

                //====================================================================================================================================
                    //#region ASSIGN SHEET VARIABLES -------------------------------------------------------------------------------------------------

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

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                await context.sync();

                //====================================================================================================================================
                    //#region LOAD SHEET VARIABLES ---------------------------------------------------------------------------------------------------
                
                        let linesArr = linesBodyRange.values;
                        let linesHeader = linesHeaderRange.values;
                        let masterArr = masterBodyRange.values;
                        let masterHeader = masterHeaderRange.values;
                        let masterRowItems = masterTableRows.items;
                        let masterArrCopy = JSON.parse(JSON.stringify(masterBodyRange.values))

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region REMOVE EXISTING BREAKOUT SHEETS (IF ANY) -------------------------------------------------------------------------------

                        console.log("Doing it...")
                        const doIt = await removeBreakoutSheets(linesArr, allSheets);
                        console.log(doIt);

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                await context.sync();

                //====================================================================================================================================
                    //#region EXIT HERE IF BUTTON IS "DELETE BREAKOUT SHEETS" ------------------------------------------------------------------------

                        if ($("#breakout").text() == "Delete Breakout Sheets") {
                            $("#breakout").text("Breakout");

                            console.log("exiting...");

                            return;
                        };
                    
                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //?===================================================================================================================================
                    //#region SUDO CODE FOR FIXING ROW COLOR SORTING SLOWDOWN ON SHAREPOINT MAYBE?? -------------------------------------------------

                        //////////////////////////////////////////////// STOPPED HERE TRYING TO FIGURE OUT ROW COLOR STORING /////////////////////////
                        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                        // for (let row of masterRowItems) {
                        //   let rowRange = row.getRange();

                        //   let rangeFill = rowRange.format.fill;
                        //   rangeFill.load(["color"])

                        //   await context.sync();
                        //   console.log("range.format.fill.color", rangeFill.color)
                        // };



                        // var allTablesPrior = context.workbook.tables;
                        // allTablesPrior.load("items/name");

                        //await context.sync();

                        // allTablesPrior.items.forEach((tablePrior) => {
                        //   console.log(tablePrior.name);
                        // });


                        // allTablesPrior.items

                        // let tempLinesObj = {}; //temporary object for storing values to push into the data object (clears each time)

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //?===================================================================================================================================

                //====================================================================================================================================
                    //#region CREATING THE BREAKOUTS -------------------------------------------------------------------------------------------------

                        //============================================================================================================================
                            //#region CREATING OBJECTS AND VARIABLES ---------------------------------------------------------------------------------

                                //====================================================================================================================
                                    //#region EMPTY VARIABLES ----------------------------------------------------------------------------------------

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
                                        let overwriteMissing = false;
                                        let formsObj = {};
                                        let tableName;

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================
    
                                //====================================================================================================================
                                    //#region MAKE ARRAY OF LINE TYPES -------------------------------------------------------------------------------

                                        globalVar.linesData = [];

                                        for (var p = 0; p < linesArr.length; p++) {

                                            let currentLines = linesArr[p][0];

                                            globalVar.linesData.push(currentLines);

                                        };
                                    
                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                //====================================================================================================================
                                    //#region CREATE FILTERED TYPE OBJECT WITH SUB OBJECTS FROM THE LINE TYPES ---------------------------------------

                                        let filteredData = {};

                                        //for each line type in the globalVar.linesData array, make a new sub-object that we can assign properties to later 
                                        //inside filteredData
                                        globalVar.linesData.forEach((line) => {

                                            filteredData[line] = [];

                                        });

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                //====================================================================================================================
                                    //#region CREATE NORMAL BREAKOUTS FORMATTING OBJECT FROM LINE TYPES ----------------------------------------------

                                        //this is similar to the Filtered Data object, but later on we will store formatting values in 
                                        //this object of arrays
                                        for (let zeLine of globalVar.linesData) {
                                            //this is a global variable
                                            globalVar.normalBreakoutsFormatting[zeLine] = [];
                };

                console.log(globalVar.normalBreakoutsFormatting);

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                //====================================================================================================================
                                    //#region CREATE MISSING DATA OBJECT WITH SUB OBJECTS FROM THE HIDDEN LINES DATA ---------------------------------

                                        let missingData = {};

                                        //hiddenglobalVar.linesData is a globally defined array containing all the items that should be marked as missing 
                                        //if they show up
                                        globalVar.hiddenLinesData.forEach((item) => {

                                            missingData[item] = [];

                                        });

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //================================= ===================================================================================

                            //#endregion -------------------------------------------------------------------------------------------------------------
                        //============================================================================================================================
               
                        //============================================================================================================================
                            //#region FOR EACH ROW IN MASTER, MOVE VALUES TO PROPER BREAKOUT ARRAY(S) ------------------------------------------------
                                globalVar.masterCellData = {};

                                for (let row of masterArr){

                                    let rowUjid= row[masterHeader[0].indexOf("UJID")];

                                    let rowArr = [];

                                    //the following matches the master table headers with the data and column index in row [z], assigning each
                                    //as a property to each header within the masterRowInfo object.

                                    for (let name of masterHeader[0]) {
                                        
                                        // createRowInfo(masterHeader, name, masterRow, masterArrCopy, masterRowInfo, z, masterSheet);
                                        const cellData = createRowInfo(masterHeader, name, row, null, null, z, masterSheet, true);

                                        rowArr.push(cellData);
                                        
                                    };

                                    globalVar.masterCellData[rowUjid] = rowArr;                                    

                                    z++;

                                };

                                        
                                await context.sync();

                                for (let masterRow of masterArr) {

                                    //================================================================================================================
                                        //#region CREATE ROW INFO OBJECT FOR CURRENT ROW IN MASTER ---------------------------------------------------

                                            //the following matches the master table headers with the data and column index in row [z], assigning each
                                            //as a property to each header within the masterRowInfo object.

                                            for (let name of masterHeader[0]) {
                                                
                                                createRowInfo(masterHeader, name, masterRow, masterArrCopy, masterRowInfo, z, masterSheet);
                                            };
                                            

                                            z++;

                                        //#endregion -------------------------------------------------------------------------------------------------
                                    //================================================================================================================

                                    await context.sync();

                                    //================================================================================================================
                                        //#region CREATE VARIABLES OF THE IMPORTANT COLUMN VALUES ----------------------------------------------------

                                            let masterForms = masterRowInfo["Forms"].value;
                                            let masterType = masterRowInfo["Type"].value;
                                            masterUJIDColumnIndex = masterRowInfo["UJID"].columnIndex;

                                            let masterExtras = masterRowInfo["Extras"].value;
                                            let masterCutoff = masterRowInfo["Cutoff"].value;
                                            let masterNotesColumnIndex = masterRowInfo["Notes"].columnIndex;

                                            let masterNotes = masterRowInfo["Notes"].value;
                                            let masterVersionNo = masterRowInfo["Version No"].value;

                                        //#endregion -------------------------------------------------------------------------------------------------
                                    //================================================================================================================

                                    //================================================================================================================
                                        // #region COMBINE NOTES -------------------------------------------------------------------------------------

                                            /* 
                                                Explanation: This makes and array and joins the array based on how many items get added. This keeps 
                                                the notes from having random " - " bits stuck on the end.
                                            */
                                            let notesArr = new Array(); 

                                            if (!(/^(?:\s|,|\s*,\s*)$/.test(masterNotes))) {
                                                notesArr.push(masterNotes);
                                            }
 
                                            if (!masterVersionNo.replace(/ /g, "") == "") {
                                                notesArr.push(masterVersionNo);
                                            }
                                            if (notesArr.length == 1 ){
                                                // If no version was pushed and no options were pushed...
                                                masterRow[masterNotesColumnIndex] = masterNotes;
                                            } else {
                                            //    masterRow[masterNotesColumnIndex] = notesArr.map((n, i)=> `${arrayKeys[i]}: ${n}`).join(" - ");
                                                masterRow[masterNotesColumnIndex] = notesArr.join(" - ")
                                            }
                                            

                                        // #endregion ------------------------------------------------------------------------------------------------
                                    //================================================================================================================

                                    //================================================================================================================
                                        //#region LOG ALL FORMATTING ATTRIBUTES OF THE ROW TO FORMS OBJ ----------------------------------------------

                                            for (let head of masterHeader[0]) {
                                                let formsAddress = masterRowInfo[head].cellProps.value[0][0].address
                                                let formsFill = masterRowInfo[head].cellProps.value[0][0].format.fill.color
                                                let formsFontColor = masterRowInfo[head].cellProps.value[0][0].format.font.color
                                                let formsFontBold = masterRowInfo[head].cellProps.value[0][0].format.font.bold
                                                let formsFontItalic = masterRowInfo[head].cellProps.value[0][0].format.font.italic
                                                let formsValue = masterRowInfo[head].value

                                                formsObj[head] = {
                                                    formsAddress,
                                                    formsFill,
                                                    formsFontColor,
                                                    formsFontBold,
                                                    formsFontItalic,
                                                    formsValue
                                                };
                                            };

                                            let formsObjCopy = JSON.parse(JSON.stringify(formsObj));

                                        //#endregion -------------------------------------------------------------------------------------------------
                                    //================================================================================================================

                                    //================================================================================================================
                                        //#region PUSH ROW INFO TO PROPER BREAKOUT ARRAY -------------------------------------------------------------

                                            try {

                                                //====================================================================================================
                                                    //#region PUSH TO DIGITAL BREAKOUT ---------------------------------------------------------------

                                                        /* if form number is between 200 & 699, then it is a digital form and should be duplicated 
                                                        into its own digital breakout (without taking it away from any other breakout it might fall 
                                                        into, which is why it is outside of the if/else area below) */
                                                        if ((masterForms == "DIGITAL") || masterForms >= 201 && masterForms <= 700) {
                                                            // masterForms > 201 && masterForms < 700
                                                            digitalBreakout.push(masterRow);
                                                            digitalFormatting.push(formsObjCopy);

                                                            missingData["DIGITAL"].push(masterRow);

                                                        };

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region PUSH TO SHIPPING BREAKOUT --------------------------------------------------------------

                                                        if (masterType == "Shipping" && (masterForms !== "IGNORE")) { 
                                                            //if type is shipping, push to just shipping array

                                                            shipping.push(masterRow);
                                                            shippingFormatting.push(formsObjCopy);

                                                            missingData["Shipping"].push(masterRow);

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region PUSH TO IGNORE BREAKOUT (IF TYPE IS IGNORE) --------------------------------------------

                                                        } else if (masterType == "Ignore" || masterForms == "IGNORE") {
                                                            ignore.push(masterRow);
                                                            ignoreFormatting.push(formsObjCopy);

                                                            missingData["IGNORE"].push(masterRow);

                                                            //========================================================================================
                                                                //#region LEGACY: PUSH EXTRAS & CUTOFFS TO SHIPPING BREAKOUT -------------------------

                                                                    //if includes extras or cutoffs, push to both type array and shipping array
                                                                    // } else if (masterExtras > 0 && masterExtras !== "" || 
                                                                    //  masterCutoff > 0 && masterCutoff !== "") { 

                                                                    //   if (masterForms == "MISSING") {

                                                                    //     missing.push(masterRow);
                                                                    //     missingFormatting.push(formsObj);

                                                                    //     missingData["MISSING"].push(masterRow);

                                                                    //     shipping.push(masterRow);
                                                                    //     shippingFormatting.push(formsObj);

                                                                    //     missingData["Shipping"].push(masterRow);
                                                                    //     filteredData[masterType].push(masterRow);
                                                                    //     globalVar.normalBreakoutsFormatting[masterType].push(formsObj);


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
                                                                    //     globalVar.normalBreakoutsFormatting[masterType].push(formsObj);

                                                                    //   };

                                                                //#endregion -------------------------------------------------------------------------
                                                            //========================================================================================

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region PUSH TO MISSING BREAKOUT ---------------------------------------------------------------

                                                        //also pushes these items to the normal type breakouts as well
                                                        } else if (masterForms == "MISSING") { 

                                                            missing.push(masterRow);
                                                            missingFormatting.push(formsObjCopy);

                                                            missingData["MISSING"].push(masterRow);
                                                            filteredData[masterType].push(masterRow);
                                                            globalVar.normalBreakoutsFormatting[masterType].push(formsObjCopy);

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region PUSH TO PRINTED BREAKOUT (IF FORM IS IGNORE) --------------------------------------------
                                                        } else if (masterForms == "PRINTED") {

                                                            printed.push(masterRow);
                                                            printedFormatting.push(formsObjCopy);

                                                            missingData["PRINTED"].push(masterRow);

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region PUSH TO NORMAL BREAKOUTS ---------------------------------------------------------------

                                                        } else if (masterForms == "APPAREL") {

                                                            filteredData['Apparel'].push(masterRow);
                                                            globalVar.normalBreakoutsFormatting['Apparel'].push(formsObjCopy);

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region PUSH TO NORMAL BREAKOUTS ---------------------------------------------------------------

                                                        } else { 
                                                            //if neither shipping type nor has extras or cutoffs, treat normally
                                                            filteredData[masterType].push(masterRow);
                                                            globalVar.normalBreakoutsFormatting[masterType].push(formsObjCopy);
                                                            // normalFormatting.push(formsObjCopy);
                                                        };

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                            } catch (e) {
                                                console.log(e);
                                                console.log("Could not find master type: ", masterType);
                                                empty.push(masterRow);
                                                missingData["Empty"].push(masterRow);
                                            };

                                            //need to alert user of any items that may appear in the empty array. This means that the type column is 
                                            //either empty or has an invlaid record

                                        //#endregion -------------------------------------------------------------------------------------------------
                                    //================================================================================================================

                                };

                            //#endregion -------------------------------------------------------------------------------------------------------------
                        //============================================================================================================================

                        //============================================================================================================================
                            //#region SHOW EMPTY UJID WARNING IF MISSING ONE OR MORE AND DON'T CREATE BREAKOUTS --------------------------------------

                                let emptyUJIDs = [];

                                if (empty.length > 0) {
                                    //show a warning box and do not generate breakout sheets

                                    if (empty.length > 1) {

                                        for (let u = 0; u < empty.length; u++) {

                                            emptyUJIDs.push(empty[u][masterUJIDColumnIndex]);

                                        }

                                    } else {
                                        emptyUJIDs.push(empty[0][masterUJIDColumnIndex]);
                                    };

                                    $("#empty-ujid").empty(); //removes all previous UJIDs from HTML element

                                    // document.getElementById("empty-ujid").innerHTML = emptyUJIDs;
                                    emptyUJIDs.forEach((emp) => {
                                        $("#empty-ujid").append(`<li>${emp}</li>`)
                                    })

                                    emptyWarning();

                                    return;
                                };

                                if (document.querySelector("#breakout").textContent == "Breakout" || showTitle == true){
 
                                    // ===============================================================================================================
                                        // #region BREAKOUT HEADERS ----------------------------------------------------------------------------------
                                            $("#loading-background").css("display", "none");
                                            $("#format-background").css("display", "flex");
                                            // Wait until they press "Go".
                                            await new Promise(resolve => {
                                                document.querySelector("#format-btn").addEventListener('click', resolve, { once: true });
                                            });
                                            globalVar.headerPrefix = $("#preview").val(); // Var to use this outside of this function.
                        
                                            $("#format-background").css("display", "none");
                                            $("#loading-background").css("display", "flex");
                                        // #endregion ------------------------------------------------------------------------------------------------
                                    // ===============================================================================================================

                                    }

                            //#endregion -------------------------------------------------------------------------------------------------------------
                        //============================================================================================================================

                        //============================================================================================================================
                            //#region CREATE & FORMAT THE CUSTOM NON-TYPE BREAKOUT TABLES (EXCEPT SHIPPING) ------------------------------------------

                                //====================================================================================================================
                                    //#region CREATE SHEETS AND TABLES IN EXCEL FOR CUSTOM BREAKOUTS -------------------------------------------------
                                            
                                        let columnsToHide = [];

                                        let missingTable = addSheetAndTable("MISSING", allSheets, missingData, masterHeader, "Missing", masterRowInfo);

                                        let printedTable = addSheetAndTable("PRINTED", allSheets, missingData, masterHeader, "Printed", masterRowInfo);

                                        let ignoreTable = addSheetAndTable("IGNORE", allSheets, missingData, masterHeader, "Ignore", masterRowInfo);

                                        let digitalTable = addSheetAndTable("DIGITAL", allSheets, missingData, masterHeader, "Digital", masterRowInfo);


                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                await context.sync();

                                //====================================================================================================================
                                    //#region GET NUMBER OF ROWS IN THE CUSTOM BREAKOUT TABLES -------------------------------------------------------

                                        let missingRowCount = missingTable.table.rows.getCount();
                                        let printedRowCount = printedTable.table.rows.getCount();
                                        let ignoreRowCount = ignoreTable.table.rows.getCount();
                                        let digitalRowCount = digitalTable.table.rows.getCount();

                                        // let shippingRowCount = shippingTable.table.rows.getCount();

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                let missingSortedValues = missingTable.table.getDataBodyRange().load("values");
                                let printedSortedValues = printedTable.table.getDataBodyRange().load("values");
                                let ignoreSortedValues = ignoreTable.table.getDataBodyRange().load("values");
                                let digitalSortedValues = digitalTable.table.getDataBodyRange().load("values");


                                await context.sync();

                                let missingSorted = missingSortedValues.values;
                                let printedSorted = printedSortedValues.values;
                                let ignoreSorted = ignoreSortedValues.values;
                                let digitalSorted = digitalSortedValues.values;


                                let missingSortedTableData = missingSorted.map(row => {
                                    let obj = {};
                                    row.forEach((cell, index) => {
                                        obj[masterHeader[0][index]] = cell;
                                    });
                                    return obj;
                                });

                                let printedSortedTableData = printedSorted.map(row => {
                                    let obj = {};
                                    row.forEach((cell, index) => {
                                        obj[masterHeader[0][index]] = cell;
                                    });
                                    return obj;
                                });

                                let ignoreSortedTableData = ignoreSorted.map(row => {
                                    let obj = {};
                                    row.forEach((cell, index) => {
                                        obj[masterHeader[0][index]] = cell;
                                    });
                                    return obj;
                                });

                                let digitalSortedTableData = digitalSorted.map(row => {
                                    let obj = {};
                                    row.forEach((cell, index) => {
                                        obj[masterHeader[0][index]] = cell;
                                    });
                                    return obj;
                                });

                                //====================================================================================================================
                                    //#region FORMAT CELLS IN THE CUSTOM BREAKOUT TABLES -------------------------------------------------------------


                                        let theMissingFormat = styleCells(
                                            missingTable.table, missingFormatting, missingRowCount.value, masterHeader, "MISSING", missingSortedTableData
                                        );
                                        let thePrintedFormat = styleCells(
                                            printedTable.table, printedFormatting, printedRowCount.value, masterHeader, "PRINTED", printedSortedTableData
                                        );
                                        let theIgnoreFormat = styleCells(
                                            ignoreTable.table, ignoreFormatting, ignoreRowCount.value, masterHeader, "IGNORE", ignoreSortedTableData
                                        );
                                        let theDigitalFormat = styleCells(
                                            digitalTable.table, digitalFormatting, digitalRowCount.value, masterHeader, "DIGITAL", digitalSortedTableData
                                        );

                                        context.workbook.tables.getItem("Digital").sort.apply([
                                            { key: 0, ascending: true } // '0' is the column number for 'Form'
                                        ]);
                                        
                                        /* let theShippingFormat = styleCells(
                                            shippingTable.table, shippingFormatting, shippingRowCount.value, masterHeader, "Shipping"
                                        );*/

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                await context.sync();

                                //====================================================================================================================
                                    //#region HIDE CERTAIN COLUMNS IN THE CUSTOM BREAKOUTS -----------------------------------------------------------

                                        hideColumns(missingTable.table, columnsToHide);
                                        hideColumns(printedTable.table, columnsToHide);
                                        hideColumns(ignoreTable.table, columnsToHide);
                                        hideColumns(digitalTable.table, columnsToHide);
                                        // hideColumns(shippingTable.table, columnsToHide);

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                //====================================================================================================================
                                    //#region SET PRINT SETTINGS FOR THE CUSTOM BREAKOUTS ------------------------------------------------------------

                                        printSettings(missingTable.sheet);
                                        printSettings(printedTable.sheet);
                                        printSettings(ignoreTable.sheet);
                                        printSettings(digitalTable.sheet);
                                        // printSettings(shippingTable.sheet);

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                            //#endregion -------------------------------------------------------------------------------------------------------------
                        //============================================================================================================================

                       

                        await context.sync();


                        //============================================================================================================================
                            //#region CREATE & FORMAT THE NORMAL TYPE BREAKOUTS ----------------------------------------------------------------------

                                for (let line of globalVar.linesData) {

                                    //================================================================================================================
                                        //#region CREATE THE TABLE NAMES FOR THE TYPE BREAKOUTS FROM LINES DATA --------------------------------------

                                            //* Excel's table names need to be formatted in a certain way. They cannot contain spaces or special 
                                            //* characters. It is also recommended to proper-case each word for readability. The problem here is that
                                            //* we are referencing a dynamic array of values based on the table types defined in the Validation sheet. 
                                            //* So we need to be able to dynamically load in these values and create proper table names from them so 
                                            //* we can reference them later on in the code. The values we are loading in however typically are not
                                            //* formatted for table names. For example, there may be a type called "Scratch-Off Postcards". We would 
                                            //* need this to read as "ScratchoffPostcards" for the table name. That is what we are doing in this 
                                            //* region of code. 

                                            //split each lineData line at the space, making an array of each word in the line
                                            let lineItemSplit = line.split(" ");

                                            //force the first word to be proper case
                                            let firstWord = lineItemSplit[0].charAt(0).toUpperCase() + lineItemSplit[0].substr(1).toLowerCase();

                                            //remove dashes from first word if it has any
                                            if (firstWord.includes("-")) {
                                                firstWord = firstWord.replace("-", "");
                                            };

                                            //removes first word from lineItemSplit
                                            lineItemSplit.shift();

                                            //set tableName to just the properly formatted first word for now
                                            tableName = firstWord;

                                            //for the remaining words in lineItemSplit (if any), do the following
                                            for (var word of lineItemSplit) {

                                                //remove any dashes from the word
                                                if (word.includes("-")) {
                                                    word = word.replace("-", "");
                                                };

                                                //proper case this word
                                                let nextWord = word.charAt(0).toUpperCase() + word.substr(1).toLowerCase(); 
                                                //+ lineItemSplit[word].slice(1);

                                                //add this word after the last one (no spaces)
                                                tableName = tableName + nextWord;

                                            };

                                            //push fully formatted table name to an array of breakout table names
                                            globalVar.listOfBreakoutTables.push(tableName);

                                        //#endregion -------------------------------------------------------------------------------------------------
                                    //================================================================================================================

                                    //================================================================================================================
                                        //#region CREATE AND FORMAT THE TYPE BREAKOUT TABLES ---------------------------------------------------------

                                            let tempTable = [];

                                            if (tableName !== "Shipping" && tableName !== "Ignore") {

                                                //====================================================================================================
                                                    //#region CREATE SHEETS AND TABLES IN EXCEL FOR TYPE BREAKOUT ------------------------------------

                                                        let thisTable = addSheetAndTable(line, allSheets, filteredData, masterHeader, tableName);
                                                    
                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region GET NUMBER OF ROWS IN TYPE BREAKOUT ----------------------------------------------------

                                                        let normalRowCount = thisTable.table.rows.getCount();

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                let tableSortedValues = thisTable.table.getDataBodyRange().load("values");


                                                await context.sync();

                                                let sortedValues = tableSortedValues.values;

                                                let sortedTableData = sortedValues.map(row => {
                                                    let obj = {};
                                                    row.forEach((cell, index) => {
                                                        obj[masterHeader[0][index]] = cell;
                                                    });
                                                    return obj;
                                                });

                                                //====================================================================================================
                                                    //#region FORMAT CELLS IN TYPE BREAKOUT ----------------------------------------------------------
                                                        let theNormalFormat = styleCells(
                                                            thisTable.table, globalVar.normalBreakoutsFormatting[line], normalRowCount.value, 
                                                            masterHeader, tableName, sortedTableData
                                                        );

                                                        if (tableName.match(/\bDIGITAL\b/gi)){
                                                            context.workbook.tables.getItem("Digital").sort.apply([
                                                                { key: 0, ascending: true } // '0' is the column number for 'Form'
                                                            ]);
                                                        } else if (tableName== "Fold Only"){
                                                            context.workbook.tables.getItem("FoldOnly").sort.apply([
                                                                { key: 5, ascending: true }, // '5' is the column number for 'product'
                                                                { key: 6, ascending: true },  // '6' is the column number for 'company'
                                                                { key: 3, ascending: true },  // '3' is the column number for 'code'
                                                            ]);
                                                        }
                                                        
                                                        await context.sync();

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region HIDE CERTAIN COLUMNS IN TYPE BREAKOUT --------------------------------------------------

                                                        hideColumns(thisTable.table, columnsToHide);

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region SET PRINT SETTINGS FOR TYPE BREAKOUT ---------------------------------------------------

                                                        printSettings(thisTable.sheet);

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                            };

                                        //#endregion -------------------------------------------------------------------------------------------------
                                    //================================================================================================================

                                    await context.sync();

                                };

                            //#endregion -------------------------------------------------------------------------------------------------------------
                        //============================================================================================================================

                        //============================================================================================================================
                            //#region CONSOLE LOG ALL BREAKOUT DATA ----------------------------------------------------------------------------------

                                console.log("FILTERED DATA: ", filteredData);
                                console.log("MISSING FORMS: ", missing);
                                console.log("IGNORED ITEMS: ", ignore);
                                console.log("EMPTY TYPE DATA: ", empty);
                                console.log("DIGITAL TYPE DATA: ", digitalBreakout);

                            //#endregion -------------------------------------------------------------------------------------------------------------
                        //============================================================================================================================


                        //============================================================================================================================
                            //#region CREATE & FORMAT SHIPPING BREAKOUT ------------------------------------------------------------------------------

                                //====================================================================================================================
                                    //#region CREATE SHEET & TABLE FOR SHIPPING BREAKOUT -------------------------------------------------------------

                                        let shippingTable = addSheetAndTable("Shipping", allSheets, missingData, masterHeader, "Shipping");

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                await context.sync();

                                //====================================================================================================================
                                    //#region COUNT ROWS IN SHIPPING BREAKOUT ------------------------------------------------------------------------

                                        let shippingRowCount = shippingTable.table.rows.getCount();

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================


                                let shippingSortedValues = shippingTable.table.getDataBodyRange().load("values");

                                await context.sync();

                                let shippingSorted = shippingSortedValues.values;

                                let shippingSortedTableData = shippingSorted.map(row => {
                                    let obj = {};
                                    row.forEach((cell, index) => {
                                        obj[masterHeader[0][index]] = cell;
                                    });
                                    return obj;
                                });

                                //====================================================================================================================
                                    //#region FORMAT SHIPPING BREAKOUT -------------------------------------------------------------------------------

                                       
                                            let theShippingFormat = styleCells(
                                                shippingTable.table, shippingFormatting, shippingRowCount.value, masterHeader, "Shipping", shippingSortedTableData
                                            );

                                            context.workbook.tables.getItem("Shipping").sort.apply([
                                                { key: 0, ascending: true } // '0' is the column number for 'Form'
                                            ]);
                                        

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                await context.sync();
                                //====================================================================================================================
                                    //#region HIDE CERTAIN COLUMNS IN SHIPPING BREAKOUT --------------------------------------------------------------
                                
                                        hideColumns(shippingTable.table, columnsToHide);

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                //====================================================================================================================
                                    //#region SET PRINT SETTINGS FOR SHIPPING BREAKOUT ---------------------------------------------------------------
                                    
                                        printSettings(shippingTable.sheet);

                                    //#endregion -----------------------------------------------------------------------------------------------------
                                //====================================================================================================================

                                await context.sync();

                            //#endregion -------------------------------------------------------------------------------------------------------------
                        //============================================================================================================================

                        //if any of the breakout sheets do not have any info in them, console log that they were empty this time
                        if (globalVar.emptySheets) {
                            console.log("The following sheets are empty:");
                            console.log(globalVar.emptySheets);
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                activateEvents(); //turn on events

                //====================================================================================================================================
                    //#region REMOVE LOADING SCREEN --------------------------------------------------------------------------------------------------

                        if ($("#breakout").text() == "Breakout") {
                            $("#breakout").text("Delete Breakout Sheets");
                        } else {
                            $("#breakout").text("Breakout");
                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

            });

            //? Maybe we can move this into remove loading screen region? But I don't wanna break anything right now so let's try this later
            $("#loading-background").css("display", "none"); 

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region HIDE COLUMNS FUNCTION ------------------------------------------------------------------------------------------------------------------

        /**
         * Hides the columns in the provided array within the provided table. Also sets custom column widths and text wrap settings for the following:
         * Address, Company, & Notes columns
         * @param {Excel Table Object} table The table object that you wish to hide some columns in
         * @param {Array} hideColArr The columns that you wish to hide, in an array
         */
        async function hideColumns(table, hideColArr) {
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

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region STYLE CELLS FUNCTION -------------------------------------------------------------------------------------------------------------------

        /**
         * Formats each cell in the table provided based on the formatting values in the formattingArr
         * @param {Excel Table Object} table The table object that you wish to style cells for
         * @param {Array} formattingArr An array containing all the formatting values for each cell in the table
         * @param {Number} rowCount The number of rows in the table
         * @param {Array} headerValues An array of the headers in the table
         * @param {String} sheetName The name of the table
         * @returns An array of objects
         */  
        function styleCells(table, formattingArr, rowCount, headerValues, sheetName, sortedData) {


            if (sheetName == "PlasticLine") {
                console.log("PLASTIC!");
            };

            let tempObj = {};

            let allFormattedCells = [];

            headerValues = headerValues[0];

            if (rowCount == 0) {
                globalVar.emptySheets.push(sheetName);
            };

            for (let u = 0; u < rowCount; u++) {
                //the problem is the u variable being 1 ahead of the v variable

                try {


                    //let arrRow = formattingArr[u];

                    //let currentRow = formattingArr[u]['UJID'].formsValue;

                    let thisRow = table.rows.getItemAt(u).getRange();

                    let currentRow = sortedData[u]['UJID'];
            


                    //! The first row in the formattingArr is actually the second row in the table for some reason (the first row seems to be missing)
                    //! thisRow is getting the first row in the table, so the wrong formatting is being applied to the first row it seems

                    //! Also formattingArr is not sorted like the actual table is, so a Z-Shelf item that is at the bottom of the excel table is not in the formattingArr

                    let rowProperties = "";

                    // Functional Array Method ↓
                    // arr.find(v => itemRow === targetValue) ; → [{... found value}]
                    // let foundValue = Find the item in globalVar.masterCellData where itemRow == targetValue
                    // let cell = thisRow.getCell(0, v).load("address");
                    // cell.format.fill ...



                    let rowFormats = globalVar.masterCellData[currentRow];

                    // Using the headerValues variable because its index will always match the column number.
                    rowFormats.forEach((cell, i)=>{
                        // Take thisRow
                        let unFormatted = thisRow.getCell(0, i).load("address"); // An Unformatted cell.
                        let currentCellProps = cell.cellProps.value[0][0];

                        unFormatted.format.fill.color = currentCellProps.format.fill.color;
                        unFormatted.format.font.color = currentCellProps.format.font.color;
                        unFormatted.format.font.bold = currentCellProps.format.font.bold;
                        unFormatted.format.font.italic = currentCellProps.format.font.italic;
                        unFormatted.format.wrapText = true;
                    });

                    //use the newProperties below instead of arrRow[headerValues[v]]






                    // for (let v = 0; v < headerValues.length; v++) {

                    //     try {
                    //         // let cell = thisRow.getCell(0, v).load("address");
                    //         // cell.format.fill.color = arrRow[headerValues[v]].formsFill;
                    //         // cell.format.font.color = arrRow[headerValues[v]].formsFontColor;
                    //         // cell.format.font.bold = arrRow[headerValues[v]].formsFontBold;
                    //         // cell.format.font.italic = arrRow[headerValues[v]].formsFontItalic;
                    //         // cell.format.wrapText = true;

                    //         let cell = thisRow.getCell(0, v).load("address");
                    //         cell.format.fill.color = rowProperties.format.fill.color;
                    //         cell.format.font.color = rowProperties.format.font.color;
                    //         cell.format.font.bold = rowProperties.format.font.bold;
                    //         cell.format.font.italic = rowProperties.format.font.italic;
                    //         cell.format.wrapText = true;


                    //         // tempObj = {
                    //         //     sheet: sheetName,
                    //         //     index: u,
                    //         //     cell: cell,
                    //         //     fill: arrRow[headerValues[v]].formsFill,
                    //         //     fontColor: arrRow[headerValues[v]].formsFontColor,
                    //         //     fontBold: arrRow[headerValues[v]].formsFontBold,
                    //         //     fontItalic: arrRow[headerValues[v]].formsFontItalic,
                    //         //     wrapText: true,
                    //         // };

                    //         tempObj = {
                    //             sheet: sheetName,
                    //             index: u,
                    //             cell: cell,
                    //             fill: rowProperties.format.fill.color,
                    //             fontColor: rowProperties.format.font.color,
                    //             fontBold: rowProperties.format.font.bold,
                    //             fontItalic: rowProperties.format.font.italic,
                    //             wrapText: true,
                    //         };

                    //         allFormattedCells.push(tempObj);

                    //     } catch (e) {
                    //         // console.log(`The value of v at error: ${v}`);
                    //         console.log(e);
                    //         loadError(e.stack)
                    //     };

                    // };
                } catch (e) {
                    console.log(e);
                    loadError(e.stack)
                };

            };

            return allFormattedCells;

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region PRINT SETTINGS FUNCTION ----------------------------------------------------------------------------------------------------------------

        /**
         * Sets the specified sheet up with the print settings that Tyndra typically uses
         * @param {Excel Sheet Object} sheet The sheet that you wish to apply print settings to
         */
function printSettings(sheet) {
    const points = 72; // 72 points in an inch
    /*
    Currently they are vertically & horizontally centered. Just keep horizontal, remove vertical centering. Also change Margin Top to 0.5� and Margin Bottom to 0.25�
    */
    sheet.pageLayout.rightMargin = 0;
    sheet.pageLayout.leftMargin = 0;
    sheet.pageLayout.topMargin = 0.5 * points;
    sheet.pageLayout.bottomMargin = 0.25 * points;
    sheet.pageLayout.headerMargin = 0;
    sheet.pageLayout.footerMargin = 0;

    sheet.pageLayout.paperSize = "Legal";

    sheet.pageLayout.centerHorizontally = true;
    sheet.pageLayout.centerVertically = false;

    let pageLayoutZoomOptions = {
        'horizontalFitToPages': 1,
        'verticalFitToPages': 0,
    };

    sheet.pageLayout.zoom = pageLayoutZoomOptions;

    // Set the first row as the title row for every page.
    sheet.pageLayout.setPrintTitleRows("$2:$2");

    sheet.pageLayout.orientation = Excel.PageOrientation.landscape;

    sheet.pageLayout.printGridlines= true;
}

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region ADD SHEET AND TABLE FUNCTION -----------------------------------------------------------------------------------------------------------
           
        /**
         * 
         * @param {String} line The name that you will be calling the sheet (also the value that shows up in the "Type" column)
         * @param {Excel Worksheet Collection} allSheets A collection of all the worksheets in the workbook
         * @param {Object} filteredData An object containing all the row values for the sheet that is going to be made
         * @param {Array} masterHeader An array of arrays containing the header values for the master sheet
         * @param {String} tableName What you want the table's name to be
         * @param {Object} masterRowInfo Information on the master row
         * @returns 
         */
         function addSheetAndTable(line, allSheets, filteredData, masterHeader, tableName, masterRowInfo) {

            let table = "";
            let tableColumnLetter = "";

            //add and name the new sheet
            let sheet = allSheets.add(line);
            sheet.load("name, position");


            let tableRowLength = filteredData[line].length;
            let tableHeaderLength = masterHeader[0].length;

            //converts the column index of the last column in the table to it's appropriate letter
            tableColumnLetter = printToLetter(tableHeaderLength);

            let tableTitleString = "A1:" + tableColumnLetter + "1"; //makes a range that is the length of the table in the first row of the sheet
            let tableTitleRange = sheet.getRange(tableTitleString); //actually makes it into a range variable now

            //merges those cells into one big cell that we will use as the title row for the table
            tableTitleRange.merge(true);

            let mergedTitleRange = sheet.getRange("A1"); //get range of the new title row

            //set formatting values of the title cell
            // console.log("Format Func:", line)
            mergedTitleRange.values = [[`${globalVar.headerPrefix} ${line}`]];
            mergedTitleRange.format.autofitColumns();
            mergedTitleRange.format.horizontalAlignment = "center";
            mergedTitleRange.format.verticalAlignment = "center";
            mergedTitleRange.format.font.size = 22;
            mergedTitleRange.format.font.bold = true;
            mergedTitleRange.format.rowHeight = 62;
            mergedTitleRange.format.fill.color = "#F3EAF7";

            let tableRangeString = "A2:" + tableColumnLetter + "2"; //where the table is going to go

            //adds a table in said range
            table = sheet.tables.add(tableRangeString, true /*hasHeaders*/);
       
            table.name = tableName; //names the table properly

            table.getHeaderRowRange().values = [masterHeader[0]]; //sets the header row values to the masterHeader values

            let sheetAndTable = {
                sheet: "",
                table: ""
            };

            //the table data for this line
            let currentSheet = filteredData[line];

            sheet.getUsedRange().format.autofitColumns();
            sheet.getUsedRange().format.autofitRows();

            // Here's the silliness. Here we go...
            let rangeA = sheet.getRange("A1"); // Stretch out Forms
            rangeA.format.columnWidth = 55; // Pixel numbers get doubled apparently?

            let rangeB = sheet.getRange("B1"); // Stretch out Type
            rangeB.format.columnWidth = 70;

            let rangeC = sheet.getRange("C1"); // Stretch out Rep
            rangeC.format.columnWidth = 65;

            let rangeD = sheet.getRange("D1"); // Stretch out AS
            rangeD.format.columnWidth = 65;

            let rangeF = sheet.getRange("F1"); // Stretch out Product
            rangeF.format.columnWidth = 100;

            let rangeI = sheet.getRange("I1"); // Stretch out Email
            rangeI.format.columnWidth = 170;

            let rangeV = sheet.getRange("V1"); // Stretch out UPS Info
            rangeV.format.columnWidth = 90;

            let rangeW = sheet.getRange("W1"); // Stretch out UJID Info
            rangeW.format.columnWidth = 80;


            const rangeA_E = sheet.getRange("A:E");
            const rangeJ_Q = sheet.getRange("J:Q");

            rangeA_E.format.horizontalAlignment = "Center";
            rangeJ_Q.format.horizontalAlignment = "Center";

            sheet.getRange(`A:W`).format.verticalAlignment="center";

            //if the table data for the line is empty, push the empty sheet and table variables to the sheetAndTable object for referencing later
            if (currentSheet == "") {
                sheet.activate();
                sheetAndTable.sheet = sheet;
                sheetAndTable.table = table;
                return sheetAndTable;
            }

            // MA/UA Sort ----------------------------------------------------------------------------------------------------------------------------
            let uaRows = currentSheet.filter(row => row[0]=="UA");


            uaRows.forEach(row => {

                // Remove this row from currentSheet
                currentSheet.splice(currentSheet.indexOf(row), 1)

                let versionRef = row[masterHeader[0].indexOf("Version No")].split("-")[1];
                let notesRef = row[masterHeader[0].indexOf("Notes")].split("-")[1];
                let codeColumn = row[masterHeader[0].indexOf("Code")];


                // Find the row in currentSheet that has the ref value as the 3rd index
                let refRow = currentSheet.find(row => (row[codeColumn] === Number(versionRef) || row[codeColumn] === Number(notesRef)))
                let targetIndex = currentSheet.indexOf(refRow)
                // Add this row after the row at the target index
                currentSheet.splice(targetIndex + 1, 0, row)

            })

            //adds the line info from filteredData as rows to the end of the table
            table.rows.add(null /*add rows t o the end of the table*/, currentSheet);

            // Type (B), EDDM (P), Order Status (Q), Version No (S), Artwork (T), Options (U), & UJID (W)
            // const columnsToHide = ['B', 'P', 'Q', 'S', 'T', 'U', 'W'];
            // columnsToHide.forEach(column => {
            //     const hideRange = sheet.getRange(`${column}:${column}`);
            //     hideRange.columnHidden = true;
            // });            

             // Have Digital and Shipping Breakouts sorted by Form numbers. For Fold Only Breakout, sort by product, then by company
             

            // Move Z-shelf to the bottom.
            let indexes = []
            let zRows = currentSheet.filter((row, i) => {
                if (row[0] == "ZSHELF") {
                    indexes.push(i)
                    return row 
                }
            });

             console.log(globalVar.normalBreakoutsFormatting);
             console.log(line);

            let formattingRows = globalVar.normalBreakoutsFormatting[line];
            
            zRows.forEach(row => {

                // Remove this row from currentSheet
                currentSheet.splice(currentSheet.indexOf(row), 1)

                // Add to the end.
                currentSheet.push(row)
            })

            // Sorting the formattingRows
             indexes.sort((a, b) => b - a);

             if (indexes) {
                 console.log(indexes);
                 // Sort formattingRows so that the items at the indexes in the array "indexes" are at the bottom
                 //let elementsToMove = indexes.map(index => formattingRows.splice(index, 1)[0]);
                 let elementsToMove = indexes.map((index) => {
                     formattingRows.splice(index, 1)[0];
                 });
                 if (elementsToMove.length > 0) {
                     formattingRows.push(...elementsToMove);
                 }
             } else {
                 console.log("There were no ZHELF items, so skipping the ZSHELF sort step...");
             };

                  

            // Post sort. format.
            // masterRowInfo['UJID'].value



            sheet.activate();

            //push the sheet and table values to the sheetAndTable object for referencing later
            sheetAndTable.sheet = sheet;
            sheetAndTable.table = table;

            return sheetAndTable;

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region CONVERT COLUMN INDEX TO COLUMN LETTER (PRINT TO LETTER) --------------------------------------------------------------------------------

        /**
         * Converts the provided column index number to it's appropriate column letter
         * @param {Number} number The column index that you are trying to convert to a column letter
         * @returns String
         */
        function printToLetter(number) {

            let alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

            let charIndex = number % alphabet.length;
            let quotient = number / alphabet.length;

            if (charIndex - 1 == -1) {

                charIndex = alphabet.length;
                quotient--;

            };

            globalVar.result = alphabet.charAt(charIndex - 1); // + globalVar.result;

            if (quotient >= 1) {

                printToLetter(parseInt(quotient));

            } else {

                return globalVar.result;

            };

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region HANDLE MISSING TYPES -------------------------------------------------------------------------------------------------------------------
        /**
         * Shows a pop-up window that informs the user of rows in the master that are missing data in the Type column. User can decide to continue,
         * where the missing types are automatically set to "All Other" or they can handle manually to exit the function and fix it themselves
         */
        function emptyWarning() {

            deactivateEvents();

            console.log("There are empty types preset. Exiting Breakout function...");

            $("#loading-background").css("display", "none")

            $("#empty-background").css("display", "flex");

            $("#set-to-ignore").on("click", () => {
                console.log("setting to ignore...");
                showTitle = true;
                setBlanksToAllOther();
                $("#empty-background").css("display", "none");
            });

            $("#handle-manually").on("click", () => {
                console.log("letting user handle...");
                showTitle = false;
                $("#empty-background").css("display", "none");
                activateEvents();
            });

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region SET ALL EMPTY TYPES TO "ALL OTHER" (SET BLANKS TO ALL OTHER) ---------------------------------------------------------------------------

        /**
         * Sets all empty Types in the master sheet to "All Other"
         */
        async function setBlanksToAllOther() {

            await Excel.run(async (context) => {

                //====================================================================================================================================
                    //#region ASSIGN SHEET VARIABLES -------------------------------------------------------------------------------------------------

                        const validation = context.workbook.worksheets.getItem("Validation");
                        const linesTable = validation.tables.getItem("Lines");
                        const linesBodyRange = linesTable.getDataBodyRange().load("values");
                        const linesHeaderRange = linesTable.getHeaderRowRange().load("values");
                        const masterSheet = context.workbook.worksheets.getItem("Master").load("name");
                        const masterTable = masterSheet.tables.getItem("Master");
                        const masterBodyRange = masterTable.getDataBodyRange().load("values");
                        const masterHeaderRange = masterTable.getHeaderRowRange().load("values");
                        const masterTableRows = masterTable.rows.load("items");

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                await context.sync();

                //====================================================================================================================================
                    //#region LOAD SHEET VARIABLES ---------------------------------------------------------------------------------------------------

                        let linesArr = linesBodyRange.values;
                        let linesHeader = linesHeaderRange.values;
                        let masterArr = masterBodyRange.values;
                        let masterHeader = masterHeaderRange.values;
                        let masterRowItems = masterTableRows.items;
                        let masterArrCopy = JSON.parse(JSON.stringify(masterBodyRange.values));

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                //====================================================================================================================================
                    //#region LOOP THROUGH MASTER ROWS AND UPDATE ALL BLANK TYPES TO "ALL OTHER" -----------------------------------------------------

                        let z = 0;
                        let masterRowInfo = new Object();

                        for (let masterRow of masterArr) {

                            //the following matches the master table headers with the data and column index in row [z], assigning each as a property
                            // to each header within the masterRowInfo object.
                            for (let name of masterHeader[0]) {
                                createRowInfo(masterHeader, name, masterRow, masterArrCopy, masterRowInfo, z, masterSheet);
                            };

                            let masterTypeColumnIndex = masterRowInfo["Type"].columnIndex;

                            //if type cell is blank, set to "All Other"
                            if (masterArr[z][masterTypeColumnIndex] == "") {
                                masterArr[z][masterTypeColumnIndex] = "All Other";
                            };

                            z = z + 1;

                        };

                    //#endregion ---------------------------------------------------------------------------------------------------------------------
                //====================================================================================================================================

                await context.sync()

                masterBodyRange.values = masterArr; //write values to Excel

                await context.sync();

                breakout(); //commence breakout function

            });

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region REMOVE BREAKOUT SHEETS -----------------------------------------------------------------------------------------------------------------

        /**
         * Sets up which worksheets exist and which ones to include or not include from the delete breakout sheets button click, then deletes them
         * @param {Array} arrOfSheets An array of arrays containing the type lines, or normal breakouts, that are the majority of the breakout sheets
         * @param {Excel Sheet Collection} allSheets A collection of all the worksheets in the workbook
         * @returns 
         */
        async function removeBreakoutSheets(arrOfSheets, allSheets) {

            return new Promise((resolve, reject) => {

                try {

                    //================================================================================================================================
                        //#region CONVERTS ARR OF SHEETS FROM NESTED ARRAY TO A SINGLE ARRAY IN LINES DATA -------------------------------------------

                            globalVar.linesData = [];

                            //builds out an array of all the normal breakout sheets
                            for (var p = 0; p < arrOfSheets.length; p++) {

                                let currentLines = arrOfSheets[p][0];

                                globalVar.linesData.push(currentLines);

                            };

                        //#endregion -----------------------------------------------------------------------------------------------------------------
                    //================================================================================================================================

                    //================================================================================================================================
                        //#region CREATES LIST OF SHEET NAMES TO DELETE AND REMOVES A FEW SPECIFIC ONES FROM THIS LIST -------------------------------

                            let arrOfSheetNames = [];

                            //sheets NOT to delete 
                            let removeThese = [
                                "Validation", "SilkE2R", "TextE2R", "DIGE2R", "Master", "Press Scheduling", "MISSING", "PRINTED", "IGNORE", "DIGITAL"
                            ];

                            //creates an array of all the sheet names of all the CURRENTLY EXISTING sheets
                            for (let sheet of allSheets.items) {
                                arrOfSheetNames.push(sheet.name);
                            };


                            for (let name of removeThese) {
                                //gets index of the sheet name from removeThese inside of arrOfSheetNames
                                let indexOfName = arrOfSheetNames.indexOf(name); 
                                if (indexOfName !== undefined) {
                                    arrOfSheetNames.splice(indexOfName, 1); //removes that sheet name from the hit list
                                };
                            };

                        //#endregion -----------------------------------------------------------------------------------------------------------------
                    //================================================================================================================================

                    //================================================================================================================================
                        //#region MAKE SURE WE DELETE A FEW CUSTOM SHEETS OUTSIDE OF THE LINES DATA INFO ---------------------------------------------

                            deleteSheets("MISSING");
                            deleteSheets("PRINTED");
                            deleteSheets("IGNORE");
                            deleteSheets("DIGITAL");

                        //#endregion -----------------------------------------------------------------------------------------------------------------
                    //================================================================================================================================

                    //================================================================================================================================
                        //#region DELETE ALL OTHER SHEETS IN LINES DATA AND ARR OF SHEET NAMES -------------------------------------------------------

                            for (let line of globalVar.linesData) {
                                for (let item of arrOfSheetNames) {
                                    if (item == line) {
                                        deleteSheets(line);
                                        arrOfSheetNames.splice(arrOfSheetNames.indexOf(item), 1);
                                        break;
                                    };
                                };
                            };

                        //#endregion -----------------------------------------------------------------------------------------------------------------
                    //================================================================================================================================

                    resolve("Done.");

                } catch (e) {
                    loadError(e.stack)
                    reject(e);
                };

                activateEvents();

            });

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region DELETE SHEET ---------------------------------------------------------------------------------------------------------------------------

        /**
         * Deletes a specific worksheet from the workbook
         * @param {String} worksheetName The name of the worksheet to delete
         */
        async function deleteSheets(worksheetName) {

            await Excel.run(async (context) => {

                try {
                    let deleteSheet = context.workbook.worksheets.getItem(worksheetName);
                    deleteSheet.delete();

                    await context.sync();
                } catch (e) {
                    console.log(`${worksheetName} does not exist, so it was not deleted!`)
                };
      
            });

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

export { breakout, removeBreakoutSheets };