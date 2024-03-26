import globalVar from "./globalVar.js";
import { deactivateEvents, activateEvents, createDataSet, conditionalFormatting, createRowInfo } from "./universalFunctions.js";
import { easyToReads } from "./E2Rs.js";
import { pressSchedulingInfoTable } from "./pressSchedulingInfo.js";
import { breakout } from "./breakout.js";

//======================================================================================================================================================
    //#region POPULATE FORMS FUNCTION -----------------------------------------------------------------------------------------------------------------
            
        /**
         * Populates the E2Rs with the proper form numbers, quantities, sheets, and hours based on the info in the Metrix Info column. 
         * Also takes this info from the E2Rs and compares it to the Master sheet, finding the items in the master in each of the E2Rs 
         * and assigning each row in the master sheet the proper form number from the E2R data and gives it a type based on the validation sheet. 
         * Finally, it builds out the taskpane info from the press scheduing Info table
         */
        async function populateForms() {

            deactivateEvents(); //turns off workbook events

            //========================================================================================================================================
                //#region TURNS ON LOADING WINDOW IN TASKPANE ---------------------------------------------------------------------------------------

                    $("#loading-background").css("display", "flex");

                    $("#loading-head").text("Populating Forms");

                //#endregion ------------------------------------------------------------------------------------------------------------------------
            //========================================================================================================================================

            try {

                await Excel.run(async (context) => { //loads context so I can directly pull and write info to Excel within this async function

                    //===============================================================================================================================
                        //#region ASSIGNS SHEET VARIABLES --------------------------------------------------------------------------------------------

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

                        //#endregion ----------------------------------------------------------------------------------------------------------------
                    //================================================================================================================================

                    await context.sync();

                    //================================================================================================================================
                        //#region LOAD SHEET VARIABLES ----------------------------------------------------------------------------------------------

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

                        //#endregion ----------------------------------------------------------------------------------------------------------------
                    //================================================================================================================================

                    //================================================================================================================================
                        //#region CREATE DATA OBJECTS FOR VALIDATION SHEET INFO ---------------------------------------------------------------------

                            //=======================================================================================================================
                                //#region CREATE SHEET HOURS DATA OBJECT FROM VALIDATION -----------------------------------------------------------

                                    var x = 0

                                    let tempArr = []; //temporary array for storing values to push into the data object (clears each time)

                                    let valueToPush;

                                    for (var item of sheetHourHeader[0]) { //for each header cell in thr sheet hour table...

                                        //y will never be larger than the number of rows in the sheet hour table
                                        for (var y = 0; y < sheetHourArr.length; y++) { 

                                            //* The last maximum value should typically be everything after the last minimum number, or infinity. 
                                            //* In the table, this is typically expressed by a "+", "-", or just an empty cell. This translates that 
                                            //* into the javascript equivilant for infinity
                                            if ((sheetHourArr[y][x] == "-" || sheetHourArr[y][x] == "+" || sheetHourArr[y][x] == "") 
                                            && sheetHourArr[y][x] !== 0) {
                                                valueToPush = Infinity;
                                            } else {
                                                //if not infinity, then just use the current value without translation
                                                valueToPush = sheetHourArr[y][x]; 
                                            };

                                            //pushes the xth cell in row y to the tempArr, then repeats for the same cell in each row
                                            tempArr.push(valueToPush); 

                                        };

                                        //assigns all the values for each column header to itself as an object in globalVar.sheetHourData
                                        globalVar.sheetHourData[item] = tempArr; 

                                        tempArr = []; //clears the tempArr so it can be clean as we loop through the parent for loop

                                        x = x + 1;

                                    };

                                //#endregion ---------------------------------------------------------------------------------------------------------
                            //========================================================================================================================

                            //========================================================================================================================
                                //#region CREATE PRODUCT DATA OBJECT FROM VALIDATION -----------------------------------------------------------------

                                    let tempProdObj = {}; //temporary object for storing values to push into the data object (clears each time)

                                    for (var u = 0; u < productArr.length; u++) {

                                        let leProduct = productArr[u];

                                        let g = 0;

                                        for (var title of productHeader[0]) {

                                            tempProdObj[title] = leProduct[g];

                                            g = g + 1;

                                        };

                                        globalVar.productData[tempProdObj["Name"]] = tempProdObj;

                                        tempProdObj = {};

                                    };

                                    //================================================================================================================
                                        //#region (LEGACY) CODE FOR OLD BREAKOUT TABLE OBJECT --------------------------------------------------------

                                            // let v = 0;

                                            // let breakoutTempArr = [];

                                            // for (var value of breakoutHeader[0]) {
                                            //   for (var w = 0; w < breakoutArr.length; w++) {
                                            //     breakoutTempArr.push(breakoutArr[w][v])
                                            //   }

                                            //   globalVar.breakoutData[value] = breakoutTempArr;

                                            //   breakoutTempArr = [];

                                            //   v = v + 1;

                                            // };

                                        //#endregion -------------------------------------------------------------------------------------------------
                                    //================================================================================================================

                                //#endregion ---------------------------------------------------------------------------------------------------------
                            //========================================================================================================================

                            //========================================================================================================================
                                //#region CREATE WASTE DATA OBJECT FROM VALIDATION -------------------------------------------------------------------

                                    let tempWasteObj = {}; //temporary object for storing values to push into the data object (clears each time)

                                    for (var t = 0; t < wasteArr.length; t++) {

                                        let currentWaste = wasteArr[t];

                                        let p = 0;

                                        for (var heading of wasteHeader[0]) {

                                            tempWasteObj[heading] = currentWaste[p];

                                            p = p + 1;

                                        };

                                        globalVar.wasteData[tempWasteObj["Type"]] = tempWasteObj;

                                        tempWasteObj = {};

                                    };

                                //#endregion ---------------------------------------------------------------------------------------------------------
                            //========================================================================================================================

                        //#endregion -----------------------------------------------------------------------------------------------------------------
                    //================================================================================================================================

                    //================================================================================================================================
                        //#region POPULATE FORM COLUMN IN E2R TABLES ---------------------------------------------------------------------------------

                            let isSilkEmpty = false;
                            let isTextEmpty = false;
                            let isDigEmpty = false;
                            
                            if (silkE2RArr.length > 0 && silkE2RArr[0][0] !== "" || silkE2RArr.length > 1) {
                                let silkE2RUpdate = easyToReads(silkE2RArr, silkE2RTableRows, silkE2RSheet, sheetHourArr);
                                //push the updated values commited to silkE2RArr into the SilkE2R table in Excel
                                silkE2RBodyRange.values = silkE2RUpdate; 
                            } else {
                                isSilkEmpty = true;
                            };

                            if (textE2RArr.length > 0 && textE2RArr[0][0] !== "" || textE2RArr.length > 1) {
                                let textE2RUpdate = easyToReads(textE2RArr, textE2RTableRows, textE2RSheet, sheetHourArr);
                                //push the updated values commited to textE2RArr into the TextE2R table in Excel
                                textE2RBodyRange.values = textE2RUpdate; 
                            } else {
                                isTextEmpty = true;
                            };

                            if (digE2RArr.length > 0 && digE2RArr[0][0] !== "" || digE2RArr.length > 1) {
                                let digE2RUpdate = easyToReads(digE2RArr, digE2RTableRows, digE2RSheet, sheetHourArr);
                                digE2RBodyRange.values = digE2RUpdate; //push the updated values commited to digE2RArr into the DIGE2R table in Excel
                            } else {
                                isDigEmpty = true;
                            };

                        //#endregion -----------------------------------------------------------------------------------------------------------------
                    //================================================================================================================================

                    //================================================================================================================================
                        //#region POPULATE MASTER INFO -----------------------------------------------------------------------------------------------

                            //========================================================================================================================
                                //#region RE-ASSIGN SHEET VARIABLES FOR MASTER TABLE AUTO-FILL -------------------------------------------------------

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

                                //#endregion ---------------------------------------------------------------------------------------------------------
                            //========================================================================================================================

                            await context.sync()

                            //========================================================================================================================
                                //#region RE-LOAD SHEET VARIABLES FOR MASTER AUTO-FILL ---------------------------------------------------------------

                                    silkE2RArr = silkE2RBodyRangeUpdate.values; //moves all values of the SilkE2R table to an array
                                    textE2RArr = textE2RBodyRangeUpdate.values; //moves all values of the TextE2R table to an array
                                    digE2RArr = digE2RBodyRangeUpdate.values; //moves all values of the TextE2R table to an array
                                    // let pressSchedulingArr = pressSchedulingInfoBodyRange.values;

                                //#endregion ---------------------------------------------------------------------------------------------------------
                            //========================================================================================================================

                            //========================================================================================================================
                                //#region BUILD E2R DATA SETS ----------------------------------------------------------------------------------------

                                    globalVar.silkDataSet = [];
                                    globalVar.textDataSet = [];
                                    globalVar.digDataSet = [];

                                    globalVar.priorityNum = 1;

                                    globalVar.silkDataSet = createDataSet(silkE2RArr, "Silk");
                                    globalVar.textDataSet = createDataSet(textE2RArr, "Text");
                                    globalVar.digDataSet = createDataSet(digE2RArr, "Digital");

                                //#endregion ---------------------------------------------------------------------------------------------------------
                            //========================================================================================================================

                            //========================================================================================================================
                                //#region UPDATE PRESS SCHEDULING TABLE ------------------------------------------------------------------------------

                                    //updates the press scheduling table in the validation sheet to the new data 
                                    //(using the "Taskpane" variable so it knows how to treat the incoming data based on the source of the change)
                                    pressSchedulingInfoTable("Populate");

                                //#endregion ---------------------------------------------------------------------------------------------------------
                            //========================================================================================================================

                            //========================================================================================================================
                                //#region BUILD HEADER, ITEM, AND COPY ARRAY VARIABLES FOR E2RS ------------------------------------------------------

                                    //assigns header range valies for each E2R
                                    const silkHeader = silkHeaderRange.values;
                                    const textHeader = textHeaderRange.values;
                                    const digHeader = digHeaderRange.values;


                                    //assigns row items for each E2R
                                    let silkRowItems = silkE2RTableRowsUpdate.items;
                                    let textRowItems = textE2RTableRowsUpdate.items;
                                    let digRowItems = digE2RTableRowsUpdate.items;


                                    //creates un-linked copies of the E2R data
                                    let silkArrCopy = JSON.parse(JSON.stringify(silkE2RBodyRangeUpdate.values));
                                    let textArrCopy = JSON.parse(JSON.stringify(textE2RBodyRangeUpdate.values));
                                    let digArrCopy = JSON.parse(JSON.stringify(digE2RBodyRangeUpdate.values));
                                    //! Could be done with map or filter, possibly better?

                                //#endregion ---------------------------------------------------------------------------------------------------------
                            //========================================================================================================================

                            //========================================================================================================================
                                //#region FILL MASTER FORMS ------------------------------------------------------------------------------------------

                                    let z = 0;

                                    let masterRowInfo = new Object();

                                    let missingForms = [];

                                    for (var masterRow of masterArr) { //for each row in the master sheet...

                                        //============================================================================================================
                                            //#region FUNCTION VARIABLES -----------------------------------------------------------------------------

                                                let wasItSilk = false;
                                                let wasItText = false;
                                                let wasItDig = false;

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                        //============================================================================================================
                                            //#region CREATE MASTER SHEET ROW INFO -------------------------------------------------------------------

                                                //the following matches the master table headers with the data and column index in row [z], 
                                                //assigning each as a property to each header within the masterRowInfo object.
                                                for (var name of masterHeader[0]) {
                                                    createRowInfo(masterHeader, name, masterRow, masterArrCopy, masterRowInfo, z, masterSheet);
                                                };

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                        //============================================================================================================
                                            //#region ASSIGN MASTER VARIABLES FROM ROW INFO ----------------------------------------------------------

                                                let masterUJID = masterRowInfo["UJID"].value;

                                                //returns the client code of the the current row in the master table
                                                let masterCode = masterRowInfo["Code"].value; 

                                                //returns the product of the current row in the master table
                                                let masterProduct = masterRowInfo["Product"].value; 

                                                //returns the version ID of the current row in the master table
                                                let masterVersion = masterRowInfo["Version No"].value; 

                                                let masterOptions = masterRowInfo["Options"].value;

                                                let masterWeeks = masterRowInfo["Wks"].value;

                                                let masterQuantity = masterRowInfo["Qty"].value;

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                        //============================================================================================================
                                            //#region HANDLE EMPTY VARIABLES IN ROW ------------------------------------------------------------------

                                                if (masterUJID == "" && masterProduct == "") {
                                                    console.log(`
                                                        Looks like row ${z + 2} is missing the UJID, Product, and possibly even more important info. 
                                                        Please update this info and run the "Populate Forms" function again to fix this line
                                                    `);
                                                    z = z + 1;
                                                    break;
                                                } else if (masterUJID == "" && masterProduct !== "") {
                                                    console.log(`
                                                        Looks like row ${z + 2} is missing a UJID. Please update this info and run the 
                                                        "Populate Forms" function again to fix this line
                                                    `);
                                                    z = z + 1;
                                                    break;
                                                } else if (masterUJID !== "" && masterProduct == "") {
                                                    console.log(`
                                                        Looks like row ${z + 2} is missing the Product. Please update this info and run the 
                                                        "Populate Forms" function again to fix this line
                                                    `);
                                                    z = z + 1;
                                                    break;
                                                };

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                        //============================================================================================================
                                            //#region SETS MASTER BREAKOUT TO EITHER FOLD ONLY OR THE PRODUCT DATA'S DEFAULT BREAKOUT VARIABLE -------

                                                let masterBreakout; //this will be used much later to set the type of the row in the master sheet

                                                let potentalFoldOnly = [
                                                    "MENU", "NonProfit80#", "x.80#custom", "x.Flyer.10.5x17", "x.Flyer.8.5x10.5", "x.Menu.10.5x17"
                                                ];

                                                let foldOnlyOverwrite = false;

                                                try {

                                                    //if product of row is a potentalFoldOnly, force masterBreakout to be "Fold Only"
                                                    for (let menuType of potentalFoldOnly) {
                                                        if (masterProduct == menuType && (masterWeeks == 0 || masterQuantity == 0)) {
                                                            masterBreakout = "Fold Only";
                                                            foldOnlyOverwrite = true;
                                                        };
                                                    };

                                                    //Otherwise, use the default breakout for the product data
                                                    if (!foldOnlyOverwrite) {
                                                        masterBreakout = globalVar.productData[masterProduct]["Breakout"];
                                                    };
                                                } catch (e) {
                                                    console.log(e);
                                                };

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                        //============================================================================================================
                                            //#region CARRY FORMS FROM E2RS TO MASTER ARRAY ----------------------------------------------------------

                                                //====================================================================================================
                                                    //#region FORM CARRY OVERRIDES -------------------------------------------------------------------

                                                        let isApperal = false;
                                                        let ignoreMissing = false;

                                                        //if masterProduct is an apparel item, set isApperal to true
                                                        for (let apparelItem of apparelArr) {
                                                            if (masterProduct == apparelItem) {
                                                                isApperal = true;
                                                            };
                                                        };

                                                        //if masterProduct is an ignore item, set ignoreMissing to true
                                                        for (let ignoreItem of defaultToIgnoreArr) {
                                                            if (masterProduct == ignoreItem) {
                                                                ignoreMissing = true;
                                                            };
                                                        };

                                                        //if master sheet Options column for the current row has ZSHELF in it, 
                                                        //overwrite the form number with "ZSHELF"
                                                        if (masterOptions.includes("ZSHELF")) {

                                                            globalVar.formToCarry = "ZSHELF";

                                                        //if master sheet Version column for the current row has UA- in it, 
                                                        //overwrite the form number with "UA"
                                                        } else if (masterVersion.includes("UA-")) {

                                                            globalVar.formToCarry = "UA";

                                                        //set form to "APPAREL" if isApperal is true
                                                        } else if (isApperal == true) {

                                                            globalVar.formToCarry = "APPAREL";

                                                        //set form to "IGNORE" if ignoreMissing is true
                                                        } else if (ignoreMissing == true) {

                                                            globalVar.formToCarry = "IGNORE";

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                } else {

                                                //====================================================================================================
                                                    //#region TRY TO FIND FORM NUMBER IN SILKE2R -----------------------------------------------------

                                                        let a = 0;

                                                        if (isSilkEmpty == false) {

                                                            for (var silkRow of silkE2RArr) {

                                                                wasItSilk = carryForm(
                                                                    silkRow, silkArrCopy, silkRowItems, silkHeader, masterUJID, a, silkE2RSheet
                                                                );

                                                                if (wasItSilk) {
                                                                    break;
                                                                };

                                                                a = a + 1; //repeat until we have gone through all the rows in the SilkE2R table

                                                            };

                                                        };

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region TRY TO FIND FORM NUMBER IN TEXTE2R -----------------------------------------------------

                                                        //if the product was not in the SilkE2R, then we move on to check if it is in the TextE2R
                                                        if (!wasItSilk) { 

                                                            if (isTextEmpty == false) {

                                                                let b = 0;

                                                                for (var textRow of textE2RArr) {

                                                                    wasItText = carryForm(
                                                                        textRow, textArrCopy, textRowItems, textHeader, masterUJID, b, textE2RSheet
                                                                    );

                                                                    if (wasItText) {
                                                                        break;
                                                                    };

                                                                    b = b + 1; //repeat until we have gone through all the rows in the SilkE2R table

                                                                };

                                                            };

                                                        };

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region TRY TO FIND FORM NUMBER IN DIGE2R ------------------------------------------------------

                                                        //if the product was not found it either the text or silkE2R, then we move onto the DIGE2R
                                                        if (!wasItSilk && !wasItText) { 

                                                            if (isDigEmpty == false) {

                                                                let c = 0;

                                                                for (var digRow of digE2RArr) {

                                                                    wasItDig = carryForm(
                                                                        digRow, digArrCopy, digRowItems, digHeader, masterUJID, c, digE2RSheet
                                                                    );

                                                                    if (wasItDig) {
                                                                        break;
                                                                    };

                                                                    c = c + 1;

                                                                };

                                                            };

                                                        };

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region IF ALL E2RS ARE EMPTY, HANDLE AS PLANNED SEPARETELY ------------------------------------

                                                        if (isSilkEmpty && isTextEmpty && isDigEmpty) {

                                                            console.log("All E2Rs are empty, so this must be a Planned Separately Form");
                                                            globalVar.plannedSep = true;

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region IF NONE OF E2RS HAVE THE FORM, ADD ROW TO MISSING FORMS ARRAY --------------------------

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

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                };

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                        //============================================================================================================
                                            //#region APPLY CORRECT FORM NUMBER TO MASTER ------------------------------------------------------------

                                                let masterFormColumnIndex = masterRowInfo["Forms"].columnIndex;

                                                //if globalVar.plannedSeperately is not true, push form to master
                                                if (!globalVar.plannedSep) {

                                                    //if formsToCarry is false, push "MISSING" to form in master
                                                    if (globalVar.formToCarry) {
                                                        masterArr[z][masterFormColumnIndex] = globalVar.formToCarry;
                                                    } else {
                                                        masterArr[z][masterFormColumnIndex] = "MISSING";
                                                    };

                                                };

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                        //============================================================================================================
                                            //#region APPLY CORRECT BREAKOUT TYPE TO MASTER ----------------------------------------------------------

                                                let masterTypeColumnIndex = masterRowInfo["Type"].columnIndex;

                                                masterArr[z][masterTypeColumnIndex] = masterBreakout;

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                        //============================================================================================================
                                            //#region APPLY DATA VALIDATION TO CERTAIN MASTER SHEET COLUMNS ------------------------------------------

                                                //====================================================================================================
                                                    //#region APPLY DATA VALIDATION TO NON-NUMBERED FORMS --------------------------------------------

                                                        let leRange = "A" + (z + 2); //row number + 1 for 0-index and + 1 for header row

                                                        let zeRange = masterSheet.getRange(leRange); //home on da RANGE UwU

                                                        zeRange.dataValidation.clear(); //dear existing data validation: clear out, fool!

                                                        zeRange.format.fill.clear(); //same goes for you cell color >:|
                                                        zeRange.format.font.bold = false; //bold font have been FALSIFIED

                                                        //? Look at microsoft's documentation as to how this function works, 
                                                        //? I have forgotten and am tired
                                                        let typeRange = masterSheet.getRangeByIndexes(z + 2, masterTypeColumnIndex, 1, 1);

                                                        typeRange.dataValidation.clear(); //clear it, son!

                                                        //* If the value in the form column is NOT a number, give it data validation and formatting!
                                                        if (!Number(masterArr[z][masterFormColumnIndex])) {
                                                            // console.log(masterArr[z][masterFormColumnIndex] + "is not a number!");
                                                            let dv = {
                                                                list: {
                                                                    inCellDropdown: true, //DROPDOWN!
                                                                    source: customFormsBodyRange //this be the Custom Forms table from the validation
                                                                    //(twas loaded in long ago at the top of the function)
                                                                }
                                                            };

                                                            zeRange.dataValidation.rule = dv;

                                                            //sets the conditional formatting for the form cell
                                                            conditionalFormatting(masterSheet, zeRange, masterArr[z][masterFormColumnIndex], null);

                                                            await context.sync();

                                                        };

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                                //====================================================================================================
                                                    //#region APPLY DATA VALIDATION TO TYPE COLUMN ---------------------------------------------------

                                                        // console.log(masterArr[z][masterFormColumnIndex] + "is not a number!");
                                                        let dvType = {
                                                            list: {
                                                                inCellDropdown: true,
                                                                source: linesBodyRange
                                                            }
                                                        };

                                                        typeRange.dataValidation.rule = dvType;

                                                    //#endregion -------------------------------------------------------------------------------------
                                                //====================================================================================================

                                            //#endregion ---------------------------------------------------------------------------------------------
                                        //============================================================================================================

                                        z = z + 1;

                                    };

                                    masterBodyRange.values = masterArr; //write masterArr to the master table in Excel

                                    console.log(`There are ${missingForms.length} missing forms, listed here:`, missingForms);

                                //#endregion ---------------------------------------------------------------------------------------------------------
                            //========================================================================================================================

                        //#endregion -----------------------------------------------------------------------------------------------------------------
                    //================================================================================================================================

                    activateEvents(); //turns on workbook events

                    location.reload(); //reloads the workbook

                });

            } catch (err) {
                console.error(err);
                // showMessage(error, "show");
            };

            //========================================================================================================================================
                //#region TURNS OFF LOADING WINDOW IN TASKPANE ---------------------------------------------------------------------------------------

                    $("#loading-background").css("display", "none");

                    $("#loading-background").css("display", "none");

                //#endregion -------------------------------------------------------------------------------------------------------------------------
            //========================================================================================================================================
        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region CARRY FORMS ----------------------------------------------------------------------------------------------------------------------------

        /**
         * Compares the product and client code of the current line in the master table to the current line in the current E2R table.
         * If it finds a match, it pushes the form value to the global globalVar.formToCarry object. Function also returns true if match is found.
         * @param {Array} row An array of all the values in the row
         * @param {Array} tableArrayCopy A shallow copy of the tableArray to be used for manipulating stuff without changing the original array
         * @param {Array} tableRowItems An array of all the items in the current row
         * @param {Range} tableHeader The header range of the table
         * @param {String} masterProduct The product from the master table to compare to the table row product
         * @param {Number} masterCode The client code from the master table to compare to the table row client code
         * @param {Object} worksheet The worksheet object of the E2R
         */
         function carryForm(row, tableArrayCopy, tableRowItems, tableHeader, masterUJID, rowIndex, worksheet) {

            globalVar.formToCarry = "";

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

                if (doesItMatch) { //if it was in the silkE2R table, then we create an object for the silk row and carry over the form number

                    //the following matches the SilkE2R table headers with the data and column index in row [a], assigning each as a property 
                    //to each header within the rowInfo object.
                    for (var rowName of tableHeader[0]) {
                        createRowInfo(tableHeader, rowName, rowValues, tableArrayCopy, rowInfo, rowIndex, worksheet);
                    };

                    globalVar.formToCarry = rowInfo["Form"].value; //gets the form number value from the row

                    // console.log(globalVar.formToCarry);

                    return true;

                };

            };

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================

//====================================================================================================================================================
    //#region COMPARE MASTER TO E2R FUNCTION ---------------------------------------------------------------------------------------------------------

        //! Don't think I use this code anymore, but I'll leave it active for now while I am reorganizing stuff

        /**
         * Compares the Client Code and Product Abbrevations from the current row in the Master Sheet to the current row 
         * in the current E2R table.
         * @param {Array} masterProdAbbr An array of all the product abbreveations listed for the current line's product in the Master sheet
         * @param {String} compareToProd The product (already abbreviated) from the current line in the E2R table to compare 
         * to the Master Product
         * @param {String} compareToCode The client code from the current line in the E2R table to compare to the Master Client Code
         * @param {String} masterCode The client code from the current line in the Master sheet 
         * @returns Boolean
         */
        function compareMasterToE2R(masterProdAbbr, compareToProd, compareToCode, masterCode) {

            for (var k = 0; k < masterProdAbbr.length; k++) {
                if (compareToProd == masterProdAbbr[k] && compareToCode == masterCode) {
                    return true;
                };
            };
            return false;

        };

    //#endregion -------------------------------------------------------------------------------------------------------------------------------------
//====================================================================================================================================================


export { populateForms };