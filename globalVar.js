export default {
    silkLayouts: [],
    sheetHourData: {},
    breakoutData: {},
    productData: {},
    wasteData: {},
    linesData:[],
    formToCarry: undefined,
    changeEvent: undefined,
    eventResult: undefined,
    formsColumnIndex: undefined,
    result: "",
    scrollErr: undefined,
    silkTable: undefined,
    textTable: undefined,
    digTable: undefined,
    monday:[],
    tuesday:[],
    wednesday:[],
    thursday:[],
    friday:[],
    silkDataSet:[],
    textDataSet:[],
    digDataSet:[],
    operators: {},
    scrollHeight: undefined,
    listOfBreakoutTables:[],
    emptySheets:[],
    normalBreakoutsFormatting: [],
    hiddenLinesData: ["MISSING", "IGNORE", "Shipping", "Empty", "PRINTED", "DIGITAL"],
    priorityNum: 1,
    singleSided: false,
    dotSelected: false,
    plannedSep: false,
    rushItem: false,
    opTableCounter: 1,
    rowMoved: false,
    pressmen:[" "],
    presses:[" "],
    daysOfWeek: [" ", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
    headerPrefix: "",
    masterCellData: {},
    tableAndSheetNames: {
        // Key = Worksheet Name.
        // Value: Table Name.
        "MISSING": "Missing",
        "Missing": "Missing", // Just in case
        "PRINTED": "Printed",
        "IGNORE": "Ignore",
        "DIGITAL": "Digital",
        "Postcard": "Postcard",
        "Scratch-Off Line": "ScratchoffLine",
        "Magnet Line": "MagnetLine",
        "Colossal": "Colossal",
        "MENU": "Menu",
        "Fold Only": "FoldOnly",
        "XL Ink Line": "XlInkLine",
        "MS-BS Flyer Line": "MsbsFlyerLine",
        "Plastic Line": "PlasticLine",
        "Envelope Inserter": "EnvelopeInserter",
        "Heidelberg Die-Cutter": "HeidelbergDiecutter",
        "Flatbed Die-Cutter": "FlatbedDiecutter",
        "MF Ink Line": "MfInkLine",
        "Apparel": "Apparel",
        "APPAREL": "Apparel", // For the Forms thing.
        "All Other": "AllOther",
        "Shipping": "Shipping"
    }
    //pressmen: [" ", "Steve", "Roberto", "Ryan", "Jamie", "Cody", "Terry", "Paul"],
    //presses: [" ", 1, 2, 3, 4, "Digital 1", "Digital 2"]
}




       
