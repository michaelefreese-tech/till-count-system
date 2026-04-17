// main.gs

//----------------------------------------------------------------------------------------------------
// DICTIONARIES AND VARIABLES
//----------------------------------------------------------------------------------------------------

// Main configuration dictionary
const Info = {
  "SPLASH": {
    "STRING": {
      "C4": "Dominos",
      "C8": "Till Count Sheet",
      "G13": "PLEASE ENTER EMPLOYEE ID:"},
    "FORMULAS": {},
    "COLORS": {
      "close": ["B2:P24"],
      "white": ["K13:M14"]},
    "VARIABLES": {},
    "OUTLINEMED": [
      "K13:M14",
      "B2:P24",],
    "OUTLINETHIN": [],
    "MERGE": [],
    "RIGHT": [],
    "MERGERIGHT": [],
    "MERGELEFT": [],
    "MERGECENTER": [
      "C4:O6",
      "C8:O9",
      "G13:J14",
      "K13:M14",],
    "ALLOWEDCELLS": ["K13"],
    "DATACELLS": []},

  "COVER": {
    "STRING": {
      "B24": "EMPLOYEE ID:",
      "C8": "TILL 1",
      "F8": "TILL 2",
      "I8": "SAFE",
      "C10":"100$",
      "C11": "50$",
      "C12": "20$",
      "C13": "10$",
      "C14" : "5$",
      "C15": "1$",
      "C16": ".25$",
      "C17": ".10$",
      "C18": ".05$",
      "C19": ".01$",
      "C21": "Total :",
      "F10":"100$",
      "F11": "50$",
      "F12": "20$",
      "F13": "10$",
      "F14" : "5$",
      "F15": "1$",
      "F16": ".25$",
      "F17": ".10$",
      "F18": ".05$",
      "F19": ".01$",
      "F21": "Total :",
      "I10":"100$",
      "I11": "50$",
      "I12": "20$",
      "I13": "10$",
      "I14" : "5$",
      "I15": "1$",
      "I16": ".25$",
      "I17": ".10$",
      "I18": ".05$",
      "I19": ".01$",
      "I21": "Total :",
      "N13": "PREVIOUS SHIFT DATE :",
      "N15": "PREVIOUS SHIFT :",
      "N17": "CURRENT DATE : ",
      "N19": "CURRENT SHIFT :",
      "N21": "DIFFERENCE :",
      "H24": "Date To Search:"},
    "FORMULAS": {
      "D21": "=SUM(D10:D19)",
      "G21": "=SUM(G10:G19)",
      "J21": "=SUM(J10:J19)",
      "O15": "=IF($N$2=\"SHIFT\",OPEN!M21,IF($N$2=\"CLOSE\",SHIFT!M21,IF($N$2=\"OPEN\",IFERROR(INDEX(WEEK_1!AG33:AG39,MATCH($O$13,WEEK_1!B33:B39,0)),IFERROR(INDEX(WEEK_2!AG33:AG39,MATCH($O$13,WEEK_2!B33:B39,0)),IFERROR(INDEX(WEEK_3!AG33:AG39,MATCH($O$13,WEEK_3!B33:B39,0)),IFERROR(INDEX(WEEK_4!AG33:AG39,MATCH($O$13,WEEK_4!B33:B39,0)),0)))),0)))",
      "O19": "=SUM(D21,G21,J21)",
      "O21": "=O15-O19",
      "O17": "=TODAY()"},
    "FORMAT": {
      "currency": ["C10:C19", "F10:F19", "I10:I19", "C21", "F21", "I21", "O19"]},
    "COLORS": {
      "yellow": ["C10:C19", "F10:F19", "I10:I19", "C21", "F21", "I21"],
      "white": ["D10:D19", "G10:G19", "J10:J19", "D21", "G21", "J21", "O13", "O15", "O17", "O19", "O21"],
      "orange": ["C8", "F8", "I8", "N13", "N15", "N17", "N19", "N21"]},
    "VARIABLES": {
      "SHIFTINDICATOR": "N2"},
    "OUTLINEMED": [
      "B7:K22", "C8:D8", "C10:C19", "D10:D19", "C21:D21", "F8:G8", "F10:F19", "G10:G19", 
      "F21:G21", "I8:J8", "I10:I19", "J10:J19", "I21:J21", "B24:C24", "D24:E24", "J24:K24", 
      "M7:P22", "N13", "O13", "N15", "O15", "N17", "O17", "N19", "N2:O5", "O19", "N21", "O21"],
    "OUTLINETHIN": [],
    "MERGE": [],
    "RIGHT": [
      "C10", "C11", "C12", "C13", "C14", "C15", "C16", "C17", "C18", "C19", "C21",
      "D10", "D11", "D12", "D13", "D14", "D15", "D16", "D17", "D18", "D19", "D21",
      "F10", "F11", "F12", "F13", "F14", "F15", "F16", "F17", "F18", "F19", "F21",
      "G10", "G11", "G12", "G13", "G14", "G15", "G16", "G17", "G18", "G19", "G21",
      "I10", "I11", "I12", "I13", "I14", "I15", "I16", "I17", "I18", "I19", "I21",
      "J10", "J11", "J12", "J13", "J14", "J15", "J16", "J17", "J18", "J19", "J21",
      "M21"],
    "MERGERIGHT": [],
    "MERGELEFT": [],
    "MERGECENTER": ["B24:C24", "D24:E24", "J24:K24", "C8:D8", "F8:G8", "I8:J8", "N2:O5"],
    "ALLOWEDCELLS": [
      "D10", "D11", "D12", "D13", "D14", "D15", "D16", "D17", "D18", "D19",
      "G10", "G11", "G12", "G13", "G14", "G15", "G16", "G17", "G18", "G19",
      "J10", "J11", "J12", "J13", "J14", "J15", "J16", "J17", "J18", "J19",
      "J24"],
    "DATACELLS": ["D10:D19", "G10:G19", "J10:J19"],
    "MENU": {
      "open": ["B7:K22", "M7:P22"],
      "shift": ["B7:K22", "M7:P22"],
      "close": ["B7:K22", "M7:P22"],
      "search": ["B7:K22", "M7:P22"]
    }},

  "OPEN": {
    "STRING": {
      "C8": "TILL 1",
      "F8": "TILL 2",
      "I8": "SAFE",

      "C10":"100$",
      "C11": "50$",
      "C12": "20$",
      "C13": "10$",
      "C14" : "5$",
      "C15": "1$",
      "C16": ".25$",
      "C17": ".10$",
      "C18": ".05$",
      "C19": ".01$",
      "F10":"100$",
      "F11": "50$",
      "F12": "20$",
      "F13": "10$",
      "F14" : "5$",
      "F15": "1$",
      "F16": ".25$",
      "F17": ".10$",
      "F18": ".05$",
      "F19": ".01$",
      "I10":"100$",
      "I11": "50$",
      "I12": "20$",
      "I13": "10$",
      "I14" : "5$",
      "I15": "1$",
      "I16": ".25$",
      "I17": ".10$",
      "I18": ".05$",
      "I19": ".01$",

      "C21": "Total :",
      "F21": "Total :",
      "I21": "Total :",},

    "FORMULAS": {
      "D21": "=SUM(D10:D19)",
      "G21": "=SUM(G10:G19)",
      "J21": "=SUM(J10:J19)",
      "M21": "=SUM(D21,G21,J21)"},

    "FORMAT": {
      "currency": ["C10:C19", "F10:F19", "I10:I19"]},
    "COLORS": {
      "open": ["B7:K22"],
      "orange": ["C8:D8", "F8:G8", "I8:J8"],
      "yellow": ["C10:C19", "F10:F19", "I10:I19", "C21", "F21", "I21"],
      "white": ["D10:D19", "G10:G19", "J10:J19", "D21", "G21", "J21"]},

    "VARIABLES": {},
    "OUTLINEMED": [
        "B7:K22", "C8:D8", "C10:C19", "D10:D19", "C21:D21", "F8:G8", "F10:F19", "G10:G19", 
        "F21:G21", "I8:J8", "I10:I19", "J10:J19", "I21:J21", 
        ],

    "OUTLINETHIN": [],
    "MERGE": [
      "B22:K22",
      "B8:B21",
      "E8:E21",
      "H8:H21",
      "K8:K21"],

    "RIGHT": [
      "C10", "C11", "C12", "C13", "C14", "C15", "C16", "C17", "C18", "C19", "C21",
      "D10", "D11", "D12", "D13", "D14", "D15", "D16", "D17", "D18", "D19", "D21",
      "F10", "F11", "F12", "F13", "F14", "F15", "F16", "F17", "F18", "F19", "F21",
      "G10", "G11", "G12", "G13", "G14", "G15", "G16", "G17", "G18", "G19", "G21",
      "I10", "I11", "I12", "I13", "I14", "I15", "I16", "I17", "I18", "I19", "I21",
      "J10", "J11", "J12", "J13", "J14", "J15", "J16", "J17", "J18", "J19", "J21",
      "M21"],

    "MERGERIGHT": [],
    "MERGELEFT": [],
    "MERGECENTER": [
      "C8:D8",
      "F8:G8",
      "I8:J8"],

    "ALLOWEDCELLS": [],
    "DATACELLS": ["D10:D19", "G10:G19", "J10:J19"]},

  "SHIFT": {
    "STRING": {
      "C8": "TILL 1",
      "F8": "TILL 2",
      "I8": "SAFE",

      "C10":"100$",
      "C11": "50$",
      "C12": "20$",
      "C13": "10$",
      "C14" : "5$",
      "C15": "1$",
      "C16": ".25$",
      "C17": ".10$",
      "C18": ".05$",
      "C19": ".01$",
      "F10":"100$",
      "F11": "50$",
      "F12": "20$",
      "F13": "10$",
      "F14" : "5$",
      "F15": "1$",
      "F16": ".25$",
      "F17": ".10$",
      "F18": ".05$",
      "F19": ".01$",
      "I10":"100$",
      "I11": "50$",
      "I12": "20$",
      "I13": "10$",
      "I14" : "5$",
      "I15": "1$",
      "I16": ".25$",
      "I17": ".10$",
      "I18": ".05$",
      "I19": ".01$",

      "C21": "Total :",
      "F21": "Total :",
      "I21": "Total :",},
    "FORMULAS": {
      "D21": "=SUM(D10:D19)",
      "G21": "=SUM(G10:G19)",
      "J21": "=SUM(J10:J19)",
      "M21": "=SUM(D21,G21,J21)"},
    "FORMAT": {
      "currency": ["C10:C19", "F10:F19", "I10:I19"]},
    "COLORS": {
      "shift": ["B7:K22"],
      "orange": ["C8:D8", "F8:G8", "I8:J8"],
      "yellow": ["C10:C19", "F10:F19", "I10:I19", "C21", "F21", "I21"],
      "white": ["D10:D19", "G10:G19", "J10:J19", "D21", "G21", "J21"]},
    "VARIABLES": {},
    "OUTLINEMED": [
        "B7:K22", "C8:D8", "C10:C19", "D10:D19", "C21:D21", "F8:G8", "F10:F19", "G10:G19", 
        "F21:G21", "I8:J8", "I10:I19", "J10:J19", "I21:J21", 
        ],
    "OUTLINETHIN": [],
    "MERGE": [],
    "RIGHT": [
      "C10", "C11", "C12", "C13", "C14", "C15", "C16", "C17", "C18", "C19", "C21",
      "D10", "D11", "D12", "D13", "D14", "D15", "D16", "D17", "D18", "D19", "D21",
      "F10", "F11", "F12", "F13", "F14", "F15", "F16", "F17", "F18", "F19", "F21",
      "G10", "G11", "G12", "G13", "G14", "G15", "G16", "G17", "G18", "G19", "G21",
      "I10", "I11", "I12", "I13", "I14", "I15", "I16", "I17", "I18", "I19", "I21",
      "J10", "J11", "J12", "J13", "J14", "J15", "J16", "J17", "J18", "J19", "J21",
      "M21"],
    "MERGERIGHT": [],
    "MERGELEFT": [],
    "MERGECENTER": [
      "C8:D8",
      "F8:G8",
      "I8:J8"],
    "ALLOWEDCELLS": [],
    "DATACELLS": ["D10:D19", "G10:G19", "J10:J19"]},

  "CLOSE": {
    "STRING": {
      "C8": "TILL 1",
      "F8": "TILL 2",
      "I8": "SAFE",

      "C10":"100$",
      "C11": "50$",
      "C12": "20$",
      "C13": "10$",
      "C14" : "5$",
      "C15": "1$",
      "C16": ".25$",
      "C17": ".10$",
      "C18": ".05$",
      "C19": ".01$",
      "F10":"100$",
      "F11": "50$",
      "F12": "20$",
      "F13": "10$",
      "F14" : "5$",
      "F15": "1$",
      "F16": ".25$",
      "F17": ".10$",
      "F18": ".05$",
      "F19": ".01$",
      "I10":"100$",
      "I11": "50$",
      "I12": "20$",
      "I13": "10$",
      "I14" : "5$",
      "I15": "1$",
      "I16": ".25$",
      "I17": ".10$",
      "I18": ".05$",
      "I19": ".01$",

      "C21": "Total :",
      "F21": "Total :",
      "I21": "Total :",},
    "FORMULAS": {
      "D21": "=SUM(D10:D19)",
      "G21": "=SUM(G10:G19)",
      "J21": "=SUM(J10:J19)",
      "M21": "=SUM(D21,G21,J21)"},
    "FORMAT": {
      "currency": ["C10:C19", "F10:F19", "I10:I19"]},
    "COLORS": {
      "close": ["B7:K22"],
      "orange": ["C8:D8", "F8:G8", "I8:J8"],
      "yellow": ["C10:C19", "F10:F19", "I10:I19", "C21", "F21", "I21"],
      "white": ["D10:D19", "G10:G19", "J10:J19", "D21", "G21", "J21"]},
    
    "VARIABLES": {},
    "OUTLINEMED": [
        "B7:K22", "C8:D8", "C10:C19", "D10:D19", "C21:D21", "F8:G8", "F10:F19", "G10:G19", 
        "F21:G21", "I8:J8", "I10:I19", "J10:J19", "I21:J21", 
        ],
    "OUTLINETHIN": [],
    "MERGE": [],
    "RIGHT": [
      "C10", "C11", "C12", "C13", "C14", "C15", "C16", "C17", "C18", "C19", "C21",
      "D10", "D11", "D12", "D13", "D14", "D15", "D16", "D17", "D18", "D19", "D21",
      "F10", "F11", "F12", "F13", "F14", "F15", "F16", "F17", "F18", "F19", "F21",
      "G10", "G11", "G12", "G13", "G14", "G15", "G16", "G17", "G18", "G19", "G21",
      "I10", "I11", "I12", "I13", "I14", "I15", "I16", "I17", "I18", "I19", "I21",
      "J10", "J11", "J12", "J13", "J14", "J15", "J16", "J17", "J18", "J19", "J21",
      "M21"],
    "MERGERIGHT": [],
    "MERGELEFT": [],
    "MERGECENTER": [
      "C8:D8",
      "F8:G8",
      "I8:J8"],
    "ALLOWEDCELLS": [],
    "DATACELLS": ["D10:D19", "G10:G19", "J10:J19"]},
  "SEARCH": {
      "STRING": {
        "C8": "TILL 1",
        "F8": "TILL 2",
        "I8": "SAFE",

        "C10":"100$",
        "C11": "50$",
        "C12": "20$",
        "C13": "10$",
        "C14" : "5$",
        "C15": "1$",
        "C16": ".25$",
        "C17": ".10$",
        "C18": ".05$",
        "C19": ".01$",
        "F10":"100$",
        "F11": "50$",
        "F12": "20$",
        "F13": "10$",
        "F14" : "5$",
        "F15": "1$",
        "F16": ".25$",
        "F17": ".10$",
        "F18": ".05$",
        "F19": ".01$",
        "I10":"100$",
        "I11": "50$",
        "I12": "20$",
        "I13": "10$",
        "I14" : "5$",
        "I15": "1$",
        "I16": ".25$",
        "I17": ".10$",
        "I18": ".05$",
        "I19": ".01$",

        "C21": "Total :",
        "F21": "Total :",
        "I21": "Total :",},
      "FORMULAS": {
        "D21": "=SUM(D10:D19)",
        "G21": "=SUM(G10:G19)",
        "J21": "=SUM(J10:J19)",
        "M21": "=SUM(D21,G21,J21)"},
      "FORMAT": {
        "currency": ["C10:C19", "F10:F19", "I10:I19"]},
      "COLORS": {
        "search": ["B7:K22"],
        "orange": ["C8:D8", "F8:G8", "I8:J8"],
        "yellow": ["C10:C19", "F10:F19", "I10:I19", "C21", "F21", "I21"],
        "white": ["D10:D19", "G10:G19", "J10:J19", "D21", "G21", "J21"]},
      
      "VARIABLES": {},
      "OUTLINEMED": [
          "B7:K22", "C8:D8", "C10:C19", "D10:D19", "C21:D21", "F8:G8", "F10:F19", "G10:G19", 
          "F21:G21", "I8:J8", "I10:I19", "J10:J19", "I21:J21", 
          ],
      "OUTLINETHIN": [],
      "MERGE": [],
      "RIGHT": [
        "C10", "C11", "C12", "C13", "C14", "C15", "C16", "C17", "C18", "C19", "C21",
        "D10", "D11", "D12", "D13", "D14", "D15", "D16", "D17", "D18", "D19", "D21",
        "F10", "F11", "F12", "F13", "F14", "F15", "F16", "F17", "F18", "F19", "F21",
        "G10", "G11", "G12", "G13", "G14", "G15", "G16", "G17", "G18", "G19", "G21",
        "I10", "I11", "I12", "I13", "I14", "I15", "I16", "I17", "I18", "I19", "I21",
        "J10", "J11", "J12", "J13", "J14", "J15", "J16", "J17", "J18", "J19", "J21",
        "M21"],
      "MERGERIGHT": [],
      "MERGELEFT": [],
      "MERGECENTER": [
        "C8:D8",
        "F8:G8",
        "I8:J8"],
      "ALLOWEDCELLS": [],
      "DATACELLS": ["D10:D19", "G10:G19", "J10:J19"]},
  
  "WEEK_1": {
    "STRING": {
      "C3": "OPEN SHIFT",
      "C5": "TILL 1",
      "M5": "TILL 2",
      "W5": "SAFE",
      "B6": "Date",
      "AG6": "TOTAL",
      "AH6": "Submitted on",
      "AI6": "Submitted by",
      "A7": "Mon",
      "A8": "Tues",
      "A9": "Wed",
      "A10": "Thur",
      "A11": "Fri",
      "A12": "Sat",
      "A13": "Sun",

      "C16": "SHIFT CHANGE",
      "C18": "TILL 1",
      "M18": "TILL 2",
      "W18": "SAFE",
      "AG19": "TOTAL",
      "AH19": "Submitted on",
      "AI19": "Submitted by",
      "A20": "Mon",
      "A21": "Tues",
      "A22": "Wed",
      "A23": "Thur",
      "A24": "Fri",
      "A25": "Sat",
      "A26": "Sun",

      "C29": "CLOSE SHIFT",
      "C31": "TILL 1",
      "M31": "TILL 2",
      "W31": "SAFE",
      "AG32": "TOTAL",
      "AH32": "Submitted on",
      "AI32": "Submitted by",
      "A33": "Mon",
      "A34": "Tues",
      "A35": "Wed",
      "A36": "Thur",
      "A37": "Fri",
      "A38": "Sat",
      "A39": "Sun",

      // Open Currency
      "C6": "100$", "D6": "50$", "E6": "20$", "F6": "10$", "G6": "5$", "H6": "1$", "I6": ".25$", "J6": ".10$", "K6": ".05$", "L6": ".01$",  
      "M6": "100$", "N6": "50$", "O6": "20$", "P6": "10$", "Q6": "5$", "R6": "1$", "S6": ".25$", "T6": ".10$", "U6": ".05$", "V6": ".01$",
      "W6": "100$", "X6": "50$", "Y6": "20$", "Z6": "10$", "AA6": "5$", "AB6": "1$", "AC6": ".25$", "AD6": ".10$", "AE6": ".05$", "AF6": ".01$",

      // Shift Currency
      "C19": "100$", "D19": "50$", "E19": "20$", "F19": "10$", "G19": "5$", "H19": "1$", "I19": ".25$", "J19": ".10$", "K19": ".05$", "L19": ".01$",
      "M19": "100$", "N19": "50$", "O19": "20$", "P19": "10$", "Q19": "5$", "R19": "1$", "S19": ".25$", "T19": ".10$", "U19": ".05$", "V19": ".01$",
      "W19": "100$", "X19": "50$", "Y19": "20$", "Z19": "10$", "AA19": "5$", "AB19": "1$", "AC19": ".25$", "AD19": ".10$", "AE19": ".05$", "AF19": ".01$",

      // Close Currency
      "C32": "100$", "D32": "50$", "E32": "20$", "F32": "10$", "G32": "5$", "H32": "1$", "I32": ".25$", "J32": ".10$", "K32": ".05$", "L32": ".01$",
      "M32": "100$", "N32": "50$", "O32": "20$", "P32": "10$", "Q32": "5$", "R32": "1$", "S32": ".25$", "T32": ".10$", "U32": ".05$", "V32": ".01$",
      "W32": "100$", "X32": "50$", "Y32": "20$", "Z32": "10$", "AA32": "5$", "AB32": "1$", "AC32": ".25$", "AD32": ".10$", "AE32": ".05$", "AF32": ".01$"},

    "FORMULAS": {
      // Open B7:B13
      "B7": "=(INFO!D3)",
      "B8": "=SUM(B7 + 1)",
      "B9": "=SUM(B8 + 1)",
      "B10": "=SUM(B9 + 1)",
      "B11": "=SUM(B10 + 1)",
      "B12": "=SUM(B11 + 1)",
      "B13": "=SUM(B12 + 1)",

      "AG7": "=SUM(C7:AF7)",
      "AG8": "=SUM(C8:AF8)",
      "AG9": "=SUM(C9:AF9)",
      "AG10": "=SUM(C10:AF10)",
      "AG11": "=SUM(C11:AF11)",
      "AG12": "=SUM(C12:AF12)",
      "AG13": "=SUM(C13:AF13)",

      // Shift B20:B16
      "B20": "=(INFO!D3)",
      "B21": "=SUM(B20 + 1)",
      "B22": "=SUM(B21 + 1)",
      "B23": "=SUM(B22 + 1)",
      "B24": "=SUM(B23 + 1)",
      "B25": "=SUM(B24 + 1)",
      "B26": "=SUM(B25 + 1)",

      "AG20": "=SUM(C20:AF20)",
      "AG21": "=SUM(C21:AF21)",
      "AG22": "=SUM(C22:AF22)",
      "AG23": "=SUM(C23:AF23)",
      "AG24": "=SUM(C24:AF24)",
      "AG25": "=SUM(C25:AF25)",
      "AG26": "=SUM(C26:AF26)",

      //Close B33:B39
      "B33": "=(INFO!D3)",
      "B34": "=SUM(B33 + 1)",
      "B35": "=SUM(B34 + 1)",
      "B36": "=SUM(B35 + 1)",
      "B37": "=SUM(B36 + 1)",
      "B38": "=SUM(B37 + 1)",
      "B39": "=SUM(B38 + 1)",

      "AG33": "=SUM(C33:AF33)",
      "AG34": "=SUM(C34:AF34)",
      "AG35": "=SUM(C35:AF35)",
      "AG36": "=SUM(C36:AF36)",
      "AG37": "=SUM(C37:AF37)",
      "AG38": "=SUM(C38:AF38)",
      "AG39": "=SUM(C39:AF39)"},

    "FORMAT": {
      "currency": ["C6:AF6", "C19:AF19", "C32:AF32"]},
    "COLORS": {
      "open": ["C3:AF4"],
      "shift": ["C16:AF17"],
      "close": ["C29:AF30"],
      "orange": ["C5:L5", "M5:V5", "W5:AF5", "C18:L18", "M18:V18", "W18:AF18", "C31:L31", "M31:V31", "W31:AF31"]},

    "VARIABLES": {},
    "OUTLINEMED": [
      //_________
      //Open
      //_________

      //Header
      "C3:AF4",

      //Till 1, Till 2, Safe
      "C5:L5", "M5:V5", "W5:AF5",

      // Till 1, Till 2, Safe Denominations, Total, Submitted on/ Submitted by
      "C6:L6", "M6:V6", "W6:AF6", "AG6", "AH6:AI6",

      // Days, Dates, DATACELLS, Totals, Submitted on, Subbmitted by
      "A7:A13","B7:B13", "C7:AF13", "AG7:AG13", "AH7:AH13", "AI7:AI13",

      //_________
      //Shift
      //_________

      //Header
      "C16:AF17",

      //Till 1, Till 2, Safe
      "C18:L18", "M18:V18", "W18:AF18",

      // Till 1, Till 2, Safe Denominations, Total, Submitted on/ Submitted by
      "C19:L19", "M19:V19", "W19:AF19", "AG19", "AH19:AI19",

      // Days, Dates, DATACELLS, Totals, Submitted on, Subbmitted by
      "A20:A26", "B20:B26", "C20:AF26", "AG20:AG26", "AH20:AH26", "AI20:AI26",

      //_________
      //Close
      //_________

      //Header
      "C29:AF30",

      //Till 1, Till 2, Safe
      "C31:L31", "M31:V31", "W31:AF31",

      // Till 1, Till 2, Safe Denominations, Total, Submitted on/ Submitted by
      "C32:L32", "M32:V32", "W32:AF32", "AG32", "AH32:AI32",

      // Days, Dates, DATACELLS, Totals, Submitted on, Subbmitted by
      "A33:A39", "B33:B39", "C33:AF39", "AG33:AG39", "AH33:AH39", "AI33:AI39"],

    "OUTLINETHIN": [],
    "MERGE": [],
    "RIGHT": [
      //Open Currency
      "C6", "D6", "E6", "F6", "G6", "H6", "I6", "J6", "K6", "L6",
      "M6", "N6", "O6", "P6", "Q6", "R6", "S6", "T6", "U6", "V6",
      "W6", "X6", "Y6", "Z6", "AA6", "AB6", "AC6", "AD6", "AE6", "AF6",

      // Shift Currency
      "C19", "D19", "E19", "F19", "G19", "H19", "I19", "J19", "K19", "L19",
      "M19", "N19", "O19", "P19", "Q19", "R19", "S19", "T19", "U19", "V19",
      "W19", "X19", "Y19", "Z19", "AA19", "AB19", "AC19", "AD19", "AE19", "AF19",

      //Close Currency
      "C32", "D32", "E32", "F32", "G32", "H32", "I32", "J32", "K32", "L32",
      "M32", "N32", "O32", "P32", "Q32", "R32", "S32", "T32", "U32", "V32",
      "W32", "X32", "Y32", "Z32", "AA32", "AB32", "AC32", "AD32", "AE32", "AF32"],

    "MERGERIGHT": [],
    "MERGELEFT": [],
    "MERGECENTER": [
      // Open
      "C3:AF4",
      "C5:L5", "M5:V5", "W5:AF5",

      //Shift
      "C16:AF17",
      "C18:L18", "M18:V18", "W18:AF18",

      //Close
      "C29:AF30",
      "C31:L31", "M31:V31", "W31:AF31"],

    "ALLOWEDCELLS": [],
    "DATACELLS": ["C7:AF13", "C20:AF26", "C33:AF39"]},

  "WEEK_2": {
    "STRING": {
      "C3": "OPEN SHIFT",
      "C5": "TILL 1",
      "M5": "TILL 2",
      "W5": "SAFE",
      "B6": "Date",
      "AG6": "TOTAL",
      "AH6": "Submitted on",
      "AI6": "Submitted by",
      "A7": "Mon",
      "A8": "Tues",
      "A9": "Wed",
      "A10": "Thur",
      "A11": "Fri",
      "A12": "Sat",
      "A13": "Sun",

      "C16": "SHIFT CHANGE",
      "C18": "TILL 1",
      "M18": "TILL 2",
      "W18": "SAFE",
      "AG19": "TOTAL",
      "AH19": "Submitted on",
      "AI19": "Submitted by",
      "A20": "Mon",
      "A21": "Tues",
      "A22": "Wed",
      "A23": "Thur",
      "A24": "Fri",
      "A25": "Sat",
      "A26": "Sun",

      "C29": "CLOSE SHIFT",
      "C31": "TILL 1",
      "M31": "TILL 2",
      "W31": "SAFE",
      "AG32": "TOTAL",
      "AH32": "Submitted on",
      "AI32": "Submitted by",
      "A33": "Mon",
      "A34": "Tues",
      "A35": "Wed",
      "A36": "Thur",
      "A37": "Fri",
      "A38": "Sat",
      "A39": "Sun",

      // Open Currency
      "C6": "100$", "D6": "50$", "E6": "20$", "F6": "10$", "G6": "5$", "H6": "1$", "I6": ".25$", "J6": ".10$", "K6": ".05$", "L6": ".01$",  
      "M6": "100$", "N6": "50$", "O6": "20$", "P6": "10$", "Q6": "5$", "R6": "1$", "S6": ".25$", "T6": ".10$", "U6": ".05$", "V6": ".01$",
      "W6": "100$", "X6": "50$", "Y6": "20$", "Z6": "10$", "AA6": "5$", "AB6": "1$", "AC6": ".25$", "AD6": ".10$", "AE6": ".05$", "AF6": ".01$",

      // Shift Currency
      "C19": "100$", "D19": "50$", "E19": "20$", "F19": "10$", "G19": "5$", "H19": "1$", "I19": ".25$", "J19": ".10$", "K19": ".05$", "L19": ".01$",
      "M19": "100$", "N19": "50$", "O19": "20$", "P19": "10$", "Q19": "5$", "R19": "1$", "S19": ".25$", "T19": ".10$", "U19": ".05$", "V19": ".01$",
      "W19": "100$", "X19": "50$", "Y19": "20$", "Z19": "10$", "AA19": "5$", "AB19": "1$", "AC19": ".25$", "AD19": ".10$", "AE19": ".05$", "AF19": ".01$",

      // Close Currency
      "C32": "100$", "D32": "50$", "E32": "20$", "F32": "10$", "G32": "5$", "H32": "1$", "I32": ".25$", "J32": ".10$", "K32": ".05$", "L32": ".01$",
      "M32": "100$", "N32": "50$", "O32": "20$", "P32": "10$", "Q32": "5$", "R32": "1$", "S32": ".25$", "T32": ".10$", "U32": ".05$", "V32": ".01$",
      "W32": "100$", "X32": "50$", "Y32": "20$", "Z32": "10$", "AA32": "5$", "AB32": "1$", "AC32": ".25$", "AD32": ".10$", "AE32": ".05$", "AF32": ".01$"},

    "FORMULAS": {
      // Open B7:B13
      "B7": "=(WEEK_1!B13 + 1)",
      "B8": "=SUM(B7 + 1)",
      "B9": "=SUM(B8 + 1)",
      "B10": "=SUM(B9 + 1)",
      "B11": "=SUM(B10 + 1)",
      "B12": "=SUM(B11 + 1)",
      "B13": "=SUM(B12 + 1)",

      "AG7": "=SUM(C7:AF7)",
      "AG8": "=SUM(C8:AF8)",
      "AG9": "=SUM(C9:AF9)",
      "AG10": "=SUM(C10:AF10)",
      "AG11": "=SUM(C11:AF11)",
      "AG12": "=SUM(C12:AF12)",
      "AG13": "=SUM(C13:AF13)",

      // Shift B20:B16
      "B20": "=(WEEK_1!B26 + 1)",
      "B21": "=SUM(B20 + 1)",
      "B22": "=SUM(B21 + 1)",
      "B23": "=SUM(B22 + 1)",
      "B24": "=SUM(B23 + 1)",
      "B25": "=SUM(B24 + 1)",
      "B26": "=SUM(B25 + 1)",

      "AG20": "=SUM(C20:AF20)",
      "AG21": "=SUM(C21:AF21)",
      "AG22": "=SUM(C22:AF22)",
      "AG23": "=SUM(C23:AF23)",
      "AG24": "=SUM(C24:AF24)",
      "AG25": "=SUM(C25:AF25)",
      "AG26": "=SUM(C26:AF26)",

      //Close B33:B39
      "B33": "=(WEEK_1!B39 + 1)",
      "B34": "=SUM(B33 + 1)",
      "B35": "=SUM(B34 + 1)",
      "B36": "=SUM(B35 + 1)",
      "B37": "=SUM(B36 + 1)",
      "B38": "=SUM(B37 + 1)",
      "B39": "=SUM(B38 + 1)",

      "AG33": "=SUM(C33:AF33)",
      "AG34": "=SUM(C34:AF34)",
      "AG35": "=SUM(C35:AF35)",
      "AG36": "=SUM(C36:AF36)",
      "AG37": "=SUM(C37:AF37)",
      "AG38": "=SUM(C38:AF38)",
      "AG39": "=SUM(C39:AF39)"},

    "FORMAT": {
      "currency": ["C6:AF6", "C19:AF19", "C32:AF32"]},
    "COLORS": {
      "open": ["C3:AF4"],
      "shift": ["C16:AF17"],
      "close": ["C29:AF30"],
      "orange": ["C5:L5", "M5:V5", "W5:AF5", "C18:L18", "M18:V18", "W18:AF18", "C31:L31", "M31:V31", "W31:AF31"]},

    "VARIABLES": {},
    "OUTLINEMED": [
      //_________
      //Open
      //_________

      //Header
      "C3:AF4",

      //Till 1, Till 2, Safe
      "C5:L5", "M5:V5", "W5:AF5",

      // Till 1, Till 2, Safe Denominations, Total, Submitted on/ Submitted by
      "C6:L6", "M6:V6", "W6:AF6", "AG6", "AH6:AI6",

      // Days, Dates, DATACELLS, Totals, Submitted on, Subbmitted by
      "A7:A13","B7:B13", "C7:AF13", "AG7:AG13", "AH7:AH13", "AI7:AI13",

      //_________
      //Shift
      //_________

      //Header
      "C16:AF17",

      //Till 1, Till 2, Safe
      "C18:L18", "M18:V18", "W18:AF18",

      // Till 1, Till 2, Safe Denominations, Total, Submitted on/ Submitted by
      "C19:L19", "M19:V19", "W19:AF19", "AG19", "AH19:AI19",

      // Days, Dates, DATACELLS, Totals, Submitted on, Subbmitted by
      "A20:A26", "B20:B26", "C20:AF26", "AG20:AG26", "AH20:AH26", "AI20:AI26",

      //_________
      //Close
      //_________

      //Header
      "C29:AF30",

      //Till 1, Till 2, Safe
      "C31:L31", "M31:V31", "W31:AF31",

      // Till 1, Till 2, Safe Denominations, Total, Submitted on/ Submitted by
      "C32:L32", "M32:V32", "W32:AF32", "AG32", "AH32:AI32",

      // Days, Dates, DATACELLS, Totals, Submitted on, Subbmitted by
      "A33:A39", "B33:B39", "C33:AF39", "AG33:AG39", "AH33:AH39", "AI33:AI39"],

    "MERGERIGHT": [],
    "MERGELEFT": [],
    "MERGECENTER": [
      // Open
      "C3:AF4",
      "C5:L5", "M5:V5", "W5:AF5",

      //Shift
      "C16:AF17",
      "C18:L18", "M18:V18", "W18:AF18",

      //Close
      "C29:AF30",
      "C31:L31", "M31:V31", "W31:AF31"],

    "ALLOWEDCELLS": [],
    "DATACELLS": ["C7:AF13", "C20:AF26", "C33:AF39"]},

  "WEEK_3": {
    "STRING": {
      "C3": "OPEN SHIFT",
      "C5": "TILL 1",
      "M5": "TILL 2",
      "W5": "SAFE",
      "B6": "Date",
      "AG6": "TOTAL",
      "AH6": "Submitted on",
      "AI6": "Submitted by",
      "A7": "Mon",
      "A8": "Tues",
      "A9": "Wed",
      "A10": "Thur",
      "A11": "Fri",
      "A12": "Sat",
      "A13": "Sun",

      "C16": "SHIFT CHANGE",
      "C18": "TILL 1",
      "M18": "TILL 2",
      "W18": "SAFE",
      "AG19": "TOTAL",
      "AH19": "Submitted on",
      "AI19": "Submitted by",
      "A20": "Mon",
      "A21": "Tues",
      "A22": "Wed",
      "A23": "Thur",
      "A24": "Fri",
      "A25": "Sat",
      "A26": "Sun",

      "C29": "CLOSE SHIFT",
      "C31": "TILL 1",
      "M31": "TILL 2",
      "W31": "SAFE",
      "AG32": "TOTAL",
      "AH32": "Submitted on",
      "AI32": "Submitted by",
      "A33": "Mon",
      "A34": "Tues",
      "A35": "Wed",
      "A36": "Thur",
      "A37": "Fri",
      "A38": "Sat",
      "A39": "Sun",

      // Open Currency
      "C6": "100$", "D6": "50$", "E6": "20$", "F6": "10$", "G6": "5$", "H6": "1$", "I6": ".25$", "J6": ".10$", "K6": ".05$", "L6": ".01$",  
      "M6": "100$", "N6": "50$", "O6": "20$", "P6": "10$", "Q6": "5$", "R6": "1$", "S6": ".25$", "T6": ".10$", "U6": ".05$", "V6": ".01$",
      "W6": "100$", "X6": "50$", "Y6": "20$", "Z6": "10$", "AA6": "5$", "AB6": "1$", "AC6": ".25$", "AD6": ".10$", "AE6": ".05$", "AF6": ".01$",

      // Shift Currency
      "C19": "100$", "D19": "50$", "E19": "20$", "F19": "10$", "G19": "5$", "H19": "1$", "I19": ".25$", "J19": ".10$", "K19": ".05$", "L19": ".01$",
      "M19": "100$", "N19": "50$", "O19": "20$", "P19": "10$", "Q19": "5$", "R19": "1$", "S19": ".25$", "T19": ".10$", "U19": ".05$", "V19": ".01$",
      "W19": "100$", "X19": "50$", "Y19": "20$", "Z19": "10$", "AA19": "5$", "AB19": "1$", "AC19": ".25$", "AD19": ".10$", "AE19": ".05$", "AF19": ".01$",

      // Close Currency
      "C32": "100$", "D32": "50$", "E32": "20$", "F32": "10$", "G32": "5$", "H32": "1$", "I32": ".25$", "J32": ".10$", "K32": ".05$", "L32": ".01$",
      "M32": "100$", "N32": "50$", "O32": "20$", "P32": "10$", "Q32": "5$", "R32": "1$", "S32": ".25$", "T32": ".10$", "U32": ".05$", "V32": ".01$",
      "W32": "100$", "X32": "50$", "Y32": "20$", "Z32": "10$", "AA32": "5$", "AB32": "1$", "AC32": ".25$", "AD32": ".10$", "AE32": ".05$", "AF32": ".01$"},

    "FORMULAS": {
      // Open B7:B13
      "B7": "=(WEEK_2!B13 + 1)",
      "B8": "=SUM(B7 + 1)",
      "B9": "=SUM(B8 + 1)",
      "B10": "=SUM(B9 + 1)",
      "B11": "=SUM(B10 + 1)",
      "B12": "=SUM(B11 + 1)",
      "B13": "=SUM(B12 + 1)",

      "AG7": "=SUM(C7:AF7)",
      "AG8": "=SUM(C8:AF8)",
      "AG9": "=SUM(C9:AF9)",
      "AG10": "=SUM(C10:AF10)",
      "AG11": "=SUM(C11:AF11)",
      "AG12": "=SUM(C12:AF12)",
      "AG13": "=SUM(C13:AF13)",


      // Shift B20:B16
      "B20": "=(WEEK_2!B26 + 1)",
      "B21": "=SUM(B20 + 1)",
      "B22": "=SUM(B21 + 1)",
      "B23": "=SUM(B22 + 1)",
      "B24": "=SUM(B23 + 1)",
      "B25": "=SUM(B24 + 1)",
      "B26": "=SUM(B25 + 1)",

      "AG20": "=SUM(C20:AF20)",
      "AG21": "=SUM(C21:AF21)",
      "AG22": "=SUM(C22:AF22)",
      "AG23": "=SUM(C23:AF23)",
      "AG24": "=SUM(C24:AF24)",
      "AG25": "=SUM(C25:AF25)",
      "AG26": "=SUM(C26:AF26)",


      //Close B33:B39
      "B33": "=(WEEK_2!B39 + 1)",
      "B34": "=SUM(B33 + 1)",
      "B35": "=SUM(B34 + 1)",
      "B36": "=SUM(B35 + 1)",
      "B37": "=SUM(B36 + 1)",
      "B38": "=SUM(B37 + 1)",
      "B39": "=SUM(B38 + 1)",

      "AG33": "=SUM(C33:AF33)",
      "AG34": "=SUM(C34:AF34)",
      "AG35": "=SUM(C35:AF35)",
      "AG36": "=SUM(C36:AF36)",
      "AG37": "=SUM(C37:AF37)",
      "AG38": "=SUM(C38:AF38)",
      "AG39": "=SUM(C39:AF39)"},

    "FORMAT": {
      "currency": ["C6:AF6", "C19:AF19", "C32:AF32"]},
    "COLORS": {
      "open": ["C3:AF4"],
      "shift": ["C16:AF17"],
      "close": ["C29:AF30"],
      "orange": ["C5:L5", "M5:V5", "W5:AF5", "C18:L18", "M18:V18", "W18:AF18", "C31:L31", "M31:V31", "W31:AF31"]},

    "VARIABLES": {},
    "OUTLINEMED": [
      //_________
      //Open
      //_________

      //Header
      "C3:AF4",

      //Till 1, Till 2, Safe
      "C5:L5", "M5:V5", "W5:AF5",

      // Till 1, Till 2, Safe Denominations, Total, Submitted on/ Submitted by
      "C6:L6", "M6:V6", "W6:AF6", "AG6", "AH6:AI6",

      // Days, Dates, DATACELLS, Totals, Submitted on, Subbmitted by
      "A7:A13","B7:B13", "C7:AF13", "AG7:AG13", "AH7:AH13", "AI7:AI13",

      //_________
      //Shift
      //_________

      //Header
      "C16:AF17",

      //Till 1, Till 2, Safe
      "C18:L18", "M18:V18", "W18:AF18",

      // Till 1, Till 2, Safe Denominations, Total, Submitted on/ Submitted by
      "C19:L19", "M19:V19", "W19:AF19", "AG19", "AH19:AI19",

      // Days, Dates, DATACELLS, Totals, Submitted on, Subbmitted by
      "A20:A26", "B20:B26", "C20:AF26", "AG20:AG26", "AH20:AH26", "AI20:AI26",

      //_________
      //Close
      //_________

      //Header
      "C29:AF30",

      //Till 1, Till 2, Safe
      "C31:L31", "M31:V31", "W31:AF31",

      // Till 1, Till 2, Safe Denominations, Total, Submitted on/ Submitted by
      "C32:L32", "M32:V32", "W32:AF32", "AG32", "AH32:AI32",

      // Days, Dates, DATACELLS, Totals, Submitted on, Subbmitted by
      "A33:A39", "B33:B39", "C33:AF39", "AG33:AG39", "AH33:AH39", "AI33:AI39"],

    "OUTLINETHIN": [],
    "MERGE": [],
    "RIGHT": [
      //Open Currency
      "C6", "D6", "E6", "F6", "G6", "H6", "I6", "J6", "K6", "L6",
      "M6", "N6", "O6", "P6", "Q6", "R6", "S6", "T6", "U6", "V6",
      "W6", "X6", "Y6", "Z6", "AA6", "AB6", "AC6", "AD6", "AE6", "AF6",

      // Shift Currency
      "C19", "D19", "E19", "F19", "G19", "H19", "I19", "J19", "K19", "L19",
      "M19", "N19", "O19", "P19", "Q19", "R19", "S19", "T19", "U19", "V19",
      "W19", "X19", "Y19", "Z19", "AA19", "AB19", "AC19", "AD19", "AE19", "AF19",

      //Close Currency
      "C32", "D32", "E32", "F32", "G32", "H32", "I32", "J32", "K32", "L32",
      "M32", "N32", "O32", "P32", "Q32", "R32", "S32", "T32", "U32", "V32",
      "W32", "X32", "Y32", "Z32", "AA32", "AB32", "AC32", "AD32", "AE32", "AF32"],

    "MERGERIGHT": [],
    "MERGELEFT": [],
    "MERGECENTER": [
      // Open
      "C3:AF4",
      "C5:L5", "M5:V5", "W5:AF5",

      //Shift
      "C16:AF17",
      "C18:L18", "M18:V18", "W18:AF18",

      //Close
      "C29:AF30",
      "C31:L31", "M31:V31", "W31:AF31"],

    "ALLOWEDCELLS": [],
    "DATACELLS": ["C7:AF13", "C20:AF26", "C33:AF39"]},

  "WEEK_4": {
    "STRING": {
      "C3": "OPEN SHIFT",
      "C5": "TILL 1",
      "M5": "TILL 2",
      "W5": "SAFE",
      "B6": "Date",
      "AG6": "TOTAL",
      "AH6": "Submitted on",
      "AI6": "Submitted by",
      "A7": "Mon",
      "A8": "Tues",
      "A9": "Wed",
      "A10": "Thur",
      "A11": "Fri",
      "A12": "Sat",
      "A13": "Sun",

      "C16": "SHIFT CHANGE",
      "C18": "TILL 1",
      "M18": "TILL 2",
      "W18": "SAFE",
      "AG19": "TOTAL",
      "AH19": "Submitted on",
      "AI19": "Submitted by",
      "A20": "Mon",
      "A21": "Tues",
      "A22": "Wed",
      "A23": "Thur",
      "A24": "Fri",
      "A25": "Sat",
      "A26": "Sun",

      "C29": "CLOSE SHIFT",
      "C31": "TILL 1",
      "M31": "TILL 2",
      "W31": "SAFE",
      "AG32": "TOTAL",
      "AH32": "Submitted on",
      "AI32": "Submitted by",
      "A33": "Mon",
      "A34": "Tues",
      "A35": "Wed",
      "A36": "Thur",
      "A37": "Fri",
      "A38": "Sat",
      "A39": "Sun",

      // Open Currency
      "C6": "100$", "D6": "50$", "E6": "20$", "F6": "10$", "G6": "5$", "H6": "1$", "I6": ".25$", "J6": ".10$", "K6": ".05$", "L6": ".01$",  
      "M6": "100$", "N6": "50$", "O6": "20$", "P6": "10$", "Q6": "5$", "R6": "1$", "S6": ".25$", "T6": ".10$", "U6": ".05$", "V6": ".01$",
      "W6": "100$", "X6": "50$", "Y6": "20$", "Z6": "10$", "AA6": "5$", "AB6": "1$", "AC6": ".25$", "AD6": ".10$", "AE6": ".05$", "AF6": ".01$",

      // Shift Currency
      "C19": "100$", "D19": "50$", "E19": "20$", "F19": "10$", "G19": "5$", "H19": "1$", "I19": ".25$", "J19": ".10$", "K19": ".05$", "L19": ".01$",
      "M19": "100$", "N19": "50$", "O19": "20$", "P19": "10$", "Q19": "5$", "R19": "1$", "S19": ".25$", "T19": ".10$", "U19": ".05$", "V19": ".01$",
      "W19": "100$", "X19": "50$", "Y19": "20$", "Z19": "10$", "AA19": "5$", "AB19": "1$", "AC19": ".25$", "AD19": ".10$", "AE19": ".05$", "AF19": ".01$",

      // Close Currency
      "C32": "100$", "D32": "50$", "E32": "20$", "F32": "10$", "G32": "5$", "H32": "1$", "I32": ".25$", "J32": ".10$", "K32": ".05$", "L32": ".01$",
      "M32": "100$", "N32": "50$", "O32": "20$", "P32": "10$", "Q32": "5$", "R32": "1$", "S32": ".25$", "T32": ".10$", "U32": ".05$", "V32": ".01$",
      "W32": "100$", "X32": "50$", "Y32": "20$", "Z32": "10$", "AA32": "5$", "AB32": "1$", "AC32": ".25$", "AD32": ".10$", "AE32": ".05$", "AF32": ".01$"},

    "FORMULAS": {
      // Open B7:B13
      "B7": "=(WEEK_3!B13 + 1)",
      "B8": "=SUM(B7 + 1)",
      "B9": "=SUM(B8 + 1)",
      "B10": "=SUM(B9 + 1)",
      "B11": "=SUM(B10 + 1)",
      "B12": "=SUM(B11 + 1)",
      "B13": "=SUM(B12 + 1)",

      "AG7": "=SUM(C7:AF7)",
      "AG8": "=SUM(C8:AF8)",
      "AG9": "=SUM(C9:AF9)",
      "AG10": "=SUM(C10:AF10)",
      "AG11": "=SUM(C11:AF11)",
      "AG12": "=SUM(C12:AF12)",
      "AG13": "=SUM(C13:AF13)",

      // Shift B20:B16
      "B20": "=(WEEK_3!B26 + 1)",
      "B21": "=SUM(B20 + 1)",
      "B22": "=SUM(B21 + 1)",
      "B23": "=SUM(B22 + 1)",
      "B24": "=SUM(B23 + 1)",
      "B25": "=SUM(B24 + 1)",
      "B26": "=SUM(B25 + 1)",

      "AG20": "=SUM(C20:AF20)",
      "AG21": "=SUM(C21:AF21)",
      "AG22": "=SUM(C22:AF22)",
      "AG23": "=SUM(C23:AF23)",
      "AG24": "=SUM(C24:AF24)",
      "AG25": "=SUM(C25:AF25)",
      "AG26": "=SUM(C26:AF26)",

      //Close B33:B39
      "B33": "=(WEEK_3!B39 + 1)",
      "B34": "=SUM(B33 + 1)",
      "B35": "=SUM(B34 + 1)",
      "B36": "=SUM(B35 + 1)",
      "B37": "=SUM(B36 + 1)",
      "B38": "=SUM(B37 + 1)",
      "B39": "=SUM(B38 + 1)",

      "AG33": "=SUM(C33:AF33)",
      "AG34": "=SUM(C34:AF34)",
      "AG35": "=SUM(C35:AF35)",
      "AG36": "=SUM(C36:AF36)",
      "AG37": "=SUM(C37:AF37)",
      "AG38": "=SUM(C38:AF38)",
      "AG39": "=SUM(C39:AF39)"},

    "FORMAT": {
      "currency": ["C6:AF6", "C19:AF19", "C32:AF32"]},
    "COLORS": {
      "open": ["C3:AF4"],
      "shift": ["C16:AF17"],
      "close": ["C29:AF30"],
      "orange": ["C5:L5", "M5:V5", "W5:AF5", "C18:L18", "M18:V18", "W18:AF18", "C31:L31", "M31:V31", "W31:AF31"]},

    "VARIABLES": {},
    "OUTLINEMED": [
      //_________
      //Open
      //_________

      //Header
      "C3:AF4",

      //Till 1, Till 2, Safe
      "C5:L5", "M5:V5", "W5:AF5",

      // Till 1, Till 2, Safe Denominations, Total, Submitted on/ Submitted by
      "C6:L6", "M6:V6", "W6:AF6", "AG6", "AH6:AI6",

      // Days, Dates, DATACELLS, Totals, Submitted on, Subbmitted by
      "A7:A13","B7:B13", "C7:AF13", "AG7:AG13", "AH7:AH13", "AI7:AI13",

      //_________
      //Shift
      //_________

      //Header
      "C16:AF17",

      //Till 1, Till 2, Safe
      "C18:L18", "M18:V18", "W18:AF18",

      // Till 1, Till 2, Safe Denominations, Total, Submitted on/ Submitted by
      "C19:L19", "M19:V19", "W19:AF19", "AG19", "AH19:AI19",

      // Days, Dates, DATACELLS, Totals, Submitted on, Subbmitted by
      "A20:A26", "B20:B26", "C20:AF26", "AG20:AG26", "AH20:AH26", "AI20:AI26",

      //_________
      //Close
      //_________

      //Header
      "C29:AF30",

      //Till 1, Till 2, Safe
      "C31:L31", "M31:V31", "W31:AF31",

      // Till 1, Till 2, Safe Denominations, Total, Submitted on/ Submitted by
      "C32:L32", "M32:V32", "W32:AF32", "AG32", "AH32:AI32",

      // Days, Dates, DATACELLS, Totals, Submitted on, Subbmitted by
      "A33:A39", "B33:B39", "C33:AF39", "AG33:AG39", "AH33:AH39", "AI33:AI39"],

    "OUTLINETHIN": [],
    "MERGE": [],
    "RIGHT": [
      //Open Currency
      "C6", "D6", "E6", "F6", "G6", "H6", "I6", "J6", "K6", "L6",
      "M6", "N6", "O6", "P6", "Q6", "R6", "S6", "T6", "U6", "V6",
      "W6", "X6", "Y6", "Z6", "AA6", "AB6", "AC6", "AD6", "AE6", "AF6",

      // Shift Currency
      "C19", "D19", "E19", "F19", "G19", "H19", "I19", "J19", "K19", "L19",
      "M19", "N19", "O19", "P19", "Q19", "R19", "S19", "T19", "U19", "V19",
      "W19", "X19", "Y19", "Z19", "AA19", "AB19", "AC19", "AD19", "AE19", "AF19",

      //Close Currency
      "C32", "D32", "E32", "F32", "G32", "H32", "I32", "J32", "K32", "L32",
      "M32", "N32", "O32", "P32", "Q32", "R32", "S32", "T32", "U32", "V32",
      "W32", "X32", "Y32", "Z32", "AA32", "AB32", "AC32", "AD32", "AE32", "AF32"],

    "MERGERIGHT": [],
    "MERGELEFT": [],
    "MERGECENTER": [
      // Open
      "C3:AF4",
      "C5:L5", "M5:V5", "W5:AF5",

      //Shift
      "C16:AF17",
      "C18:L18", "M18:V18", "W18:AF18",

      //Close
      "C29:AF30",
      "C31:L31", "M31:V31", "W31:AF31"],

    "ALLOWEDCELLS": [],
    "DATACELLS": ["C7:AF13", "C20:AF26", "C33:AF39"]},

  "INFO": {
    "STRING": {
      "B2": "Period number:",
      "B3": "First Day of Period:",
      "B5": "Week range on cover:"},
    "FORMULAS": {},
    "COLORS": {
      "splash": [],
      "white": []},
    "VARIABLES": {},
    "OUTLINEMED": [],
    "OUTLINETHIN": [],
    "MERGE": [],
    "RIGHT": ["B2", "B3", "B5"],
    "MERGERIGHT": [],
    "MERGELEFT": [],
    "MERGECENTER": [],
    "ALLOWEDCELLS": [],
    "DATACELLS": []}
};

// Dictionary of color names and corresponding hex codes
const ColorDict = {
  "open": "#6ac16b",
  "shift": "#a64d79",
  "close": "#4285f4",
  "search": "#f1cb65",
  "orange": "#FFA500",
  "light blue": "#ADD8E6",
  "red": "#FF0000",
  "yellow": "#FFFF00",
  "purple": "#800080",
  "pink": "#FFC0CB",
  "brown": "#A52A2A",
  "black": "#000000",
  "white": "#ffffff"
};

const TempCellBackup = {};

const REQUIRED_SHEETS = [
  "SPLASH",
  "COVER",
  "OPEN",
  "SHIFT",
  "CLOSE",
  "SEARCH",
  "WEEK_1",
  "WEEK_2",
  "WEEK_3",
  "WEEK_4",
  "INFO"
];

const CurrentShiftIndicator = "N2";
const LAST_LIVE_SHIFT_PROPERTY = "LAST_LIVE_SHIFT";
const VALID_LIVE_SHIFTS = ["OPEN", "SHIFT", "CLOSE"];
const VALID_VIEW_MODES = ["OPEN", "SHIFT", "CLOSE", "SEARCH"];

let CurrentDate = new Date();
let CurrentWeek = null;
//----------------------------------------------------------------------------------------------------
// FUNCTIONS FOR OPENING THE FILE
//----------------------------------------------------------------------------------------------------

// Runs when the spreadsheet is opened
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const splashSheet = ss.getSheetByName("SPLASH");

  BuildMenus();

  if (!IsFileInitialized()) {
    BuildSetupMenu();
  }

  if (!splashSheet) {
    Logger.log("SPLASH sheet not found.");
    return;
  }

  // Force splash first
  splashSheet.showSheet();
  splashSheet.activate();

  ss.getSheets().forEach(sheet => {
    if (sheet.getName() !== "SPLASH") {
      sheet.hideSheet();
    }
  });

  // Then rebuild file state
  GetWeekInfo();
  // Restore();

  SpreadsheetApp.getUi().alert(
  "Welcome to the Dominos Till Count Sheet!\nPlease enter Emp Id and click the button to begin."
);

  
}

// Transition from splash screen to cover sheet
function StartApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const splashSheet = ss.getSheetByName("SPLASH");
  const coverSheet = ss.getSheetByName("COVER");

  if (!splashSheet || !coverSheet) {
    Logger.log("SPLASH or COVER sheet not found.");
    return;
  }

  const employeeId = splashSheet.getRange("K13").getValue().toString().trim();

  if (!employeeId) {
    SpreadsheetApp.getUi().alert("Please enter an employee ID before continuing.");
    splashSheet.activate();
    return;
  }

  Restore();

  coverSheet.showSheet();
  coverSheet.activate();
  splashSheet.hideSheet();

  // Keep all other sheets hidden during normal use
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name !== "COVER" && name !== "SPLASH") {
      sheet.hideSheet();
    }
  });

  coverSheet.getRange("D24").setValue(employeeId);
  splashSheet.getRange("K13").clearContent();

  // Resume the last live shift, never SEARCH, when entering from splash
  const lastLiveShift = GetLastLiveShift();
  coverSheet.getRange(CurrentShiftIndicator).setValue(lastLiveShift);

  UpdateCoverSheet();

  SpreadsheetApp.getUi().alert("EMPLOYEE ID : " + employeeId);
}

function BuildMenus() {
  const props = PropertiesService.getDocumentProperties();
  const isLocked = props.getProperty("docLocked") === "true";

  const security = SpreadsheetApp.getUi().createMenu("🔒 Security");

  if (isLocked) {
    security.addItem("🔓 Unlock Document", "Unlock");
  } else {
    security.addItem("🔐 Lock Document", "Lock");
  }

  security
    .addSeparator()
    .addItem("🛠 Set/Change Password", "SetupPassword")
    .addItem("❓ Recover Password", "RecoverPassword")
    .addToUi();

  const tools = SpreadsheetApp.getUi().createMenu("🛠 Script Tools");

  if (!isLocked) {
    tools
      .addItem("Set New Version", "PromptNewVersion")
      .addSeparator();
  }

  tools
    .addItem("Show Current Version", "ShowVersion")
    .addItem("View Changelog", "ViewChangeLog")
    .addToUi();
}

function IsFileInitialized() {
  const props = PropertiesService.getDocumentProperties();
  return props.getProperty("FILE_INITIALIZED") === "true";
}

function BuildSetupMenu() {
  SpreadsheetApp.getUi().createMenu("⚙️ Setup")
    .addItem("Initialize This File", "InitializeThisFile")
    .addToUi();
}

function InitializeThisFile() {
  EnsureRequiredSheetsExist();

  SetupPassword();
  ApplySheetProtections();
  InstallChangeTrigger();

  GetWeekInfo();
  Restore();
  UpdateCoverSheet();

  const props = PropertiesService.getDocumentProperties();
  props.setProperty("FILE_INITIALIZED", "true");

  SpreadsheetApp.getUi().alert("Setup complete.");

  // Do not call onOpen() here.
  // Just return the file to splash mode directly.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const splashSheet = ss.getSheetByName("SPLASH");

  if (splashSheet) {
    splashSheet.showSheet();
    splashSheet.activate();
  }

  ss.getSheets().forEach(sheet => {
    if (sheet.getName() !== "SPLASH") {
      sheet.hideSheet();
    }
  });
}

function EnsureRequiredSheetsExist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingNames = ss.getSheets().map(s => s.getName());

  REQUIRED_SHEETS.forEach(name => {
    if (!existingNames.includes(name)) {
      ss.insertSheet(name);
    }
  });
}

function GetLastLiveShift() {
  const props = PropertiesService.getDocumentProperties();
  const stored = (props.getProperty(LAST_LIVE_SHIFT_PROPERTY) || "").toUpperCase().trim();
  return VALID_LIVE_SHIFTS.includes(stored) ? stored : "OPEN";
}

function SetLastLiveShift(shiftType) {
  const props = PropertiesService.getDocumentProperties();
  const normalized = (shiftType || "").toUpperCase().trim();

  if (VALID_LIVE_SHIFTS.includes(normalized)) {
    props.setProperty(LAST_LIVE_SHIFT_PROPERTY, normalized);
  }
}

function GetCurrentViewMode() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coverSheet = ss.getSheetByName("COVER");
  if (!coverSheet) return "OPEN";

  const mode = coverSheet.getRange(CurrentShiftIndicator).getValue().toString().trim().toUpperCase();
  return VALID_VIEW_MODES.includes(mode) ? mode : "OPEN";
}
