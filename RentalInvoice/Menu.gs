function onOpen() {
    var menuEntries = [{
            name: "Generate (Single) Invoice for This Month",
            functionName: "createThisMonthInvoice"
        }, {
            name: "Generate (Single) Invoice for Last Month",
            functionName: "createLastMonthInvoice"
        },{
            name: "Generate (Single) Invoice for This Month (Override red lines on invoice)",
            functionName: "createThisMonthInvoiceOverride"
        }, {
            name: "Generate (Single) Invoice for Last Month (Override red lines on invoice)",
            functionName: "createLastMonthInvoiceOverride"
        },{
            name: "Generate (Single) Invoice for 2 Month Ago (Override red lines on invoice)",
            functionName: "createTwoMonthsAgeInvoiceOverride"
        }, {
            name: "Generate (All) Invoices for This Month",
            functionName: "createAllInvoicesCurrentMonth"
        }, {
            name: "Generate (All) Invoices for Last Month",
            functionName: "createAllInvoicesLastMonth"
        }, {
            name: "Email Me Totals for (All) Invoices for This Month",
            functionName: "createAllInvoicesCurrentMonthStatsOnly"
        }, {
            name: "Email Me Totals for (All) Invoices for Last Month",
            functionName: "createAllInvoicesLastMonthStatsOnly"
        },{
            name: "Go to the Last Row",
            functionName: "ToBottom"
        },{
            name: "Move Single Tab to the Next Alphabetical Tab",
            functionName: "sortGoogleSheets"
        }
    ];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.addMenu("Jonathan's Magic", menuEntries);
}

function createAllInvoicesCurrentMonth() {
    processAllInvoices("Current", false, true);
}

function createAllInvoicesLastMonth() {
    processAllInvoices("Last", false, true);
}

// function processAllInvoices(invoiceMonth, printPDF, emailReport)

function createAllInvoicesCurrentMonthStatsOnly() {
    processAllInvoices("Current", false, false);
}

function createAllInvoicesLastMonthStatsOnly() {
    processAllInvoices("Last", false, false);
}




function createAllInvoicesAndPrintOn28th() {
    processAllInvoices("Current", true, true);
}

function createThisMonthInvoice() {
    createInvoice("Current", null, false, true);
}

function createLastMonthInvoice() {
    createInvoice("Last", null, false, true);
}


// function processAllInvoices(invoiceMonth, printPDF, emailReport)
function createThisMonthInvoiceOverride() {
    createInvoice("Current", null, false, true, true);
}

function createLastMonthInvoiceOverride() {
    createInvoice("Last", null, false, true, true);
}

function createTwoMonthsAgeInvoiceOverride() {
    createInvoice("twoMonthsAgo", null, false, true, true);
}



