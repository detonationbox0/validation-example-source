/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Ready.")
    }
});

$("#setup").on("click", () => tryCatch(setup));
$("#require-approved-name").on("click", () => tryCatch(requireApprovedName));
$("#bind-table-event").on("click", () => tryCatch(bindTable));

async function bindTable() {
    await Excel.run(async (context) => {

        let table = context.workbook.tables.getItem("NameOptionsTable");
        table.onChanged.add(onChange);

        await context.sync();
        console.log("A handler has been registered for the onChanged event");

    })
}

async function onChange(event) {
    await Excel.run(async (context) => {
        console.log("Handler for table onChanged event has been triggered. Data changed address:" + event.address)
    })
}

async function requireApprovedName() {
    await Excel.run(async (context) => {

        const sheet = context.workbook.worksheets.getItem("Decision");
        const nameRange = sheet.tables
            .getItem("NameOptionsTable")
            .columns.getItem("Baby Name")
            .getDataBodyRange();

        // When you are developing, it is a good practice to
        // clear the dataValidation object with each run of your code.
        nameRange.dataValidation.clear();

        const nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

        let approvedListRule = {
            list: {
                inCellDropDown: true,
                source: nameSourceRange
            }
        };

        nameRange.dataValidation.rule = approvedListRule;

        await context.sync();
    });
}

async function setup() {
    await Excel.run(async (context) => {
        context.workbook.worksheets.getItemOrNullObject("Decision").delete();
        const decisionSheet = context.workbook.worksheets.add("Decision");

        const optionsTable = decisionSheet.tables.add("A1:C4", true /*hasHeaders*/);
        optionsTable.name = "NameOptionsTable";
        optionsTable.showBandedRows = false;

        optionsTable.getHeaderRowRange().values = [["Baby Name", "Ranking", "Comments"]];

        decisionSheet.getUsedRange().format.autofitColumns();
        decisionSheet.getUsedRange().format.autofitRows();

        // The names that will be allowed in the Baby Name column are
        // listed in a range on the Names sheet.
        context.workbook.worksheets.getItemOrNullObject("Names").delete();
        const namesSheet = context.workbook.worksheets.add("Names");

        namesSheet.getRange("A1:A3").values = [["Sue"], ["Ricky"], ["Liz"]];
        decisionSheet.activate();
        await context.sync();
    });
}



/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}
