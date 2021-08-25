function myFunction(doc, range) {

    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = spreadSheet.getSheetByName('Operating Budget');
    const activeRange = activeSheet.getActiveRange();

    if (activeRange.getWidth() != 7 || activeRange.getA1Notation() !== 'C6:I15') {
        throw new Error(`'choose a correct range in a sheet, not ${activeRange.getA1Notation()}'`)
    }

    const sqlSheet = spreadSheet.getSheetByName('SQL Staff');

    if (!sqlSheet) {
        throw new Error('There\'s no SQL sheet');
    }

    const firstRow = activeRange.getRow();
    const lastRow = activeRange.getLastRow();
    const formulas = [];
    let iteratee = 0;

    while(formulas.length < 8) {
        formulas.push(
            [`=SUMIF('SQL Staff'!$G:$G,$A${firstRow + iteratee},'SQL Staff'!A:A)`,
                ``,
                `=SUMIF('SQL Staff'!$G:$G,$A${firstRow + iteratee},'SQL Staff'!B:B)`,
                ``,
                `=SUMIF('SQL Staff'!$G:$G,$A${firstRow + iteratee},'SQL Staff'!C:C)`,
                ``,
                `=SUMIF('SQL Staff'!$G:$G,B${firstRow + iteratee},'SQL Staff'!D:D)`
            ]
        );

        iteratee++;
    }

    formulas.push(
        [
            `=SUM(C${firstRow - 1}:C${lastRow - 2})`,
            '',
            `=SUM(E${firstRow - 1}:E${lastRow - 2})`,
            '',
            `=SUM(G${firstRow - 1}:G${lastRow - 2})`,
            '',
            `=SUM(I${firstRow - 1}:I${lastRow - 2})`
        ],
        [
            `=C${lastRow - 1}`,
            '',
            `=E${lastRow - 1}`,
            '',
            `=G${lastRow - 1}`,
            '',
            `=I${lastRow - 1}`,
        ]
    )

    activeRange.setFormulas(formulas);
}
