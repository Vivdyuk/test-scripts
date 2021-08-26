function myFunction(doc, range) {

    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = spreadSheet.getSheetByName('Operating Budget');
    const activeRange = activeSheet.getActiveRange();

    if (activeRange.getWidth() != 7) {
        throw new Error('range has to be 7 columns wide')
    }

    const sqlSheet = spreadSheet.getSheetByName('SQL Staff');

    if (!sqlSheet) {
        throw new Error('There\'s no SQL sheet');
    }

    const firstRow = activeRange.getRow();
    const lastRow = activeRange.getLastRow();
    const formulas = _.times((lastRow - firstRow + 1), n => {
        if (activeSheet.getRange(`B${firstRow + n}`).getValue().startsWith('Total')) {
            return [
                `=SUM(C${firstRow}:C${lastRow - 2})`,
                '',
                `=SUM(E${firstRow}:E${lastRow - 2})`,
                '',
                `=SUM(G${firstRow}:G${lastRow - 2})`,
                '',
                `=SUM(I${firstRow}:I${lastRow - 2})`
            ]
        }

        if (activeSheet.getRange(`B${firstRow +n}`).getValue().startsWith('TOTAL')) {
            return [
                `=C${lastRow - 1}`,
                '',
                `=E${lastRow - 1}`,
                '',
                `=G${lastRow - 1}`,
                '',
                `=I${lastRow - 1}`
            ]
        }
        return (
            [`=SUMIF('SQL Staff'!$G:$G,$A${firstRow + n},'SQL Staff'!A:A)`,
                ``,
                `=SUMIF('SQL Staff'!$G:$G,$A${firstRow + n},'SQL Staff'!B:B)`,
                ``,
                `=SUMIF('SQL Staff'!$G:$G,$A${firstRow + n},'SQL Staff'!C:C)`,
                ``,
                `=SUMIF('SQL Staff'!$G:$G,B${firstRow + n},'SQL Staff'!D:D)`
            ]
        )
    })

    activeRange.setFormulas(formulas);
}
