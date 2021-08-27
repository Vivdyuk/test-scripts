function myFunction(doc, range) {
    var _ = Underscore.load();

    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = spreadSheet.getSheetByName('Operating Budget');
    const activeRange = activeSheet.getRange('B22:I45');
    const neededWidth = 8;
    const getTheValue = (letter, number) => activeSheet.getRange(`${letter}${number}`).getValue();


    if (activeRange.getWidth() != neededWidth) {
        throw new Error(`'choose a correct range in a sheet, not ${activeRange.getA1Notation()}'`);
    }

    const sqlSheet = spreadSheet.getSheetByName('SQL');

    if (!sqlSheet) {
        throw new Error('There\'s no SQL sheet');
    }

    const firstRow = activeRange.getRow();
    const lastRow = activeRange.getLastRow();
    const rangeHeight = lastRow - firstRow + 1;
    let startRow = firstRow;
    const filler = new Array(neededWidth).fill('');

    const formulas = _.times((rangeHeight), n => {
        const currentCell = firstRow + n;
        const cellCheckValue = getTheValue('A', currentCell);

        if (cellCheckValue === '') {
            startRow = currentCell + 1;
            const result = [...filler];
            result.splice(0, 1, `"${getTheValue('B', currentCell)}"`);

            return result;
        }

        if (cellCheckValue === 'un-subtotal') {
            return [
                `"${activeSheet.getRange(currentCell, activeRange.getColumn()).getValue()}"`,
                `=SUM(C${startRow}:C${currentCell - 1})`,
                '',
                `=SUM(E${startRow}:E${currentCell - 1})`,
                '',
                `=SUM(G${startRow}:G${currentCell - 1})`,
                '',
                `=SUM(I${startRow}:I${currentCell - 1})`
            ]
        }

        return [
            `="  "&VLOOKUP(A${currentCell},SQL!G:M,7,false)`,
            `=SUMIFS(SQL!$A:$A,SQL!$E:$E,"UNRESTRICTED",SQL!$G:$G,A${currentCell})`,
            '',
            `=SUMIFS(SQL!$B:$B,SQL!$E:$E,"UNRESTRICTED",SQL!$G:$G,A${currentCell})`,
            '',
            `=SUMIFS(SQL!$C:$C,SQL!$E:$E,"UNRESTRICTED",SQL!$G:$G,A${currentCell})`,
            '',
            `=SUMIFS(SQL!$D:$D,SQL!$E:$E,"UNRESTRICTED",SQL!$G:$G,A${currentCell})`
        ]
    })

    const previousRangeInfo = activeRange.getValues();

    const diffRange = activeSheet.getRange(
        firstRow,
        activeRange.getColumn() + 1 + neededWidth,
        rangeHeight,
        neededWidth
    );

    activeRange.copyFormatToRange(
        diffRange.getGridId(),
        activeRange.getColumn() + 1 + neededWidth,
        activeRange.getColumn() + 1 + neededWidth + neededWidth,
        firstRow,
        rangeHeight
    )

    diffRange.setValues(previousRangeInfo)
    activeRange.setFormulas(formulas);
    const actualValues = activeRange.getValues();

    const compareColours = _.map(previousRangeInfo, (row, rowIndex) =>
        _.map(row, (cell, cellIndex) =>
            cell === actualValues[rowIndex][cellIndex] ? '#fff' : '#f00'
        )
    );

    diffRange.setBackgrounds(compareColours);
}
