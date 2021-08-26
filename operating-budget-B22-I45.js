function myFunction(doc, range) {
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = spreadSheet.getSheetByName('Operating Budget');
    const activeRange = activeSheet.getActiveRange();

    if (activeRange.getA1Notation() !== 'B22:I45') {
        throw new Error(`'choose a correct range in a sheet, not ${activeRange.getA1Notation()}'`)
    }

    const sqlsheet = spreadSheet.getSheetByName('SQL');

    if(!sqlsheet) {
        throw new Error ('There\'s no SQL sheet');
    }

    const  fillPartsOfRange = (start, end) => {
        const result = [];
        let n = 0;

        while ((start + n) < end) {
            result.push([
                `="  "&VLOOKUP(A${start + n},SQL!G:M,7,false)`,
                `=SUMIFS(SQL!$A:$A,SQL!$E:$E,"UNRESTRICTED",SQL!$G:$G,A${start + n})`,
                '',
                `=SUMIFS(SQL!$B:$B,SQL!$E:$E,"UNRESTRICTED",SQL!$G:$G,A${start + n})`,
                '',
                `=SUMIFS(SQL!$C:$C,SQL!$E:$E,"UNRESTRICTED",SQL!$G:$G,A${start + n})`,
                '',
                `=SUMIFS(SQL!$D:$D,SQL!$E:$E,"UNRESTRICTED",SQL!$G:$G,A${start + n})`
            ]);

            n++;
        }

        result.push([
            '',
            `=SUM(C${start}:C${end - 1})`,
            '',
            `=SUM(E${start}:E${end - 1})`,
            '',
            `=SUM(G${start}:G${end - 1})`,
            '',
            `=SUM(I${start}:I${end - 1})`
        ])

        return result;
    }

    const firstPart = fillPartsOfRange(22, 30);   // 22-30 , 9
    const secondPart = fillPartsOfRange(32, 37);  // 32-37, 6
    const thirdPart = fillPartsOfRange(39, 45);   // 39-45, 7
    const filler = new Array(8).fill('');

    const fullFormulas = [].concat(firstPart, [filler], secondPart, [filler],thirdPart);

    activeRange.setFormulas(fullFormulas);
}
