let ivySheet, markitSheet;
const dateColumns = ['Maturity date', 'End Date' , 'Death', 'Trade Date', 'Settle Date', 'Accrue Date', 'Additional Payment 1 Date', 'Fixed Start Date'];

document.getElementById('loadBtn').addEventListener('click', handleFileLoad);
document.getElementById('compareBtn').addEventListener('click', handleFileCompare);

const mapping = [
    { "comparison": "Trade date", "ivy": "Trade Date", "markit": "Trade Date" },
    { "comparison": "Settle date", "ivy": "Settle Date", "markit": "Additional Payment 1 Date" }, //need to check from Jiju
    { "comparison": "Accrue date", "ivy": "Accrue Date", "markit": "Fixed Start Date" },
    { "comparison": "Maturity date", "ivy": "Death", "markit": "End Date" },
    { "comparison": "Direction - pay/recv fixed", "ivy": "direction", "markit": "Direction" },
    { "comparison": "Quantity/notional", "ivy": "Quantity", "markit": "Notional" },
    { "comparison": "Payment Frequency - leg 1", "ivy": "Swap Frequency Leg 1", "markit": "Fixed Payment Freq" },
    { "comparison": "Payment Frequency - leg 2", "ivy": "Swap Frequency Leg 2", "markit": "Float Payment Freq" },
    { "comparison": "Reset frequency – Float leg", "ivy": "Reset Frequency", "markit": "Float Reset Freq" },
    { "comparison": "Index Name", "ivy": "RATE SOURCE2", "markit": "F/Rate Index" },
    { "comparison": "Swap Level (Fix rate/100)", "ivy": "Swap Level", "markit": "KEY (Don't touch)" },
    { "comparison": "Settlement Amount", "ivy": "Net Money", "markit": "Additional Payment 1 Amount" },
    { "comparison": "Settlement Direction", "ivy": "Net Money", "markit": "Additional Payment 1 Direction" },
    { "comparison": "Settlement Currency", "ivy": "Settle Ccy", "markit": "Brokerage Currency" },
    { "comparison": "Trade ID", "ivy": "Trade ID", "markit": "Trade ID" },
    { "comparison": "Daycount Leg 1", "ivy": "Daycount Type Leg 1", "markit": "Fixed Day Basis" },
    { "comparison": "Daycount Leg 2", "ivy": "Daycount Type Leg 2", "markit": "Float Day Basis" },
    { "comparison": "Roll Convention", "ivy": "Roll Convention", "markit": "Roll Day" },
    { "comparison": "IRS Type", "ivy": "ir_swap_type", "markit": "Product" },
    { "comparison": "Contra Broker", "ivy": "Contra Broker", "markit": "Broker Code" }
  ];

  const reviewRequiredByUser = ['Index Name', 'Roll Convention'];

function handleFileLoad() {
    const fileUpload = document.getElementById('fileUpload').files[0];
    if (!fileUpload) {
        alert('Please upload an Excel file');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        ivySheet = XLSX.utils.sheet_to_json(workbook.Sheets['Ivy'], { 
            header: 1, 
            raw: false,
            dateNF: 'dd/mm/yyyy',
            cellDates: true,
            rawDates: false,
            defval: '',
            parseCells: parseCell
        });
        markitSheet = XLSX.utils.sheet_to_json(workbook.Sheets['Markit'], { 
            header: 1, 
            raw: false,
            dateNF: 'dd/mm/yyyy',
            cellDates: true,
            rawDates: false,
            defval: '',
            parseCells: parseCell
        });

        if (!ivySheet || !markitSheet) {
            alert('Both sheets Ivy and Markit must be present');
            return;
        }

        populateSwapLevelDropdown(ivySheet);
    };
    reader.readAsArrayBuffer(fileUpload);
}

function parseCell(cell, row, col, sheetName) {
    const columnName = sheetName === 'Ivy' ? ivySheet[0][col] : markitSheet[0][col];
    if (dateColumns.includes(columnName)) {
        if (typeof cell === 'number') {
            // Convert Excel serial date to JS Date object
            alert(1)
            return new Date((cell - 25569) * 86400 * 1000);
        } else if (typeof cell === 'string') {
            // If it's already a string, try to parse it as a date
            const parsedDate = new Date(cell).toLocaleDateString('en-GB');
            console.log(parsedDate)
            return isNaN(parsedDate.getTime()) ? cell : parsedDate;
        }
    }
    return cell;
}

function populateSwapLevelDropdown(ivySheet) {
    const swapLevelIndex = ivySheet[0].indexOf('Swap Level');
    if (swapLevelIndex === -1) {
        alert('Swap Level column not found in Ivy sheet');
        return;
    }

    const swapLevelDropdown = document.getElementById('swapLevelDropdown');
    swapLevelDropdown.innerHTML = ''; // Clear existing options
    const swapLevels = new Set();

    for (let i = 1; i < ivySheet.length; i++) {
        swapLevels.add(ivySheet[i][swapLevelIndex]);
    }

    swapLevels.forEach(level => {
        const option = document.createElement('option');
        option.value = level;
        option.text = level;
        swapLevelDropdown.add(option);
    });

    swapLevelDropdown.style.display = 'block';
    document.getElementById('compareBtn').style.display = 'inline';
}

function handleFileCompare() {
    const selectedSwapLevel = document.getElementById('swapLevelDropdown').value;
    if (!selectedSwapLevel) {
        alert('Please select a Swap Level');
        return;
    }

    const comparisonSheet = compareSheets(ivySheet, markitSheet, Number(selectedSwapLevel));

    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.aoa_to_sheet(comparisonSheet);
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Comparison');

    const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });

    const link = document.getElementById('downloadLink');
    link.href = URL.createObjectURL(blob);
    link.download = 'Comparison.xlsx';
    link.style.display = 'block';
    link.textContent = 'Download Comparison';
}

function compareSheets(ivySheet, markitSheet, selectedSwapLevel) {
    const comparisonSheet = [['Column Name', 'Ivy Value', 'Markit Value', 'Status']];
    const ivyHeader = ivySheet[0];
    const markitHeader = markitSheet[0];

    const swapLevelIndexIvy = ivyHeader.indexOf('Swap Level');
    const swapLevelIndexMarkit = markitHeader.indexOf("KEY (Don't touch)");

    for (let i = 1; i < ivySheet.length; i++) {
        if (Number(ivySheet[i][swapLevelIndexIvy]) === selectedSwapLevel) {
            const markitRow = markitSheet.find(row => Number(row[swapLevelIndexMarkit]) === Number(selectedSwapLevel));

            if (!markitRow) continue;

            mapping.forEach(mappingItem => {
                const ivyIndex = ivyHeader.indexOf(mappingItem.ivy);
                const markitIndex = markitHeader.indexOf(mappingItem.markit);

                let ivyValue, markitValue;
                if (dateColumns.includes(mappingItem.ivy)) {
                    ivyValue = new Date(ivySheet[i][ivyIndex]).toLocaleDateString('en-GB');
                } else if (mappingItem.ivy === 'Quantity') {
                    ivyValue = Math.abs(ivySheet[i][ivyIndex]);
                } else if (mappingItem.ivy === 'Net Money') {
                    ivyValue = ivySheet[i][ivyIndex] > 0 ? 'Rec' : 'Pay';
                } else {
                    ivyValue = ivySheet[i][ivyIndex];
                }
                
                if (['Fixed Payment Freq', 'Float Payment Freq'].includes(mappingItem.markit)) {
                    thenum = markitRow[markitIndex].match(/\d+/)[0] // "3"
                    markitValue = thenum;
                } else if (mappingItem.markit === 'Notional') {
                    markitValue = Number(markitRow[markitIndex]);
                } else {
                    markitValue = markitRow[markitIndex];
                }

                let status = 'Matched';
                if (['Daycount Leg 1', 'Daycount Leg 2'].includes(mappingItem.comparison) && (ivyValue.includes(markitValue) || markitValue.includes(ivyValue))) {

                } else if (reviewRequiredByUser.includes(mappingItem.comparison)) {
                    status = 'Review required by user';
                } else if (mappingItem.comparison === 'Swap Level (Fix rate/100)') {

                } else if (ivyValue !== markitValue) {
                    status = 'Unmatched';
                }

                comparisonSheet.push([mappingItem.comparison, ivyValue, markitValue, status]);
            });
            break; // Assuming only one row per Swap Level, remove this break if multiple rows per level
        }
    }

    return comparisonSheet;
}
