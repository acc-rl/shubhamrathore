let ivySheet, markitSheet;
const dateColumns = ['Maturity date', 'End Date', 'Death', 'Trade Date', 'Settle Date', 'Accrue Date', 'Additional Payment 1 Date', 'Fixed Start Date'];

document.getElementById('compareBtn').addEventListener('click', handleFileLoad);

const PaymentFrequencyLeg1Mapping = new Map([
    [1, '12M'],
    [2, '6M'],
    [4, '3M'],
    [12, '1M'],
    [6, '2M'],
    [52, '7D']
]);
const contraBroker = {
    BBG: ['1769597', '1769481'],
    TWEPCCP: ['1769598', '1769480']
}


const mapping = [
    { "comparison": "Trade date", "ivy": "Trade Date", "markit": "Trade Date" },
    { "comparison": "Settle date", "ivy": "Settle Date", "markit": "Additional Payment 1 Date" },
    { "comparison": "Accrue date", "ivy": "Accrue Date", "markit": "Fixed Start Date" },
    { "comparison": "Maturity date", "ivy": "Death", "markit": "End Date" },
    { "comparison": "Direction - pay/recv fixed", "ivy": "direction", "markit": "Direction" },
    { "comparison": "Quantity/notional", "ivy": "Quantity", "markit": "Notional" },
    { "comparison": "Payment Frequency - leg 1", "ivy": "Swap Frequency Leg 1", "markit": "Fixed Payment Freq" },
    { "comparison": "Payment Frequency - leg 2", "ivy": "Swap Frequency Leg 2", "markit": "Float Payment Freq" },
    { "comparison": "Reset frequency – Float leg", "ivy": "Reset Frequency", "markit": "Float Reset Freq" },
    { "comparison": "Index Name", "ivy": "RATE SOURCE2", "markit": "F/Rate Index" },
    // { "comparison": "Swap Level (Fix rate/100)", "ivy": "Swap Level", "markit": "KEY (Don't touch)" },
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
    reader.onload = function (e) {
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
        handleFileCompare()
    };
    reader.readAsArrayBuffer(fileUpload);
}

function parseCell(cell, row, col, sheetName) {
    const columnName = sheetName === 'Ivy' ? ivySheet[0][col] : markitSheet[0][col];
    if (dateColumns.includes(columnName)) {
        if (typeof cell === 'number') {
            // Convert Excel serial date to JS Date object
            return new Date((cell - 25569) * 86400 * 1000);
        } else if (typeof cell === 'string') {
            // If it's already a string, try to parse it as a date
            const parsedDate = new Date(cell).toLocaleDateString('en-GB');
            return isNaN(parsedDate.getTime()) ? cell : parsedDate;
        }
    }
    return cell;
}

function getAllSwapLevels(ivySheet) {
    const swapLevelIndex = ivySheet[0].indexOf('Swap Level');
    if (swapLevelIndex === -1) {
        alert('Swap Level column not found in Ivy sheet');
        return;
    }

    const swapLevels = new Set();

    for (let i = 1; i < ivySheet.length; i++) {
        if (ivySheet[i][swapLevelIndex]) {
        swapLevels.add(ivySheet[i][swapLevelIndex]);
        }
    }
    return swapLevels;
}

function handleFileCompare() {
    const newWorkbook = XLSX.utils.book_new();
    const allSwapLevels = getAllSwapLevels(ivySheet);

    allSwapLevels.forEach(level => {
        let comparisonSheet = compareSheets(ivySheet, markitSheet, level);
        const newSheet = XLSX.utils.aoa_to_sheet(comparisonSheet);

        applyConditionalFormatting(newSheet, comparisonSheet);

        XLSX.utils.book_append_sheet(newWorkbook, newSheet, `Comparison ${level}`);
    });

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
        if (Number(ivySheet[i][swapLevelIndexIvy]) === Number(selectedSwapLevel)) {
            const markitRow = markitSheet.find(row => Number(row[swapLevelIndexMarkit]) === Number(selectedSwapLevel));
            if (!markitRow) continue;

            mapping.forEach(mappingItem => {
                const ivyIndex = ivyHeader.indexOf(mappingItem.ivy);
                const markitIndex = markitHeader.indexOf(mappingItem.markit);
                let status = 'Matched';

                let ivyValue = ivySheet[i][ivyIndex];
                let markitValue = markitRow[markitIndex];

                if (ivyIndex === -1 || markitIndex === -1) {
                    status = `ERROR - Column not found in ${ivyIndex === -1 ? 'Ivy' : 'Markit'}`;
                    comparisonSheet.push([mappingItem.comparison, '-', '-', status]);
                    return;
                }

                if (dateColumns.includes(mappingItem.ivy)) {
                    ivyValue = new Date(ivyValue).toLocaleDateString('en-GB');
                    markitValue = markitValue;
                }

                if (mappingItem.ivy === 'Net Money') {
                    ivyValue = ivyValue > 0 ? 'Rec' : 'Pay';
                }

                if (mappingItem.markit === 'Notional') {
                    markitValue = Number(markitValue);
                }

                switch (mappingItem.comparison) {
                    case 'Daycount Leg 1':
                    case 'Daycount Leg 2': {
                        if (ivyValue.includes(markitValue) || markitValue.includes(ivyValue)) {
                            status = 'Matched';
                        }
                        break;
                    }
                    case 'Payment Frequency - leg 1':
                    case 'Payment Frequency - leg 2': {
                        console.log(mappingItem)
                        console.log(ivyValue, markitValue)
                        console.log(PaymentFrequencyLeg1Mapping.get(Number(ivyValue)), markitValue.toUpperCase())
                        console.log(typeof PaymentFrequencyLeg1Mapping.get(Number(ivyValue)), typeof markitValue.toUpperCase())
                        if (PaymentFrequencyLeg1Mapping.get(Number(ivyValue)) !== markitValue.toUpperCase()) {
                           console.log("went in if", status)
                            status = 'Unmatched';
                        }
                        break;
                    }
                    case 'Quantity/notional': {
                        if (Math.abs(ivyValue) !== Number(markitValue)) {
                            status = 'Unmatched';
                        }
                        break;
                    }
                    case 'Direction - pay/recv fixed': {
                        if (ivyValue === 'Receive' && markitValue === 'Rec') {
                            status = 'Matched';
                        }
                        break;
                    }
                    case 'Trade ID': {
                        status = '';
                        break;
                    }
                    case 'Contra Broker': {
                        if (!contraBroker[markitValue]?.includes(ivyValue)) {
                            status = 'Unmatched';
                        }
                        break;
                    }
                    default: {
                        if (
                            reviewRequiredByUser.includes(mappingItem.comparison) ||
                            (markitValue !== 0 && !markitValue) ||
                            (ivyValue !== 0 && !ivyValue)
                        ) {
                            status = 'Review required by user';
                        } else if (ivyValue !== markitValue) {
                            status = 'Unmatched';
                        }
                        break;
                    }
                }
                
                comparisonSheet.push([mappingItem.comparison, ivyValue, markitValue, status]);
            });
        }
    }

    return comparisonSheet;
}


function applyConditionalFormatting(sheet, comparisonSheet) {
    const statusColumnIndex = comparisonSheet[0].indexOf('Status');

    // Iterate over all rows except the header
    for (let i = 1; i < comparisonSheet.length; i++) {
        const cellAddress = XLSX.utils.encode_cell({ c: statusColumnIndex, r: i });
        const cell = sheet[cellAddress];

        if (cell && cell.v === 'Matched') {
            // Green color for "Matched"
            cell.s = {
                fill: {
                    patternType: 'solid',
                    fgColor: { rgb: 'C6EFCE' } // Light green background
                },
                font: {
                    color: { rgb: '006100' } // Dark green text
                }
            };
        } else if (cell && cell.v === 'Unmatched') {
            // Light red color for "Unmatched"
            cell.s = {
                fill: {
                    patternType: 'solid',
                    fgColor: { rgb: 'FFC7CE' } // Light red background
                },
                font: {
                    color: { rgb: '9C0006' } // Dark red text
                }
            };
        }
    }
}