let globalSheet;
let workbookGlobal;

function loadExcelFile() {
    const fileInput = document.getElementById('fileInput');
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = e.target.result;
        workbookGlobal = XLSX.read(data, {type: 'binary'});

        if (workbookGlobal.SheetNames.length > 1) {
            displaySheetSelectors(workbookGlobal.SheetNames);
        } else {
            globalSheet = workbookGlobal.Sheets[workbookGlobal.SheetNames[0]];
            displayColumnSelectors(globalSheet);
        }
    };
    reader.readAsBinaryString(fileInput.files[0]);
}

function displaySheetSelectors(sheetNames) {
    const sheetSelectorDiv = document.getElementById('sheetSelector');
    sheetSelectorDiv.innerHTML = '';

    sheetNames.forEach((name, index) => {
        const radio = document.createElement('input');
        radio.type = 'radio';
        radio.id = 'sheet' + index;
        radio.name = 'sheet';
        radio.value = name;
        if (index === 0) radio.checked = true; // Check the first sheet by default

        const label = document.createElement('label');
        label.htmlFor = 'sheet' + index;
        label.appendChild(document.createTextNode(name));
        sheetSelectorDiv.appendChild(radio);
        sheetSelectorDiv.appendChild(label);
        sheetSelectorDiv.appendChild(document.createElement('br'));
    });

    const button = document.createElement('button');
    button.textContent = 'Select Sheet';
    button.onclick = selectSheet;
    sheetSelectorDiv.appendChild(button);

    sheetSelectorDiv.style.display = 'block';
}

function selectSheet() {
    const selectedSheet = document.querySelector('#sheetSelector input[type="radio"]:checked').value;
    globalSheet = workbookGlobal.Sheets[selectedSheet];
    displayColumnSelectors(globalSheet);
}

function displayColumnSelectors(sheet) {
    const columnSelectorDiv = document.getElementById('columnSelectors');
    columnSelectorDiv.innerHTML = '';

    const columnNames = XLSX.utils.sheet_to_json(sheet, { header: 1 })[0];
    
    // 添加搜索列和显示列的标签和复选框
    ['search', 'display'].forEach(type => {
        const label = document.createElement('p');
        label.textContent = `请选择要${type === 'search' ? '搜索' : '显示'}的列:`;
        columnSelectorDiv.appendChild(label);

        const container = document.createElement('div');
        container.id = `${type}ColumnCheckboxes`;
        columnNames.forEach((name, index) => {
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `${type}Col${index}`;
            checkbox.value = name;
            checkbox.checked = false;

            const label = document.createElement('label');
            label.htmlFor = `${type}Col${index}`;
            label.appendChild(document.createTextNode(name));

            container.appendChild(checkbox);
            container.appendChild(label);
            container.appendChild(document.createElement('br'));
        });
        columnSelectorDiv.appendChild(container);
    });
    document.getElementById('columnSelectors').style.display = 'block';
}

function isDateColumn(columnName) {
    return columnName.toLowerCase().includes('date') || columnName.toLowerCase().includes('time');
}

function excelDateToJSDate(serial) {
    if (isNaN(serial)) return ""; 
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);

    if (isNaN(date_info.getTime())) {
        return ""; 
    }

    const formattedDate = date_info.toISOString().split('T')[0];
    return formattedDate;
}

function highlightText() {
    const keywordInput = document.getElementById('keyword').value;
    let keywordPattern;

    // 尝试创建一个正则表达式，如果失败，则捕获错误并提示
    try {
        // 'gi' 标志表示全局搜索且不区分大小写
        keywordPattern = new RegExp(keywordInput, 'gi');
    } catch (e) {
        alert('无效的正则表达式: ' + e.message);
        return;
    }

    const searchColumns = Array.from(document.querySelectorAll('#searchColumnCheckboxes input[type=checkbox]:checked'))
                               .map(input => input.value);
    const displayColumns = Array.from(document.querySelectorAll('#displayColumnCheckboxes input[type=checkbox]:checked'))
                                .map(input => input.value);
    const jsonSheet = XLSX.utils.sheet_to_json(globalSheet, { defval: "" });

    const resultTable = document.createElement('table');
    const headerRow = resultTable.insertRow();
    displayColumns.forEach(col => {
        const headerCell = document.createElement('th');
        headerCell.textContent = col;
        headerRow.appendChild(headerCell);
    });

    jsonSheet.forEach(row => {
        const isKeywordPresent = searchColumns.some(col => (row[col] || "").match(keywordPattern));
        if (isKeywordPresent) {
            const tableRow = resultTable.insertRow();
            displayColumns.forEach(col => {
                const cell = tableRow.insertCell();
                let cellValue = row[col] || "";
                if (isDateColumn(col)) {
                    const numericValue = parseFloat(cellValue);
                    if (!isNaN(numericValue)) {
                        cellValue = excelDateToJSDate(numericValue);
                    }
                }
                if (searchColumns.includes(col) && keywordPattern.test(cellValue.toString())) {
                    // 为了高亮显示，重建正则表达式确保不受之前测试的影响
                    const highlightPattern = new RegExp(keywordInput, 'gi');
                    cell.innerHTML = cellValue.toString().replace(highlightPattern, match => `<mark>${match}</mark>`);
                } else {
                    cell.textContent = cellValue;
                }
            });
        }
    });

    const resultDiv = document.getElementById('result');
    resultDiv.innerHTML = '';
    resultDiv.appendChild(resultTable);
}


