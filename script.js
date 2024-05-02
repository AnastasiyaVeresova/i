
const select = document.getElementById('table-select');
const selectedRowsContainer = document.getElementById('selected-rows');
const pivotTableContainer = document.getElementById('pivot-table');
let workbook;

var url = "https://docs.google.com/spreadsheets/d/132NpYta9xqVm06Jlh-Ip-n3VxomjvOVj/edit?usp=sharing&ouid=106327186958861822886&rtpof=true&sd=true";

/* set up async GET request */
var req = new XMLHttpRequest();
req.open("GET", url, true);
req.responseType = "arraybuffer";

req.onload = function(e) {
  var workbook = XLSX.read(req.response);

  /* DO SOMETHING WITH workbook HERE */
};

req.send();

/*const file = 'book.xlsx';
fetch(file)
    .then(response => response.arrayBuffer())
    .then(data => {
        workbook = XLSX.read(data, { type: 'array' });
        workbook.SheetNames.forEach(sheetName => {
            const option = document.createElement('option');
            option.textContent = sheetName;
            option.value = sheetName;
            select.appendChild(option);
        });
    });
*/
select.addEventListener('change', () => {
    const selectedSheetName = select.value;
    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[selectedSheetName], { header: 1 });
    
    selectedRowsContainer.innerHTML = '';
    sheetData.forEach(rowData => {
        const row = document.createElement('div');
        row.textContent = rowData.join(' | ');
        row.addEventListener('click', () => {
            const pivotRow = document.createElement('div');
            pivotRow.textContent = rowData.join(' | ');
            pivotTableContainer.appendChild(pivotRow);
        });
        selectedRowsContainer.appendChild(row);
    });
});




// Подключение библиотеки SheetJS
/*var XLSX = require('xlsx');

1. Подключается библиотека SheetJS с помощью require. 
2. Загружается файл Excel с именем 'book.xlsx' с помощью XLSX.readFile(). 
3. Создается объект sheetTables, который сопоставляет названия листов Excel с соответствующими таблицами. 
4. Создается список (ul) с id 'sheet-list' и контейнер (div) с id 'sheet-info' на HTML странице. 
5. Для каждого названия листа Excel создается элемент списка (li), который при наведении мыши на него отображает информацию с соответствующего листа в таблице. 
6. При наведении на заголовок листа вызывается функция, которая преобразует данные с листа в формат JSON и отображает их в виде таблицы на HTML странице. 
7. Создается таблица (table) для отображения данных с листа. 
8. Добавляется возможность добавления новых строк данных через prompt и их отображения в таблице на HTML странице. 
 
Этот скрипт позволяет загружать и отображать данные из файла Excel на HTML странице, а также добавлять новые строки данных и отображать их в отдельной таблице.


// Загрузка файла Excel
var file = 'book.xlsx';
var workbook = XLSX.readFile(file);

// Сопоставление названий листов с соответствующими таблицами
const sheetTables = {
    'Таблица ячеек трансформатора 100-330': 'Таблица Т1',
    'Таблица2': 'Таблица Т2',
};

// Отображение списка названий листов на HTML странице
const list = document.getElementById('sheet-list');
const infoContainer = document.getElementById('sheet-info');

Object.keys(sheetTables).forEach(sheetName => {
    const item = document.createElement('li');
    item.textContent = sheetName;
    item.addEventListener('mouseover', () => {
        const targetSheet = sheetTables[sheetName];
        const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[targetSheet], { header: 1 });
        // Отображение информации с нужного листа при наведении на заголовок
        infoContainer.innerHTML = ''; // Очистить контейнер перед добавлением новой информации
        const table = document.createElement('table');
        sheetData.forEach(rowData => {
            const row = document.createElement('tr');
            rowData.forEach(cellData => {
                const cell = document.createElement('td');
                cell.textContent = cellData;
                row.appendChild(cell);
            });
            table.appendChild(row);
        });
        infoContainer.appendChild(table);
    });
    list.appendChild(item);
});

// Отображение списка названий листов в выпадающем списке
const dropdown = document.getElementById('sheet-dropdown');
Object.keys(sheetTables).forEach(sheetName => {
    const option = document.createElement('option');
    option.value = sheetName;
    option.textContent = sheetName;
    dropdown.appendChild(option);
});

dropdown.addEventListener('change', () => {
    const selectedSheetName = dropdown.value;
    const targetSheet = sheetTables[selectedSheetName];
    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[targetSheet], { header: 1 });
    // Отображение информации с выбранного листа при выборе из выпадающего списка
    infoContainer.innerHTML = ''; // Очистить контейнер перед добавлением новой информации
    const table = document.createElement('table');
    sheetData.forEach(rowData => {
        const row = document.createElement('tr');
        rowData.forEach(cellData => {
            const cell = document.createElement('td');
            cell.textContent = cellData;
            row.appendChild(cell);
        });
        table.appendChild(row);
    });
    infoContainer.appendChild(table);
});



const selectedRowsTable = document.getElementById('selected-rows');
let selectedRows = [];

function addRow() {
    const selectedRow = prompt('Введите данные для новой строки через запятую');
    if (selectedRow) {
        selectedRows.push(selectedRow.split(','));
        renderSelectedRows();
    }
}

function renderSelectedRows() {
    selectedRowsTable.innerHTML = '';
    selectedRows.forEach(rowData => {
        const row = document.createElement('tr');
        rowData.forEach(cellData => {
            const cell = document.createElement('td');
            cell.textContent = cellData;
            row.appendChild(cell);
        });
        selectedRowsTable.appendChild(row);
    });
}*/
