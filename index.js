const xl = require('exceljs');

const workbook = new xl.Workbook();

workbook.creator = 'Abhishek_Dhillon';
workbook.created = new Date(2020, 05, 24);
workbook.modified = new Date();

const sheet = workbook.addWorksheet('TimeSheet');

//Create an Array of Date and Days of current month
var date = new Date(),
    l = new Date(date.getFullYear(), date.getMonth() + 1, 0),
    dateArr = [],
    dayArr = [],
    temp = [{ name: '1' }, { name: '2' }, { name: 'Date:' }],
    vals = ['Client', 'Project', 'Task Description']
for (let f, i = 1; i < 32; i++) {
    f = new Date(date.getFullYear(), date.getMonth(), i);
    if (f.getDate() == l.getDate()) {
        dateArr.push({ name: f.toLocaleDateString() });
        dayArr.push(f.toDateString().split(' ')[0]);
        break;
    }
    else {
        dateArr.push({ name: f.toLocaleDateString() });
        dayArr.push(f.toDateString().split(' ')[0]);
    }
}

var col1 = [...temp, ...dateArr.slice(0, 7), { name: '-' }];
var row1 = [...vals, ...dayArr.slice(0, 7), 'Total'];


sheet.addTable({
    name: 'Table1',
    ref: 'C3',
    headerRow: true,
    style: {
        border: {
            top: "thick",
            left: "thick",
            right: "thick",
            bottom: "thick"
        }
    },
    columns: col1,
    rows: [
        row1,
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['Total', '', '', '', '', '', '', '', '', ''],
    ]
});


var col1 = [...temp, ...dateArr.slice(7, 14), { name: '-' }];
var row1 = [...vals, ...dayArr.slice(7, 14), 'Total'];
sheet.addTable({
    name: 'Table2',
    ref: 'C13',
    headerRow: true,
    style: {
        border: {
            top: "thick",
            left: "thick",
            right: "thick",
            bottom: "thick"
        }
    },
    columns: col1,
    rows: [
        row1,
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['Total', '', '', '', '', '', '', '', '', ''],
    ]
});

var col1 = [...temp, ...dateArr.slice(14, 21), { name: '-' }];
var row1 = [...vals, ...dayArr.slice(14, 21), 'Total'];
sheet.addTable({
    name: 'Table3',
    ref: 'C23',
    headerRow: true,
    style: {
        border: {
            top: "thick",
            left: "thick",
            right: "thick",
            bottom: "thick"
        }
    },
    columns: col1,
    rows: [
        row1,
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['Total', '', '', '', '', '', '', '', '', ''],
    ]
});


var col1 = [...temp, ...dateArr.slice(21, 28), { name: '-' }];
var row1 = [...vals, ...dayArr.slice(21, 28), 'Total'];
sheet.addTable({
    name: 'Table4',
    ref: 'C33',
    headerRow: true,
    style: {
        border: {
            top: "thick",
            left: "thick",
            right: "thick",
            bottom: "thick"
        }
    },
    columns: col1,
    rows: [
        row1,
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', ''],
        ['Total', '', '', '', '', '', '', '', '', ''],
    ]
});


sheet.getColumn(3).alignment = {
    vertical: 'top',
    horizontal: 'left',
    wrapText: true
};
sheet.getColumn(3).font = {
    bold: true
}
sheet.getColumn(3).width = 20

sheet.getColumn(4).alignment = {
    vertical: 'top',
    horizontal: 'left',
    wrapText: true
};
sheet.getColumn(4).font = {
    bold: true
}
sheet.getColumn(4).width = 20

sheet.getColumn(5).alignment = {
    vertical: 'top',
    horizontal: 'left',
    wrapText: true
};
sheet.getColumn(5).width = 20

workbook.xlsx.writeFile("Monthly-TimeSheet.xlsx");

