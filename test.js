const xlsx = require('xlsx');

// 读取Excel文件
const workbook = xlsx.readFile('mainland_textApplicationForm.xlsx');
const sheetName = workbook.SheetNames[0]; // 获取第一个工作表的名字
const worksheet = workbook.Sheets[sheetName];

// 将工作表转换为JSON格式
let jsonData = xlsx.utils.sheet_to_json(worksheet);

// 创建一个Map来存储最新的InGameKey及其对应的mtime
let latestEntries = new Map();

jsonData.forEach(row => {
    let inGameKey = row['客户端读取的key(InGameKey)'];
    let mtime = row['修改时间(_mtime)'];

    // 如果Map中不存在此InGameKey或者当前记录的时间较新，则更新Map
    if (!latestEntries.has(inGameKey) || new Date(latestEntries.get(inGameKey).toString()) < new Date(mtime.toString())) {
        latestEntries.set(inGameKey, mtime);
    }
});

// 转换Map为数组形式，包含原始数据中的所有字段
let filteredData = jsonData.filter(row => 
    latestEntries.get(row['客户端读取的key(InGameKey)']).toString() === row['修改时间(_mtime)'].toString()
);

console.log(filteredData);

// 如有需要，将filteredData保存到新的Excel文件
const newWorkbook = xlsx.utils.book_new();
const newWorksheet = xlsx.utils.json_to_sheet(filteredData);
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'FilteredData');
xlsx.writeFile(newWorkbook, 'filtered_mainland_textApplicationForm.xlsx');