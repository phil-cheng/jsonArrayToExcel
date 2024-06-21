const fs = require('fs')
const xlsx = require('node-xlsx')

// 定义表头-需要跟json数组内对象的key对应上
const headerArray = ['begda', 'ename', 'icnum', 'infty', 'message', 'pernr', 'status'];
// 往excel插入表头行
const excelDate = [
  headerArray
]

// 加载json文件
const jsonFile = fs.readFileSync('1.json', 'utf-8')
const jsonObj = JSON.parse(jsonFile)
const jsonArray = jsonObj.data

// 往excel插入业务行数据
jsonArray.forEach(item => {
  // 准备每行的数据
  let rowArray = [];
  headerArray.forEach(headerKey => {
    rowArray.push(item[headerKey]);
  });
  // 把行数据插入excel
  excelDate.push(rowArray)
})

// 设置列宽 第一列和第二列 都是30
const sheetOptions = { '!cols': [{ wch: 30 }, { wch: 30 }] }
// 写入数据
const bufferZh = xlsx.build([{ name: '结果', data: excelDate }], {
  sheetOptions
})

// 导出生成文件
fs.writeFileSync('./result.xlsx', bufferZh, { flag: 'w' })


