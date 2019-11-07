const xlsx = require('node-xlsx').default;
const fs = require('fs');
const data = require('./data.json')
let rowLength = []
let colLength = []
// let newdata = []
// data.forEach((element,index) => {
//     let centerData = []
//     element.newArr.forEach(ele=>{
//         centerData.push(ele.programName)
//     })
//     newdata.push(centerData)
// });
// console.log(newdata)
newdata = [
    ['ab\ncde','ddd',null,null,'dddddddd\nddddd'],
    ['abcde','ddd',null,null,'ddddddddddddd'],
    ['abcde','ddd',null,null,'ddddddddddddd'],
    ['abcde','ddd',null,null,'ddddddddddddd'],
    ['abcde','ddd',null,null,'ddddddddddddd'],
    ['abcde','ddd',null,null,'ddddddddddddd']
]
newdata.forEach(()=>{
    let a = {hpx: 20}
    rowLength.push(a)
})
const form = {
    name: '模拟数据表',
    data: newdata
}

const mockData = []
const form2 = {
    name: '认真的表格',
    data: mockData
}

form.data.map((v, i) => {
    
        const line = []
        v.map((item, i) => {
            line.push({
                v: item,
                s: {
                    alignment: {
                        vertical: 'center',
                        horizontal: 'center',
                        wrapText:true
                    },
                    font: {
                        size: 19,
                        color: {rgb: '000000'}
                    }
                }
            })
        })
        mockData.push(line)
})

console.log(mockData)
const range = [
    {s: {c: 0, r:0 }, e: {c:3, r:6}},
    {s: {c: 0, r:7 }, e: {c:3, r:8}},
    {s: {c: 0, r:9 }, e: {c:3, r:12}},
    {s: {c: 4, r:0 }, e: {c:7, r:4}},
    {s: {c: 4, r:5 }, e: {c:7, r:12}},
    {s: {c: 4, r:13 }, e: {c:7, r:14}},
];

const options = {
    '!cols': [ //设置宽度
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
        {wpx: 50},
    ],
    //高度设置无效
    '!rows': rowLength,
    '!merges':range,
    '!margins': {left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3},
}


const xlsxData = xlsx.build([form2], options)


console.log("准备写入文件");
fs.writeFile('input.xlsx', xlsxData, function (err) {
    if (err) {
        return console.error(err);
    }
    console.log("数据写入成功！");
    console.log("--------我是分割线-------------")

});