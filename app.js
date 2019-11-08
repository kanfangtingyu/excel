const xlsx = require('node-xlsx').default;
const fs = require('fs');

// 引入基本数据
const dataList = require('./newdata.json')
let rowLength = []
let colLength = []

// 创建一个空二维数组，对应excel的行列
let allData = []
for(let i = 0;i<=10000;i++){
    allData[i] = []
    for(let y = 0;y<=100;y++){
        allData[i].push('')
    }
}

// 创建一个空数组，存放每一个小时的高度
let oneHourHeight = []
// 创建一个空数组,定义单元格的合并
let range = []


// 循环data，对每一项进行处理
// 第一个数组为第几个本案，代案
// 第二个数组为本案，代案中有多少单独项
dataList.forEach((element,index) => {
     // 初始高度
    let start = 0
    let end = 0
    console.log(element)
        element.newArr.forEach((ele,idx)=>{
            // 先根据本案进行一小时高度的设定
            oneHourHeight[index] += ele.height
            // 循环每一项，对每一项进行单元格的合并
            // 判断出每个单元格的起始于截至位置
            end += ele.height
            let mergeSingle = {
                s:{c:index,r:start},
                e:{c:index,r:end-1}
            }
            range.push(mergeSingle)
            // 针对每一个单独的对象数据，保存到单元格中
            allData[start][index] = ele.programName + "\n" + ele.starttime + "\n" + ele.endtime
            start = end
        })
    }
);
// 设置单个表格高度
for(let i=0;i<=1000;i++){
    let a = {hpx: 5}
    rowLength.push(a)
}
// 设置单个表格宽度
for(let i=0;i<=100;i++){
    let a = {wpx: 150}
    colLength.push(a)
}

const form = {
    data: allData
}

const mockData = []
const form2 = {
    name: '报表',
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


const options = {
    '!cols': colLength,
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