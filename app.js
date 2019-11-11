const xlsx = require('node-xlsx').default;
const fs = require('fs');

// 引入基本数据
const dataList = require('./newdata.json')

// 单个表格高度汇总
let rowLength = []
// 单个表格宽度汇总
let colLength = [{wpx: 40},{wpx: 40}]
// 设置单个表格高度
for(let i=0;i<=1000;i++){
    let a = {hpx: 1}
    rowLength.push(a)
}
// 设置单个表格宽度
for(let i=0;i<=100;i++){
    let a = {wpx: 170}
    colLength.push(a)
}

// 计算所需要的总行数
let allConNum = 1000
dataList[0].newArr.forEach((element,index)=>{
    allConNum += element.height
})

// 计算所需的总列数
let allRowNum = 5
dataList.forEach(()=>{
    allRowNum += 1
})


// 创建一个空二维数组，对应excel的行列
let allData = []
for(let i = 0;i<=allConNum;i++){
    allData[i] = []
    for(let y = 0;y<=allRowNum;y++){
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
    let start = 100
    let end = 0
    // 一小时的时间高度
    let oneHour = 0
    let eachHeight = 0
    // 暂时存放初始start变量
    let startInit = 0
    // 第二列后的每一列总高度计算标识
    let colFlag = true
        element.newArr.forEach((ele,idx)=>{
            // 先根据本案进行一小时高度的设定
            if(index === 0){
                oneHour = oneHour + ele.height
            }
            // 存放该高度到数组列表中
            if(index === 0 && idx+1 === element.newArr.length){
                oneHourHeight.push(oneHour)
            }
            // 默认阶梯默认初始化改变
            // 先计算出总高度
            if(colFlag === true){
                element.newArr.forEach((eleHeight,idxHeight)=>{
                    eachHeight += eleHeight.height
                })
            }
            if(index !== 0 && idx === 0){

                start += oneHourHeight[0]-eachHeight
            }
            // 初始时存放start高度
            if(index === 0 && idx === 0){
                startInit = start
                console.log(startInit)
            }
            // 合并小时数单元格
            if(index === 0 && idx+1 === element.newArr.length){
                let mergeSingle = {
                    s:{c:index+1,r:startInit},
                    e:{c:index+1,r:startInit+oneHourHeight[0]}
                }
                range.push(mergeSingle)
                let a = ele.starttime.slice(0,2)
                allData[startInit][index+1] = a + '时'
            }
            // 循环每一项，对每一项进行单元格的合并
            // 判断出每个单元格的起始于截至位置
            end = start+ele.height
            let mergeSingle = {
                s:{c:index+2,r:start},
                e:{c:index+2,r:end-1}
            }
            range.push(mergeSingle)
            // 针对每一个单独的对象数据，保存到单元格中
            allData[start][index+2] = ele.programName + "\n" + ele.starttime + "-" + ele.endtime
            start = end
        })
    }
);
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
                            sz: 10,
                            bold: false,
                            color: {rgb: '000000'}
                        },
                        border: {
                            top:{ style: 'thin', color: { rgb: "d8d8d8" } },
                            bottom:{ style: 'thin', color: { rgb: "d8d8d8" }},
                            left:{ style: 'thin', color: { rgb: "d8d8d8" } },
                            right:{ style: 'thin', color: { rgb: "d8d8d8" } }
                        }
                    }
                })
        })
        mockData.push(line)
})

console.log(range)

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