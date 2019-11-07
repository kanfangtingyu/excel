const xlsx = require('node-xlsx').default;
const fs = require('fs')
const form = {
    name: '模拟数据表',
    data: [
        ['东京吃货',null,null,null,'海贼王'],
        ['东京吃货',null,null,null,'海贼王'],
        ['东京吃货',null,null,null,'海贼王'],
        ['东京吃货',null,null,null,'海贼王'],
        ['东京吃货',null,null,null,'海贼王'],
        ['东京吃货',null,null,null,'海贼王'],
        ['东京吃货',null,null,null,'海贼王'],
        ['东京吃货',null,null,null,'海贼王'],
        ['东京吃货',null,null,null,'海贼王'],
        ['东京吃货',null,null,null,'海贼王'],
    ]
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
                        horizontal: 'center'
                    },
                    font: {
                        size: 19,
                        color: {rgb: 'ff280c'}
                    }
                }
            })
        })
        mockData.push(line)


})
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
    '!rows': [//设置高度
        {hpx: 20}, //1
        {hpx: 20},//2
        {hpx: 20},//3
        {hpx: 20},//4
        {hpx: 20},//20
        {hpx: 20},//6
        {hpx: 20},//7
        {hpx: 20},//8
        {hpx: 20},//9
        {hpx: 20}, //1
        {hpx: 20},//2
        {hpx: 20},//3
        {hpx: 20},//4
        {hpx: 20},//20
        {hpx: 20},//6
        {hpx: 20},//7
        {hpx: 20},//8
        {hpx: 20},//9
    ],
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