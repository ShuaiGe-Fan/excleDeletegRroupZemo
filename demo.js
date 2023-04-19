const fs = require('fs')
const xlsx = require('node-xlsx');      // 读写xlsx的插件
//文件名 修改1.csv为需要修改的表格
let list = xlsx.parse("./input.xlsx")

let number = 0 //记录某一列需要删除的前置列数
const ABC = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'//
const ABCObj = {}
ABC.split('').forEach((item, index) => {
    ABCObj[item] = index
})

let data = list[0].data
data.forEach((item, index) => {
    // console.log(item[0])
    if (item[0]) {
        //这里写的是最后两行的第一列的名字,如果到这里就不执行了
        if (item[0].indexOf('SKU 汇总') > -1 || item[0].indexOf('总计') > -1) {
            return
        } else {
            //这里的F输入第几列为0 ,P为暂时记录第几列删除
            if (item[ABCObj['E']] === 0) {
                item[ABCObj['P']] = number;
            }
            number = 0
        }
    } else {
        number += 1
    }

});
for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][ABCObj['E']] === 0) {
        let number = data[i][ABCObj['P']]
        data.splice(i - number, number + 1)
        i = i - number
    }
}
// for (let i = 0; i < data.length; i++) {
//     console.log(data[i][0])
//     if (data[i][0]) {
//         data.splice(i , 1)
//         i = i - 1
//     }
// }

let xlsxObj = [{
    name: 'sheet',
    data: data
}]
setTimeout(() => {
    console.table(data)
}, 1000)

fs.writeFileSync('./output.xlsx', xlsx.build(xlsxObj), "binary");
