// excel read ，data deal and export to excel
const fs = require('fs')
const convertExcel = require('excel-as-json').processFile
const json2xls = require('json2xls');

console.log('-------- convert start ----------')
let options = {
        sheet:'1',
        isColOriented: false,
        omitEmtpyFields: false
    },
    result = [],
    fields = ['cityName1','cityId1','cityName2','cityId2','calResult']

// 两两计算
const calFunc = (item1, item2) => {
    let param1 = item1.s1 * item2.s1 + item1.s2 * item2.s2 + item1.s3 * item2.s3,
        param2 = (Math.pow(item1.s1, 2) + Math.pow(item1.s2, 2) + Math.pow(item1.s3, 2)) *
                 (Math.pow(item2.s1, 2) + Math.pow(item2.s2, 2) + Math.pow(item2.s3, 2))

    return param1 / Math.sqrt(param2)
}

convertExcel('./citys.xlsx', './citys.json', options, (err, data) => {
    if(err) {
        console.log("JSON conversion failure: ", err)
    }

    let cityData = require('./citys.json')
    
    console.log('-------- data dealing ----------')

    cityData.map((item, index) => {
        cityData.map(($item, $index) => {
            if($index <= index) return

            result.push({
                cityName1: item.cityName,
                cityId1:   item.cityId,
                cityName2: $item.cityName,
                cityId2:   $item.cityId,

                calResult: calFunc(item, $item)

            })
        })
    })

    console.log('-------- data export to excel ----------')

    let xls = json2xls(result);
    fs.writeFileSync('cityCalResult.xlsx', xls, 'binary')

    console.log('-------- cal end ----------')
})

