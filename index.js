
var xlsx = require('node-xlsx');
var fs = require('fs');
var parseString = require('xml2js').parseString;
var sheets = xlsx.parse('./from/from.xlsx');
var rowList = [];
sheets.forEach(function(sheet){
    for(var rowId in sheet['data']){
        rowList.push(sheet['data'][rowId][0]);
    }
});
var nodeList = [];
rowList.forEach((xml_string, index)=> {
        parseString(xml_string, { explicitArray : false, ignoreAttrs : true }, function (err, result) {
            if (Object.prototype.toString.call((result)) == '[object Object]') {
                var row = result.note;
                nodeList.push([row.to, row.from, row.heading, row.body]);
            }
        });
});
writeXls(nodeList);
function writeXls(datas) {
    var buffer = xlsx.build([
        {
            name:'sheet1',
            data:datas   
        }
    ]);
    fs.writeFileSync('./dist/dist.xlsx',buffer,{'flag':'w'});   //生成excel
}
