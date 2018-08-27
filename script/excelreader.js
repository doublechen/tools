var fs = require('fs');
var xlsx = require('node-xlsx');

if (process.argv.length != 4){
    console.log('Excel Parser 参数错误！');
    process.exit();
}

var infile = process.argv[2];
fs.exists(infile, function(exists){
    if (!exists){
        console.log(infile + ' 文件不存在！');
        process.exit();
    }
});

var outfile = process.argv[3];

var obj = xlsx.parse(fs.readFileSync(infile));

if (!obj.length){
    console.log('解析文件失败！');
    process.exit();
}

var data = obj;

fs.writeFileSync(outfile, JSON.stringify(data));