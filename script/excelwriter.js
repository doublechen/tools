var fs = require('fs');
var xlsx = require('node-xlsx');

if (process.argv.length != 4){
    console.log('Excel Parser 参数错误！');
    process.exit();
}

//读入的json文件
var infile = process.argv[2];
fs.exists(infile, function(exists){
    if (!exists){
        console.log(infile + ' 文件不存在！');
        process.exit();
    }
});

var jsonfile = fs.readFileSync(infile);

try {
    var json = JSON.parse(jsonfile);
}catch(e){
    console.log('JSON 文件解析错误');
    process.exit(); 
}

var buffer = xlsx.build(json);

//要写入的文件地址
var outfile = process.argv[3];
fs.writeFileSync(outfile, buffer);
