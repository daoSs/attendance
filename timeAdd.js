var xlsx = require('node-xlsx');
var fs = require('fs');
//读取文件内容


const FILENAME = 'input.xlsx';


var obj = xlsx.parse(__dirname+'/'+FILENAME);//配置excel文件的路径

var excelObj=obj[0].data;//excelObj是excel文件里第一个sheet文档的数据，obj[i].data表示excel文件第i+1个sheet文档的全部内容
console.log(excelObj[0],excelObj[1]);
//一个sheet文档中的内容包含sheet表头 一个excelObj表示一个二维数组，excelObj[i]表示sheet文档中第i+1行的数据集（一行的数据也是数组形式，访问从索引0开始）

//console.log(excelObj)

var map = {};
var arr = [];
arr[0] = excelObj[0];
for(var i in excelObj){
    if(i>0){
        let str = excelObj[i][3]+excelObj[i][5]+excelObj[i][8]+excelObj[i][9];

        let row = arr[arr.length-1];
        let _Str = row[3]+row[5]+row[8]+row[9];
        if(str == _Str){
            arr[arr.length-1] = [...arr[arr.length-1],excelObj[i][13]]
        }else{
            arr = [...arr,excelObj[i]]
        }
    }
}

console.log(arr)





// for(var i in excelObj){
//     var arr=[];
//     var value=excelObj[i];
//     for(var j in value){
//         arr.push(value[j]);
//     }

//     var urlStrForBase64 = '/destination/ttd/index.html#sightList/category.html?DistrictId='+ arr[0]+'&DistrictName='+ arr[1]+'&CateId='+ arr[2];
//     console.log(urlStrForBase64)
    
//     var b = new Buffer(urlStrForBase64);
//     var s = b.toString('base64');

   

//     arr.push('ctrip://wireless/h5?url='+ s +'&type=1');
//     //console.log(arr)

//     data.push(arr);
// }
var buffer = xlsx.build([
    {
        name:'sheet1',
        data:arr
    }        
]);

//将文件内容插入新的文件中
fs.writeFileSync('test22222.xlsx',buffer,{'flag':'w'});