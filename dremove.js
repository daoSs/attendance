var xlsx = require('node-xlsx');
var fs = require('fs');
//读取文件内容


const FILENAME = 'input.xlsx';


var obj = xlsx.parse(__dirname+'/'+FILENAME);//配置excel文件的路径
console.log(obj.length)
var excelObj=obj[0].data;//excelObj是excel文件里第一个sheet文档的数据，obj[i].data表示excel文件第i+1个sheet文档的全部内容


console.log(excelObj[0],excelObj[1],excelObj[2]);
//return;
//一个sheet文档中的内容包含sheet表头 一个excelObj表示一个二维数组，excelObj[i]表示sheet文档中第i+1行的数据集（一行的数据也是数组形式，访问从索引0开始）

//console.log(excelObj)

var map = {};
var arr = [];
var darr = [];
arr = [excelObj[0],excelObj[1]];
for(var i in excelObj){
    if(i>0){
        let str = excelObj[i][2]+excelObj[i][3];
        //console.log(str)
        if(!map[str]){
            arr = [...arr,excelObj[i]];
            map[str] = true;
        }else{
            console.log('序号'+excelObj[i][1]+'重复,姓名与房间号为：'+ str)
            darr.push(['序号'+excelObj[i][1]+'重复,姓名与房间号为：'+ str])
        }
    }
}

//console.log(arr)





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

var buffer1 = xlsx.build([
    {
        name:'sheet1',
        data:darr
    }        
]);

//将文件内容插入新的文件中
fs.writeFileSync('test23333.xlsx',buffer,{'flag':'w'});
fs.writeFileSync('重复项.xlsx',buffer1,{'flag':'w'});
