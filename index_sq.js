var xlsx = require('node-xlsx');
//var sql = require('sqlite3');
var fs = require('fs');
//读取文件内容
var sqlite3 = require('sqlite3').verbose();

console.log(__dirname+"Backup.db")
var db = new sqlite3.Database(__dirname+"Backup.db", function(e){
 //if (err) throw err;
 console.log(e)
});

db.serialize(function() {
  db.run("show tables");
});
 
db.close();

return

const FILENAME = 'Original_May.xlsx';


var obj = xlsx.parse(__dirname+'/'+FILENAME);//配置excel文件的路径

var excelObj=obj[0].data;//excelObj是excel文件里第一个sheet文档的数据，obj[i].data表示excel文件第i+1个sheet文档的全部内容
// console.log(excelObj[1]);
//一个sheet文档中的内容包含sheet表头 一个excelObj表示一个二维数组，excelObj[i]表示sheet文档中第i+1行的数据集（一行的数据也是数组形式，访问从索引0开始）
var getH = function(date){
    return parseInt(date*24);
}

var getM = function(date){
    return parseInt((date*24*60)%60);
}

var getS = function(date){
    return parseInt((date*24*60*60)%60);
}


var arr = [];
for(var i in excelObj[1]){
    if(i<=4){
        arr[i] = excelObj[1][i];
    }else{
        arr[i] = {
            H:getH(excelObj[1][i]),
            M:getM(excelObj[1][i]),
            S:getS(excelObj[1][i])
        };
    }
}

//console.log(arr)



var workAttendanceArr = [];


excelObj.map(function(value,key){
    if(key === 0){
            workAttendanceArr[key] = [
                '部门名称',
                '人员编号',
                '姓名',
                '日期',
                '打卡次数',
                '最早打卡时间',
                '最晚打卡时间'
            ];
    }else{
        workAttendanceArr[key] = [];
        var max = 0;
        for(var i in value){
            if(i<=4){
                workAttendanceArr[key][i] = value[i];
            }else if(i==5){
                //console.log(!!value[i])
                if(!!value[i]){
                    workAttendanceArr[key][i] = value[i];
                }
            }else{
                max =  value[i];
            }
        }
        if(max!=0){
            workAttendanceArr[key][6] = max;
        }
    }
    
});

console.log(workAttendanceArr);


var b = new Buffer('JavaScript');
var s = b.toString('base64');


var data = [];

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
        data:workAttendanceArr
    }        
]);

//将文件内容插入新的文件中
fs.writeFileSync('test1.xlsx',buffer,{'flag':'w'});