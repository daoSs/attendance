var xlsx = require('node-xlsx');
var fs = require('fs');
//读取文件内容


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



var workAttendanceArr = [];


excelObj.map(function(value,key){
    if(key === 0){
            workAttendanceArr[key] = excelObj[key];
    }else{
        workAttendanceArr[key] = [];
        for(var i in value){
            if(i<=4){
                workAttendanceArr[key][i] = value[i];
            }else{
                workAttendanceArr[key][i] = {
                    H:getH(value[i]),
                    M:getM(value[i]),
                    S:getS(value[i])
                };
            }
        }
       
    }
    
});

//console.log(workAttendanceArr);
var obj = {};
var setTimeSlot = function(value,year){
    //value= { H: 8, M: 10, S: 35 }
    var obj = {};
    var sw,gw,start,end;
    if(value.M>=10){
        sw = parseInt(value.M/10);
        gw = value.M%10;
    }else{
        gw = value.M;
        sw = 0;
    }

    if(gw > 4 && gw<=9){
        start = sw*10+5;
        end = sw*10+9;
    }else if(gw >= 0 && gw<5){
        start = sw*10;
        end = sw*10+4;
    }
    //console.log(start,end,value.H+'时'+ start+'分'+'到'+value.H+'时'+end+'分',value)
    return {
        start: start,
        end: end,
        str: value.H+'时'+ start+'分'+'到'+value.H+'时'+end+'分',
        sort: Date.parse(new Date(year+" "+value.H+':'+value.M+':'+value.S))
    }
}

var arrHaveTimeSlot = function(arr,slotStr){
    var flag = false,key = 0;
    arr.map(function(v,k){
        if(v[1] == slotStr){
            flag = true;
            key = k;
        }
    })
    return {flag:flag,key:key};
}

// var a = [ 1,
//     [ '工程部',
//       '000029381',
//       '李帅',
//       '2018-05-01',
//       '2',
//       { H: 8, M: 2, S: 22 },
//       { H: 19, M: 30, S: 41 } ] ];

workAttendanceArr.map(function(value,key){
    if(key>0){
        if(!obj[value[3]]){
            obj[value[3]] = [];
            // [
            //     [],
            //     [],
            //     [],
            // ]

        }
        value.map(function(v,i){
            if(i>4){
                let timeSlotObj =  setTimeSlot(value[i],value[3]);
                //console.log(timeSlotObj)
                let _arrHaveTimeSlot = arrHaveTimeSlot(obj[value[3]],timeSlotObj.str);
                //console.log(_arrHaveTimeSlot)
                if(_arrHaveTimeSlot.flag){
                    obj[value[3]][_arrHaveTimeSlot.key].push(value[2]+value[i].H+'时'+value[i].M+'分');
                }else{
                    let arr = [];
                    arr[0] = parseInt(timeSlotObj.sort);
                    arr[1] = timeSlotObj.str;
                    arr.push(value[2]+value[i].H+'时'+value[i].M+'分')
                    obj[value[3]].push(arr);
                    
                }
            }
        })        
    }
})

var resultArr = [];
for(let a in obj){
    //console.log(a,obj[a])
    resultArr.push({
        name: a,
        data: obj[a]
    })
}


var b = new Buffer('JavaScript');
var s = b.toString('base64');


var data = [];

var buffer = xlsx.build(resultArr);

//将文件内容插入新的文件中
fs.writeFileSync('test2.xlsx',buffer,{'flag':'w'});