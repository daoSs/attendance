var xlsx = require('node-xlsx');
var fs = require('fs');
//读取文件内容


const FILENAME = 'input.xlsx';


var obj = xlsx.parse(__dirname+'/'+FILENAME);//配置excel文件的路径

var excelObj=obj[7].data;//excelObj是excel文件里第一个sheet文档的数据，obj[i].data表示excel文件第i+1个sheet文档的全部内容
 //console.log(excelObj[0],excelObj[1]);
//一个sheet文档中的内容包含sheet表头 一个excelObj表示一个二维数组，excelObj[i]表示sheet文档中第i+1行的数据集（一行的数据也是数组形式，访问从索引0开始）

var minute = 1000 * 60;
    var hour = minute * 60;
    var day = hour * 24;
    var month = day * 30;


var getH = function(date){
    return parseInt(date*24);
}

var getM = function(date){
    return parseInt((date*24*60)%60);
}

var getS = function(date){
    return parseInt((date*24*60*60)%60);
}
//console.log(excelObj[3],excelObj[3][15]);

// excelObj.map(function(value,key){
// 	console.log(value[15])
// })


//return
var a = new Date(1900,0,excelObj[3][12]).toLocaleString();
var b = new Date(1900,0,excelObj[3][15]).toLocaleString();
//console.log(a,b)

var ca = new Date(b).getTime()-new Date(a).getTime();
// console.log(ca)
// console.log(ca/minute);
// console.log(excelObj[4][15])
var da = new Date(1900,0,excelObj[4][15]).toLocaleString();
// console.log(da)
// console.log(new Date(da).getMinutes());
// console.log(new Date(da).getSeconds());



var result = [];
result[0] = excelObj[0]
excelObj.map(function(v,k){
	if(k>0){

		//console.log(v[12],v[15])
		var wechat = new Date(1900,0,v[12]).toLocaleString();
		var pos = new Date(1900,0,v[15]).toLocaleString();
	//console.log(a,b)
		if((new Date(wechat)).getMinutes()!=0 || (new Date(wechat)).getSeconds()!=0){
			//var ca =Math.abs(new Date(pos).getTime()-new Date(wechat).getTime());
			var ca =new Date(pos).getTime()-new Date(wechat).getTime();
			//console.log(ca)
			if(isNaN(ca)){
				console.log(ca)
				result.push(v)
			}else{
				if(ca/minute<=30){
					console.log(v[1],ca/minute)
					result.push(v)
				}
			}
			// if(ca/minute<=30){
			// 	//console.log(v[1],ca/minute)
			// }
		}else{
			//console.log(v[1])
		}
	}
	
})

console.log(result)
return




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
        data:result
    }        
]);

//将文件内容插入新的文件中
fs.writeFileSync('test1.xlsx',buffer,{'flag':'w'});