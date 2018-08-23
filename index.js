var xlsx = require('node-xlsx');
var fs = require('fs');
//读取文件内容


const FILENAME = 'input.xlsx';


var obj = xlsx.parse(__dirname+'/'+FILENAME);//配置excel文件的路径

// obj.map((v,k)=>{
// 	console.log(v.name,k)
// })

//console.log(obj[6].data[3],obj[6].data[4])
var a = obj[6].data[3];


console.log(a[14],a[15],a[16])


var gll7 = obj[6].data;
var fxt7 = obj[7].data;



//excelObj是excel文件里第一个sheet文档的数据，obj[i].data表示excel文件第i+1个sheet文档的全部内容
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
var aaa = new Date(1900,0,a[15]).toLocaleString();
//var bbb = new Date(a[16]).getTime();
var ccc = new Date(aaa).getTime()
console.log(ccc,'qqqqqqqqq')
//return
//var ca = new Date(b).getTime()-new Date(a).getTime();
// console.log(ca)
// console.log(ca/minute);
// console.log(excelObj[4][15])
//var da = new Date(1900,0,excelObj[4][15]).toLocaleString();
// console.log(da)
// console.log(new Date(da).getMinutes());
// console.log(new Date(da).getSeconds());



function main(excelObj){
	var resultwp1d = [];
	resultwp1d[0] = excelObj[0];
	resultwp1d[1] = excelObj[1];
	resultwp1d[2] = excelObj[2];

	var resultpa1d = [...resultwp1d];
	var resultwp2h = [...resultwp1d];
	var resultpa2h = [...resultwp1d];
	excelObj.map(function(v,k){
		if(k>2){
	
			//console.log(v[12],v[15])
			var wechat = new Date(1900,0,v[14]).toLocaleString();
			var pos = new Date(1900,0,v[15]).toLocaleString();
			var app = new Date(v[16]).getTime();
		//console.log(a,b)
			//if((new Date(wechat)).getMinutes()!=0 || (new Date(wechat)).getSeconds()!=0){
				//var ca =Math.abs(new Date(pos).getTime()-new Date(wechat).getTime());
				var wp = Math.abs(new Date(wechat).getTime()-new Date(pos).getTime());
				var pa = Math.abs(new Date(pos).getTime()-app);
				//console.log(ca)
				 
				//console.log(wp/day,pa/day,wp/hour,pa/hour)
				if(wp/day<=1){
					//console.log(wp/day)
					resultwp1d.push(v)
				}

				if(pa/day<=1){
					console.log(pa/day)
					console.log(pos,v[16])
					resultpa1d.push(v)
				}

				if(wp/hour<=2){
					//console.log(wp/hour)
					resultwp2h.push(v)
				}

				if(pa/hour<=2){
					resultpa2h.push(v)
				}
				
				
				// if(ca/minute<=30){
				// 	//console.log(v[1],ca/minute)
				// }
			// }else{
			// 	//console.log(v[1])
			// }
		}
		
	})

	return {resultwp1d,resultpa1d,resultwp2h,resultpa2h}
}
var _obj = main(fxt7)


//console.log(result)
//return




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
        name:'微信pos1天内',
        data:_obj.resultwp1d
	},
	{
        name:'posapp1天内',
        data:_obj.resultpa1d
	},
	{
        name:'微信pos2小时',
        data:_obj.resultwp2h
	},
	{
        name:'posapp2小时',
        data:_obj.resultpa2h
    }  
	       
]);

//将文件内容插入新的文件中
fs.writeFileSync('test2.xlsx',buffer,{'flag':'w'});