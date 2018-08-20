var xlsx = require('node-xlsx');
var fs = require('fs');
//读取文件内容


const FILENAME = 'input_v7.xlsx';
const FILENAME_v1 = 'input_V7_2.xlsx';
const FILENAME_v2 = 'input_V7_3.xlsx';

var obj = xlsx.parse(__dirname+'/'+FILENAME);//配置excel文件的路径
var obj_v1 = xlsx.parse(__dirname+'/'+FILENAME_v1);
var obj_v2 = xlsx.parse(__dirname+'/'+FILENAME_v2);
// console.log(obj[3].name)
// console.log(obj[4].name)
// console.log(obj[5].name)
// console.log(obj[6].name)


var yhData = obj_v2[0].data

var v7Data = obj[3].data;

// var _40Data = obj[4].data;

// var _42Date = obj[5].data;

//console.log(v7Data[2],v7Data[3])
//console.log('房号：'+v7Data[3][2],'姓名：'+v7Data[3][3],'类型：'+v7Data[3][21])
// console.log('房号：'+_40Data[3][2],'姓名：'+_40Data[3][3],'类型：'+_40Data[3][21])
// console.log('定金记账日期:'+_40Data[3][22],'定金凭证号:'+_40Data[3][23],'首期记账日期:'+_40Data[3][24],'首期凭证号:'+_40Data[3][25],'银行按揭记账日期:'+_40Data[3][26],'银行按揭凭证号:'+_40Data[3][27])


// console.log('房号：'+_42Date[3][2],'姓名：'+_42Date[3][3],'类型：'+_42Date[3][20])
// console.log('定金记账日期:'+_42Date[3][21],'定金凭证号:'+_42Date[3][22],'首期记账日期:'+_42Date[3][23],'首期凭证号:'+_42Date[3][24],'银行按揭记账日期:'+_42Date[3][25],'银行按揭凭证号:'+_42Date[3][26])


var v7_2 = obj_v1[0].data;
var v7_aArr = [];
v7_2.map(function(v,k){
	
	if(!!v[6] && v[6]>0){
		if(!!v[5] && v[5].indexOf('本期合计')==-1 && v[5].indexOf('本年累计')==-1 ){
			//console.log(v[6],v[5])
			v7_aArr.push(v);
		}
	}

})

// function get40date(name){

// 	_40Data.map(function(value,key){
// 		//console.log(name, (value[2]+value[3]))
// 		if(name === (value[2]+value[3])){
// 			return [value[22],value[23],value[24],value[25],value[26],value[27]]
// 		}
// 	})
// 	return [];
// }

// function get42date(name){
// 	_42Date.map(function(value,key){
// 		if(name === (value[2]+value[3])){
// 			return [value[21],value[22],value[23],value[24],value[25],value[26]]
// 		}
// 	})
// 	return [];
// }
var resultArr = [];
resultArr[0] = v7Data[2];
v7Data.map(function(value,key){
	if(key>2){
		let value2 = value[2];
		let value3 = value[3];
		
		if(value2 || value3){
			resultArr[key-2] = value;
		}
	}
	
})

console.log(resultArr[1])

function getSameNameObj(name){
	//console.log(name)
	let a = null;
	if(name.length==1){
		v7_aArr.map(function(value,key){
		//console.log(value[5],name)
			if(value[5].indexOf(name[0])>-1){
				//console.log(value[5],name[0])
				a=value;
			}
		})
	}else if(name.length==2){
		v7_aArr.map(function(value,key){
		//console.log(value[5],name)
			if(value[5].indexOf(name[0])>-1){
				if(value[5].indexOf(name[1])>-1){
					//console.log(value[5],name[0],name[1])
					a=value;
				}	
			}
		})
	}
	
	return a;
}

function sepeName(name){
	if(name.indexOf(';')>-1){
		let arr = name.split(';');
		//console.log(arr)
		if(!arr[1]){
			return [arr[0]]
		}else{
			return arr
		}
	}else if(name.indexOf('/')>-1){
		
		let arr1 = name.split('/');
		//console.log(arr1)
		if(!arr1[1]){
			return [arr1[0]]
		}else{
			return arr1
		}
	}else{
		return [name]
	}
}


//yinhang yhData

// yhData.map(function(value,key){
// 	console.log(value[5])
	
// })


function getSameNameObjFromYH(name){
	let a = [];
	if(name.length==1){
		yhData.map(function(value,key){
		//console.log(value[5],name)
			if(value[5]){
				if(value[5].indexOf(name[0])>-1){
					//console.log(value[5],name[0])
					a.push([value[2],value[4],value[5]]);
				}
			}
			
		})
	}else if(name.length==2){
		yhData.map(function(value,key){
		//console.log(value[5],name)
			if(value[5]){
				if(value[5].indexOf(name[0])>-1){
					if(value[5].indexOf(name[1])>-1){
						//console.log(value[5],name[0],name[1])
						a.push([value[2],value[4],value[5]]);
					}	
				}
			}
			
		})
	}
	
	return a;
}

function getShou(arr){
	let sou=[],ding=[],dj=[],aj=[];
	arr.map(function(v,k){
		//console.log('数据：'+v,'长度：'+arr.length)

		if(v[2].indexOf('首期')>-1){
			if(v[2].indexOf('车位')==-1){
				sou.push(v);
			}
			
		}else if(v[2].indexOf('定期')>-1){
			
			if(v[2].indexOf('车位')==-1){
				ding.push(v);
			}
		}else if(v[2].indexOf('定金')>-1){
			
			if(v[2].indexOf('车位')==-1){
				dj.push(v);
			}
		}else if(v[2].indexOf('银行按揭')>-1){
			
			if(v[2].indexOf('车位')==-1){
				aj.push(v);
			}
		}
	})
	return {
		sou:sou,
		ding:ding,
		dj:dj,
		aj:aj
	}
}

function gethao(str){
	if(str.indexOf('#')>-1){
		let arr = str.split('#');
		return arr[1];
	}else if(str.indexOf('幢')>-1){
		let arr1 = str.split('幢');
		return arr1[1];
	}
}

function getObjByhao(hao,arr){
	let _arr = [];
	arr.map(function(v,k){
		if(v[2].indexOf(hao)>-1){
			_arr.push(v)
		}
	})
	return _arr;

}

function getminTime(arr){
	let _arr ,time=21181212;
	arr.map(function(v,k){
		if(~~(v[0].split('-').join(''))<time){
			time = ~~(v[0].split('-').join(''));
			_arr = arr[k];
		}
	})
	return _arr;
}

resultArr.map(function(value,key){
	if(key>0){
		let name = value[3];
		let now = getSameNameObj(sepeName(name));
		let yh = getSameNameObjFromYH(sepeName(name));
		//console.log(yh.length)
		//console.log(now)
		let obj= getShou(yh);
		let minaj = [];
		let minsou = [];
		//console.log('首期:'+obj.sou.length,'定期:'+obj.ding.length,'定金:'+obj.dj.length,'按揭:'+obj.aj.length)
		if(obj.aj.length>0){
			//console.log(obj.aj)
			let hao = gethao(value[2]);
			//console.log(hao)
			let ajarr = getObjByhao(hao,obj.aj);
			if(ajarr.length>0){
				minaj =  getminTime(ajarr);
				//console.log(minaj)
			}
			
			// if(ajarr.length==0||ajarr.length>2){
			// }
			
		}

		if(obj.sou.length>0){
			//console.log(obj.sou)
			let hao1 = gethao(value[2]);
			// //console.log(hao)
			let souarr = getObjByhao(hao1,obj.sou);
			//console.log(souarr)
			// if(souarr.length>0){
			// 	console.log(souarr)
			// }
			
			if(souarr.length>0){
				minsou =  getminTime(souarr);
				console.log(minsou)
			}
			
			// if(ajarr.length==0||ajarr.length>2){
			// }
			
		}



		
		

		if(now){
			resultArr[key][22] = now[2];
			resultArr[key][23] = now[4];
			resultArr[key][28] = now[5];
		}

		if(minaj.length>0){
			resultArr[key][26] = minaj[0];
			resultArr[key][27] = minaj[1];
			resultArr[key][29] = minaj[2];;
		}
		if(minsou.length>0){
			resultArr[key][24] = minsou[0];
			resultArr[key][25] = minsou[1];
			resultArr[key][30] = minsou[2];;
		}
		resultArr[key][31] = value[3];
	}
	
	
})


//return



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
        data:resultArr
    }        
]);

//将文件内容插入新的文件中
fs.writeFileSync('output.xlsx',buffer,{'flag':'w'});