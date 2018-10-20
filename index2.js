var xlsx = require('node-xlsx');
var fs = require('fs');
//读取文件内容


const FILENAME = 'input.xlsx';

//var arr = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
var arr = ['10月'];

var obj = xlsx.parse(__dirname+'/'+FILENAME);//配置excel文件的路径

console.log(obj[0].data);
var mobileArr = obj[0].data;
var mobileObj = {};

mobileArr.map(function(v,k){
    if(k>0){
        console.log(v[0])
        mobileObj[v[0]] = {};
    }
})

console.log(mobileObj)

//return;
var data = {};
var number = {};

arr.map(function(v,k){
	try{
		console.log('正在尝试读取'+v+'数据');
		data[(k+1)+'月'] = xlsx.parse(__dirname+'/'+v+'.xlsx');//配置excel文件的路径
	}catch(e){
		//console.log(e)
		console.log(v+'数据异常,如需生成'+v+'数据,请添加 '+v+'.xlsx 文件');
	}

	if(data[(k+1)+'月']){
		let _obj = data[(k+1)+'月'];
		let _data = _obj[0].data;

		var b = _data[1];
		console.log(b[10],'carno')
		console.log(b[21],'mobile')

		let _map = {};
		_data.map(function(va,ke){
			if(ke>0){
                
                
                for(let num in mobileObj){
                    if(va[21] == num){
                        if(parseInt(va[25]) == 0){
                            let count = mobileObj[num].count?parseInt(mobileObj[num].count):0;
                            let _count = count+1
                            mobileObj[num] = {
                                ...mobileObj[num],
                                count: _count
                            }
                        }else{
                            let count = mobileObj[num].count?parseInt(mobileObj[num].count):0;
                            let _count = count+1;

                            let price = mobileObj[num].price?mobileObj[num].price:0;
                            let _price = price+ (parseInt(va[25])/100)
                            mobileObj[num] = {
                                ...mobileObj[num],
                                count: _count,
                                price: _price
                            }
                        }
                    }
                }
				
			} 
		})

		console.log(mobileObj)

		// let _result = [];
		// _result[0] = ['手机号','次数','价格总数'];
		// for(let abc in mobileObj){
		// 	var $arr = [abc,mobileObj[abc].count?mobileObj[abc].count:0,mobileObj[abc].price?mobileObj[abc].price:0];
		// 	_result.push($arr);

			
		// }
		// console.log(_result)
		// let buffer = xlsx.build([
		// 	{
		// 		name: v + '数据',
		// 		data:_result
		// 	}		   
		// ]);
		
		// //将文件内容插入新的文件中
		// fs.writeFileSync(v+'output.xlsx',buffer,{'flag':'w'});

	}

	data[(k+1)+'月'] = null;
})

let _result = [];
_result[0] = ['手机号','次数','价格总数'];
for(let abc in mobileObj){
    var $arr = [abc,mobileObj[abc].count?mobileObj[abc].count:0,mobileObj[abc].price?mobileObj[abc].price:0];
    _result.push($arr);

    
}
console.log(_result)
let buffer = xlsx.build([
    {
        name: '数据',
        data:_result
    }		   
]);

//将文件内容插入新的文件中
fs.writeFileSync('10月会员优惠.xlsx',buffer,{'flag':'w'});


