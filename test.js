var xlsx = require('node-xlsx');
var fs = require('fs');
//读取文件内容


const FILENAME = 'input.xlsx';

var arr = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];

var data = {};
var number = {};


// var asd = xlsx.parse(__dirname+'/7月.xlsx');

// console.log(asd[0].data)
// console.log(asd[0].data[1][10],'carno')
// console.log(asd[0].data[1][21],'mobile')
// return

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
				if(!!va[21]){
					if(_map[va[21]]){
						if(_map[va[21]].indexOf(va[10]) == -1){
							_map[va[21]].push(va[10]);
						}
						
					}else{
						_map[va[21]] = [va[10]];
					}
				}
				
			} 
        })
        

		//console.log(_map)

		let _result = [];
		_result[0] = ['手机号','车牌号'];
		for(let abc in _map){
			var $arr = [abc,..._map[abc]];
			_result.push($arr);

			if(number[abc]){
				_map[abc].map(function(value,key){
					if(number[abc].indexOf(value) == -1){
						number[abc].push(value);
					}
				})
			}else{
				number[abc] = _map[abc];
			}
		}
        //console.log(_result)
        
        _result.map(function(v,k){
            //console.log(v)
            (function(va,ke){
                setTimeout(function(){
                    if(v.length>2){
                        console.log('******************** '+'mobile number: '+va[0]+' repeats.Please Check it!'+' ********************')
                    }else{
                        console.log('******************** '+'mobile number: '+va[0]+' is normal.'+' ********************')
                    }
                },ke*20)
            })(v,k)
            
        })

		let buffer = xlsx.build([
			{
				name: v + '数据',
				data:_result
			}		   
		]);
		
		//将文件内容插入新的文件中
		fs.writeFileSync(v+'output.xlsx',buffer,{'flag':'w'});

	}

	data[(k+1)+'月'] = null;
})

var sres = [];

sres[0] = ['手机号','车牌号'];

for(let a in number){
	if(number[a] && number[a].length && number[a].length>1){
		let $arr = [a,...number[a]];
		sres.push($arr);
	}
}

let buffer1 = xlsx.build([
	{
		name: '重复车牌号汇总',
		data:sres
	}		   
]);

//将文件内容插入新的文件中
fs.writeFileSync('同一手机对应多个车牌.xlsx',buffer1,{'flag':'w'});

