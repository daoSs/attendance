var xlsx = require('node-xlsx');
var fs = require('fs');
//读取文件内容


const FILENAME = 'iP_origin.xlsx';


var obj = xlsx.parse(__dirname+'/'+FILENAME);//配置excel文件的路径
// console.log(obj[1].data)
var excelObj=obj[1].data;//excelObj是excel文件里第一个sheet文档的数据，obj[i].data表示excel文件第i+1个sheet文档的全部内容
//console.log(excelObj[0]);
//一个sheet文档中的内容包含sheet表头 一个excelObj表示一个二维数组，excelObj[i]表示sheet文档中第i+1行的数据集（一行的数据也是数组形式，访问从索引0开始）
var getIPFirstTwoPosition = function(ip){
    ip = ''+ip;
    let arr = ip.split('.');
    return arr[0]+'.'+arr[1];
}





var ipObj = {};
var nameObj = {};
var ipTwoObj = {};
excelObj.map(function(v,k){
    if(k>0){
        let ip = v[5];
        let name = v[4];
        let ipTwo = getIPFirstTwoPosition(v[5]);
        console.log(ipTwo)
        if(!ipObj[ip]){
            ipObj[ip] = [excelObj[k]];
        }else{
            ipObj[ip].push(excelObj[k])
        }

        if(!nameObj[name]){
            nameObj[name] = [excelObj[k]];
        }else{
            nameObj[name].push(excelObj[k])
        }

        if(!ipTwoObj[ipTwo]){
            ipTwoObj[ipTwo] = [excelObj[k]];
        }else{
            ipTwoObj[ipTwo].push(excelObj[k])
        }

    }
})



var resultIpTwoObj = {};
var ipTwoSameCompany = {};
for(let a in ipTwoObj){
    if(ipTwoObj[a].length>1){
        //resultIpObj[a] = ipObj[a];
        let num = 0, numSameIp = 0 ,_map = {},_mapIp = {},nameArr = [];
        ipTwoObj[a].map(function(v,k){
            //console.log(v[4])
            if(!_map[v[4]]){
                _map[v[4]] = true;
                nameArr.push(v[4]);
                //console.log(num)
                num++
            }
            if(!_mapIp[v[5]]){
                _mapIp[v[5]] = true;
                numSameIp++;
            }
        })
        if(num>1){
            if(numSameIp>1){
                ipTwoObj[a].map(function(v,k){
                    ipTwoObj[a][k] = [...ipTwoObj[a][k],a]
                })
                resultIpTwoObj[a] = ipTwoObj[a];
                ipTwoSameCompany[a] = nameArr;
            }
            
        }    
    }
}
console.log(resultIpTwoObj)

//console.log(ipObj)
var resultIpObj = {};
var ipSameCompany = {};
for(let a in ipObj){
    if(ipObj[a].length>1){
        //resultIpObj[a] = ipObj[a];
        let num = 0,_map = {},nameArr = [];
        ipObj[a].map(function(v,k){
            //console.log(v[4])
            if(!_map[v[4]]){
                _map[v[4]] = true;
                nameArr.push(v[4]);
                //console.log(num)
                num++
            }
        })
        if(num>1){
            
            ipObj[a].map(function(v,k){
                ipObj[a][k] = [...ipObj[a][k],nameArr.join(',')]
            })
            resultIpObj[a] = ipObj[a];
            ipSameCompany[a] = nameArr;
        }    
    }
}


var resultNameObj = {};
var ipDifferenceButSameCompany = {};
for(let a in nameObj){
    if(nameObj[a].length>1){
        let num = 0,_map = {},nameArr = [];
        nameObj[a].map(function(v,k){
            //console.log(v[4])
            if(!_map[v[5]]){
                _map[v[5]] = true;
                nameArr.push(v[5]);
                //console.log(num)
                num++
            }
        })
        if(num>1){
            
            nameObj[a].map(function(v,k){
                nameObj[a][k] = [...nameObj[a][k],nameArr.join('---')]
            })
            resultNameObj[a] = nameObj[a];
            ipDifferenceButSameCompany[a] = nameArr;
        }    
    }
}

//console.log(resultNameObj)

var sameIpTwoArr = [[ 
    '公司名称',
    '归属项目',
    '采购方案名称',
    '采购方案编号',
    '供应商名称',
    'IP地址',
    '回标时间',
    '操作类型',
    'ip前两位' ]];

for(let a in resultIpTwoObj){
    sameIpTwoArr = [...sameIpTwoArr,...resultIpTwoObj[a]]
}

var sameIpArr = [[ 
    '公司名称',
    '归属项目',
    '采购方案名称',
    '采购方案编号',
    '供应商名称',
    'IP地址',
    '回标时间',
    '操作类型',
    '相同的公司名称罗列' ]];

for(let a in resultIpObj){
    sameIpArr = [...sameIpArr,...resultIpObj[a]]
}

var sameNameArr = [[ 
    '公司名称',
    '归属项目',
    '采购方案名称',
    '采购方案编号',
    '供应商名称',
    'IP地址',
    '回标时间',
    '操作类型',
    '相同的IP罗列' ]];

for(let a in resultNameObj){
    sameNameArr = [...sameNameArr,...resultNameObj[a]]
}


var resultArr = [];



var b = new Buffer('JavaScript');
var s = b.toString('base64');


var buffer = xlsx.build([
    {
        name:'源数据',
        data: excelObj
    },
    {
        name:'相同ip',
        data: sameIpArr
    },
    {
        name:'相同公司不同ip',
        data: sameNameArr
    },
    {
        name:'相同ip前两位',
        data: sameIpTwoArr
    }       
]);

fs.writeFileSync('test1.xlsx',buffer,{'flag':'w'});

//将文件内容插入新的文件中
//fs.writeFileSync('test2.xlsx',buffer,{'flag':'w'});