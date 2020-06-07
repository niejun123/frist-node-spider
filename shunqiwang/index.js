const { Builder, By, Key, until } = require('selenium-webdriver');
const request  = require('request')
const cheerio = require('cheerio');
const fs = require('fs');
const xlsx = require('node-xlsx');//表格写入模块
const ourl = "http://so.11467.com/cse/search?q=%E5%B9%BF%E5%B7%9E%E5%81%A5%E5%BA%B7%E6%9C%8D%E5%8A%A1%E5%85%AC%E5%8F%B8&click=1&s=662286683871513660&nsid=1"
var driver = new Builder().forBrowser('chrome').build();
(async function start() {

    try {
        await driver.get(ourl);
        getData()
    } finally {
        
    }
})();
// pager-next-foot n
var result = []
var then = true;
var currentPage = 0;
var data = [
    ['企业名','负责人','联系方式','地址','备注']//表头
] 
  const options = {
    '!cols': [
      {wpx: 100},//1-变更名称
      {wpx: 100},//2-变更描述
      {wpx: 140},//3-计划上线测试时间
      {wpx: 140}, //4-计划上线时间
      {wpx: 250}, //5-子系统、模块名称
      {wpx: 120}, //6-依赖模块
      {wpx: 195},//7-功能点
      {wpx: 195}, //8-详细描述
      {wpx: 195}, //9-测试要点
      {wpx: 205}, //10-对应需求
      {wpx: 150}, //11-是否
      {wpx: 150}, //12-开发A
      {wpx: 150}, //13-开发B
      {wpx: 110}, //14-关联版本
      {wpx: 110}, //15-代码走查
    ],
    '!rows': [
      {hpx: 40,},
      {hpx: 60},
      {hpx: 80},
      {hpx: 100},
    ],
    '!margins': {left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3},
  }

  const range = {s: {c: 0, r: 0}, e: {c: 0, r: 2}}; // A1:A4
  options['!merges'] = [range]
//写入表格格式
async function getData() {
        let flag = true;
        while(then){
    
        try {
            let items = await driver.findElements(By.css('#results .result'))
            
            for (let i = 0; i <= items.length; i++) {
                let item = items[i]
                let companyName = await item.findElement(By.css('.c-title')).getText()
                let companyInformation = await item.findElement(By.css('.c-title a')).getAttribute('href')
                
                    getInfo(companyInformation)
               
                    //获取内容
                // let info = {
                //     companyName,
                //     companyInformation
                // }
                // result.push(info)
            }
          
           
        } catch (e) {
           
            if (e) flag = false

        } finally { 
            //
            currentPage++;
            if (currentPage <8) {
                await driver.findElement(By.css('#pageFooter .pager-next-foot')).click()
            } else {
               //爬完数据后写入excal表格
                var buffer = xlsx.build([{name: "企业名单", data: data}],options);
                fs.writeFileSync("./headit_enterprice_machine_test.xlsx",buffer,err=>{
                    console.log(err)
                })
                then = false
            }   
            if (flag) break
        }
 
    }
}
 
async function getInfo(item){
    
    try{
        request(item,(err,res,body)=>{
            const $ = cheerio.load(body);
         const rules = isNaN(parseInt($('.codl>dd:nth-child(4)').text()))==false || isNaN(parseInt($('.codl>dd:nth-child(8)').text()))==false &&  $('.codl>dd:nth-child(6)').text().length < 4 
            if(rules) {
                let enterprice = $('.codl>tbody>tr:nth-child(1)>td:nth-child(2)').text();
                let president = $('.codl>dd:nth-child(6)').text();
                let tel = $('.codl>dd:nth-child(4)').text()+'---photo:'+$('.codl>dd:nth-child(8)').text();
                let address = $('.codl>dd:nth-child(2)').text()
                data.push([enterprice,president,tel,address])
                console.log(enterprice,president,tel,address )
            } 
          
        })
    }catch (e){
        console.log(e)
    }finally{
      
    }
   
    
}