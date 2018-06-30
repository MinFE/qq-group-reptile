
const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');
const xlsx = require('node-xlsx').default;
const completeDownloadGroups = getCompleteDownloadGroups();
let allGroupList = [];
let allGroupData = {};
let idx = 0;
let browser;

start();

async function start() {
    const page = await launch();

    // 劫持请求
    page.on('request', interceptedRequest => {
        let postData = interceptedRequest.postData();
        interceptedRequest.continue({
            postData: checkRequestIsSearchGroup(interceptedRequest)
                ? postData.replace(/end=.+&/, 'end=99999&')
                : postData
        });
    });
    page.on('requestfinished', async interceptedRequest => {
        if (checkRequestGroupList(interceptedRequest)) {
            let res = await interceptedRequest.response().json();
            Object.keys(res)
                .filter(key => Array.isArray(res[key]))
                .forEach(key => res[key].forEach(item => allGroupList.push(item)));
            
            setTimeout(() => {
                page.$$eval('.my-all-group .my-group-list li', lis => {
                    lis.forEach(li => li.click());
                });
            });
        } else if (checkRequestIsSearchGroup(interceptedRequest)) {
            let postData = interceptedRequest.postData();
            let match = postData.match(/gc=([0-9]+)&/);
            let res = await interceptedRequest.response().json();
            allGroupData[match[1]] = res;
            idx ++;

            idx === allGroupList.length && createDownloadData();
        }
    });

    await page.goto('https://qun.qq.com/member.html');
}

// 启动 Chromium
async function launch() {
    browser = await puppeteer.launch({
        headless: false
    });
    // const context = await browser.createIncognitoBrowserContext();
    const page = await browser.newPage();
    await page.setRequestInterception(true);
    return page;
}

// 获取已爬取的群
function getCompleteDownloadGroups() {
    return fs.readdirSync('./data')
        .filter(item => path.extname(item) === '.xlsx')
        .map(item => {
            let match = item.match(/.+-id(.+)\.(.+)/);
            return match[1];
        });
}

// 判断是否为获取群成员的请求
function checkRequestIsSearchGroup(request) {
    const url = request.url();
    const method = request.method();

    return url === 'https://qun.qq.com/cgi-bin/qun_mgr/search_group_members' && method === 'POST';
}

// 判断是否为获取群列表的请求
function checkRequestGroupList(request) {
    const url = request.url();
    const method = request.method();

    return url === 'https://qun.qq.com/cgi-bin/qun_mgr/get_group_list' && method === 'POST';
}

// 写入文件
function createDownloadData() {
    Object.keys(allGroupData)
        .forEach(groupId => {
            let res = allGroupData[groupId];
            let groupInfo = allGroupList.find(item => item.gc == groupId);

            if (!completeDownloadGroups.includes(groupId.toString())) {
                const option = {'!merges': [
                    {s: {c: 0, r: 0}, e: {c: 6, r: 0}}
                ]};
                const data = [
                    [`${groupInfo.gn}(${groupId})`],
                    ['成员', '群名片', 'QQ号', '性别', 'Q龄', '入群时间', '最后发言']
                ];

                res.mems.forEach(user => {
                    data.push([
                        user.nick,
                        user.card,
                        user.uin.toString(),
                        formatSex(user.g),
                        user.qage,
                        formatDate(user.join_time * 1000),
                        formatDate(user.last_speak_time * 1000)
                    ]);
                });
        
                const buffer = xlsx.build(
                    [{name: "群成员名单", data}],
                    option
                );
                console.log(`[download data]: ${groupInfo.gn}`);
                fs.writeFileSync(`./data/${groupInfo.gn}-id${groupId}.xlsx`, buffer);
            }
        });
    
    console.log('数据导出完毕');
    browser.close();
}

// 格式性别
function formatSex(type) {
    switch(type) {
        case 0: return '男';
        case 1: return '女';
        default: return '未知';
    }
}

// 格式化时间戳
function formatDate(time) {
    let d = time ? new Date(time * 1) : new Date();
    let year = d.getFullYear().toString();
    let month = fix(d.getMonth() + 1, 2);
    let date = fix(d.getDate(), 2);
    
    function fix(num, length) {
        return ('' + num).length < length ? ((new Array(length + 1)).join('0') + num).slice(-length) : '' + num;
    }

    return `${year}-${month}-${date}`;
}
