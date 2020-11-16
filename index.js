const os = require('os');
const path = require('path')
const fs = require('fs')
const puppeteer = require('puppeteer');
const xlsx = require('node-xlsx')

const xlsxPath = xlsx.parse('khzl.xlsx');
let data = xlsxPath[0].data, page
/**
 * 打开浏览器
 */
async function initBrowser() {
    const executablePath = await getExecutablePath();
    consolea(executablePath)
    const options = {
        headless: false,
        ignoreDefaultArgs: ['--enable-automation'],//隐藏Chrome正受到自动测试软件的控制
        args: [
            '--disable-web-security', // 禁用同源策略(允许跨域)
            '--disable-translate',	//禁用翻译提示
            '--no-sandbox',//禁用沙箱
            '--disable-gpu',//禁用GUP
            '--ignore-certificate-errors', //忽略证书错误
            '--ignore-certificate-errors-spki-list',//忽略证书错误公钥列表
            '--allow-running-insecure-content', //允许运行不安全的网站
            '--disable-setuid-sandbox',
            '--disable-background-timer-throttling',//禁用后台标签性能限制
            '--disable-site-isolation-trials',//禁用网站隔离
            '--disable-renderer-backgrounding',//禁止降低后台网页进程的渲染优先级
            // '--start-maximized',//以最大化启动
            '--test-type',//可以隐藏提示
            '--suppress-message-center-popups',//隐藏消息
            '--no-default-browser-check',//禁用检查是否为默认浏览器
        ],
        executablePath: executablePath,
        defaultViewport: { width: 1920, height: 1080 } //默认窗口大小
        //devtools: true
    }
    const browser = await puppeteer.launch(options);
    page = await browser.newPage();
    await page.goto('https://www.tianyancha.com/search?key=%E7%A7%91%E6%8A%80%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8');
    // 请先扫码登录
    await page.waitFor(15000)   
    try{
        await query(browser)
    }catch(e){
        console.log(e)
        var buffer = xlsx.build(xlsxPath);
        fs.writeFile('result.xlsx',buffer,function(e){
            if(e) throw e
        })
    }
    
}

async function query(browser){
    for(let i = 0; i < data.length; i++){
        const newPagePromise = new Promise(resolve => browser.once('targetcreated', target => resolve(target.page())));
        await page.waitForXPath('//input[@id="header-company-search"]')
        let inputSearch = await page.$x('//input[@id="header-company-search"]')
        await set(inputSearch[0],data[i][1])
        let divSearch =  await page.$x('//input[@id="header-company-search"]//parent::div//parent::form/following-sibling::div[1]')
        await divSearch[0].click()
        await page.waitFor(3000)
        await page.waitForXPath("//div[contains(@class,'result-list sv-search-container')]")
        let company = await page.$x(`//div[contains(@class,'result-list sv-search-container')]/div[1]//div[@class='content']//div[@class='header']//a`)
        await company[0].click()
        const newPage = await newPagePromise; //利用targetcreated得到page
        await newPage.bringToFront() //切换到newPage
        const url = newPage.url()
        await newPage.close(); //关闭newPage
        const copyNewPage = await browser.newPage()
        await setPage(copyNewPage)
        await copyNewPage.goto(url, {waitUntil: 'networkidle2'})
        try{
            await copyNewPage.waitForXPath('//div[@id="_container_baseInfo"]')
            let zczbEle = await copyNewPage.$x(`//div[@id="_container_baseInfo"]//td[contains(text(),'注册资本')]/following-sibling::td[1]`)
            let zczb = await (await zczbEle[0].getProperty('innerText')).jsonValue()
            let cbrsEle = await copyNewPage.$x(`//div[@id="_container_baseInfo"]//td[contains(text(),'参保人数')]/following-sibling::td[1]`)
            let cbrs = await (await cbrsEle[0].getProperty('innerText')).jsonValue()
            let zcdzEle = await copyNewPage.$x(`//div[@id="_container_baseInfo"]//td[contains(text(),'注册地址')]/following-sibling::td[1]`)
            let zcdz = await (await zcdzEle[0].getProperty('innerText')).jsonValue()
            zcdz = zcdz.replace('附近公司','')
            data[i].push(zczb,cbrs,zcdz)
        }catch(error){
            consolea(error)
        }

        consolea(data[i])
        await page.bringToFront()
        await copyNewPage.close()
        await page.reload({
            waitUntil: "networkidle0"
        })
    }
    var buffer = xlsx.build(xlsxPath);
    fs.writeFile('result.xlsx',buffer,function(e){
        if(e) throw e
    })
    
}

async function set(ele,value){
    await ele.focus()
    await inputClear()
    await ele.type(`${value}`)
    console.log(`${value}`)
}

async function inputClear() {
    await page.keyboard.down('Control');
    await page.keyboard.down('KeyA');
    await page.keyboard.press('Backspace'); //Backspace/Delete
    await page.keyboard.up('Control');
    await page.keyboard.up('KeyA');
}

//获取chrome路径
async function getExecutablePath() {
    const executablePath1 = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe'
    const executablePath2 = 'C:/Users/Administrator/AppData/Local/Google/Chrome/Application/chrome.exe'
    const executablePath3 = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
    const executablePath4 = path.join(os.homedir(), '/AppData/Local/Google/Chrome/Application/chrome.exe')
    const executablePath5 = '/opt/google/chrome/google-chrome'//linux
    const executablePath6 = 'C:/Program Files/Google/Chrome/Application/chrome.exe'

    if (await fs.existsSync(executablePath1)) {
        return executablePath1
    } else if (await fs.existsSync(executablePath2)) {
        return executablePath2
    } else if (await fs.existsSync(executablePath3)) {
        return executablePath3
    } else if (await fs.existsSync(executablePath4)) {
        return executablePath4
    } else if (await fs.existsSync(executablePath5)) {
        return executablePath5
    } else if (await fs.existsSync(executablePath6)) {
        return executablePath6
    } else {
        return ''
    }

}

async function setPage(page) {
    //有些网站屏蔽chrome浏览器，修改为Firefox
    //await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:71.0) Gecko/20100101 Firefox/71.0');
    //webdriver
    await page.evaluateOnNewDocument(() => {
        Object.defineProperty(navigator, 'webdriver', {
            get: () => false,
        });
    });
    //chrome
    await page.evaluateOnNewDocument(() => {
        // We can mock this in as much depth as we need for the test.
        window.navigator.chrome = {
            runtime: {},
            // etc.
        };
    });
    //permissions
    await page.evaluateOnNewDocument(() => {
        const originalQuery = window.navigator.permissions.query;
        window.navigator.permissions.query = (parameters) => (
            parameters.name === 'notifications' ?
                Promise.resolve({ state: Notification.permission }) :
                originalQuery(parameters)
        );
    });
    //plugins
    await page.evaluateOnNewDocument(() => {
        Object.defineProperty(navigator, 'plugins', {
            // This just needs to have `length > 0` for the current test,
            // but we could mock the plugins too if necessary.
            get: () => [1, 2, 3, 4, 5],
        });
    });
    //languages
    await page.evaluateOnNewDocument(() => {
        Object.defineProperty(navigator, 'languages', {
            get: () => ['en-US', 'en'],
        });
    });
    //设置Viewport
    const dimensions = await page.evaluate(() => {
        return {
            deviceScaleFactor: window.devicePixelRatio,
            width: window.screen.width,
            height: window.screen.height
        };
    });
    await page.setViewport(dimensions);
    
}

function consolea(n){
    console.log('*****')
    console.log(n)
    console.log('*****')
}

initBrowser()
