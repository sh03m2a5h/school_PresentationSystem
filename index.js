const puppeteer = require('puppeteer');
const datas = require('./data.js');
const execSync = require('child_process').execSync;
const EventEmitter = require('events').EventEmitter;

var app = require('express')();
var http = require('http').createServer(app);
var io = require('socket.io')(http);

const event = new EventEmitter;
event.setMaxListeners(10);

app.get('/', function (req, res) {
    res.sendFile(__dirname + '/index.html');
});

io.on('connection', function (socket) {
    console.log('a user connected');
    socket.on('message', (msg) => {
        console.log(msg);
        event.emit('control', msg);
    });
});

http.listen(3000, function () {
    console.log('listening on *:3000');
});

(async () => {
    const browser = await puppeteer.launch({ headless: false, args: ['--start-fullscreen', '--disable-infobar'], timeout: 15000 });
    const page = await browser.newPage(); const { targetInfos: [{ targetId }] } = await browser._connection.send(
        'Target.getTargets'
    )

    const { windowId } = await browser._connection.send(
        'Browser.getWindowForTarget', {
            targetId
        }
    )

    const myresize = async (c, wid, w, h) => {
        await page.setViewport({ width: w, height: h })
        await c.send('Browser.setWindowBounds', {
            bounds: {
                height: h,
                width: w
            },
            windowId: wid
        })
    }
    await page.setViewport({ width: 1280, height: 720 });
    await page.goto('https://office.com');

    {
        const disps = execSync('python close_imbotbar.py').toString();
        const json = JSON.parse(disps);
        datas.display.width = json.width;
        datas.display.height = json.height;
        console.log(datas.display);
    }


    {   /* ログイン処理 */
        await page.waitForSelector('#hero-banner-sign-in-to-office-365-link');
        await page.click('#hero-banner-sign-in-to-office-365-link');
        await page.waitForSelector('input[type="email"]');
        await page.type('input[type="email"]', datas.user.id);
        await page.click('#idSIButton9');
        await page.waitFor(2000);
        await page.waitForSelector('input[type="password"]');
        await page.type('input[type="password"]', datas.user.passwd);
        await page.click('#idSIButton9');
        await page.waitForSelector('#idBtn_Back');
        await page.click('#idBtn_Back');
    }

    const ppURL = await (
        /**
         * OneDrive探索処理
         * @returns {string}
         */
        async (text) => {
            await page.goto('https://onedrive.live.com');
            await page.waitForSelector('div[data-automationid="ドキュメント"]');
            await page.click('div[data-automationid="ドキュメント"]');
            await page.waitForSelector(`div[data-automationid*="${text}"]`);
            return await page.evaluate((text) => {
                return document.querySelector(`div[data-automationid*="${text}"] > a`).href
            }, text);
        })('Quantum Computer.pptx');

    myresize(browser._connection, windowId, datas.display.width, datas.display.height)


    const frame = await (
        /**
         * PowerPointスライドショー表示処理
         * @returns {puppeteer.Frame}
         */
        async () => {
            await page.goto(ppURL);
            await page.waitForSelector('iframe[name="wac_frame"]');
            const frame = await page.frames().find(f => f.name() === 'wac_frame');
            await frame.waitFor(5000);
            await frame.waitForSelector(`button[data-automation-id*="View"]`);
            await frame.click(`button[data-automation-id="View"]`);
            await frame.waitFor(2000);
            await frame.waitForSelector(`button[data-automation-id="PlayFromBeginning"]`);
            await frame.click(`button[data-automation-id="PlayFromBeginning"]`);
            return frame;
        })();

    {   /* スライド */
        await frame.evaluate(() => {
            document.querySelector('iframe#SlideShowHostFrame').setAttribute('name', 'SlideShowHostFrame');
        });
        const Slide = (await frame.childFrames().find(f => f.name() === 'SlideShowHostFrame')).childFrames()[0];
        await Slide.waitForSelector('#browserLayerViewId');
        const onControl = async (msg) => {
            switch (msg) {
                case 'next':
                    await frame.evaluate(() => { document.querySelector('#buttonNextSlide').click() }).catch(async () => {
                        event.removeListener('control', onControl);
                        await browser.close();
                    });
                    break;
                case 'prev':
                    await frame.evaluate(() => { document.querySelector('#buttonPrevSlide').click() });
                    break;
            }
            await frame.waitFor(100);
            io.emit('note', await Slide.evaluate(() => {
                var result = '';
                document.querySelectorAll('#OutlineView ul > *').forEach((elements) => {
                    result += elements.innerText + '\n';
                });
                return result;
            }));
        };
        event.addListener('control', onControl);
    }
})();
