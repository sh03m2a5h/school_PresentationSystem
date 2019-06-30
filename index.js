const puppeteer = require('puppeteer-core');
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
    const browser = await puppeteer.launch({ headless: false, args: ['--start-fullscreen', '--disable-infobar'], timeout: 15000, executablePath: datas.chromepath });
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

    const ppURLs = await (
        /**
         * OneDrive探索処理
         * @returns {Array<HTMLAnchorElement>}
         */
        async (text) => {
            await page.goto('https://onedrive.live.com');
            await page.waitForSelector('div[data-automationid="ドキュメント"]');
            await page.click('div[data-automationid="ドキュメント"]');
            await page.waitForSelector(`div[data-automationid*="${text}"]`);
            return await page.evaluate((text) => {
                return document.querySelectorAll('.ms-List-surface a.ms-Tile-link')
            }, text);
        })();

    myresize(browser._connection, windowId, datas.display.width, datas.display.height)

    const appFrame = await (
        /**
         * PowerPointスライドショー表示処理
         * @returns {puppeteer.Frame}
         */
        async () => {
            await page.goto(ppURL);
            await page.waitForSelector('iframe[name="wac_frame"]');
            const appFrame = await page.frames().find(f => f.name() === 'wac_frame');
            await appFrame.waitFor(5000);
            await appFrame.waitForSelector(`button[data-automation-id*="View"]`);
            await appFrame.click(`button[data-automation-id="View"]`);
            await appFrame.waitFor(2000);
            await appFrame.waitForSelector(`button[data-automation-id="PlayFromBeginning"]`);
            await appFrame.click(`button[data-automation-id="PlayFromBeginning"]`);
            return appFrame;
        })();

    {   /* スライド */
        await appFrame.evaluate(() => {
            document.querySelector('iframe#SlideShowHostFrame').setAttribute('name', 'SlideShowHostFrame');
        });
        const slideFrame = (await appFrame.childFrames().find(f => f.name() === 'SlideShowHostFrame')).childFrames()[0];
        await slideFrame.waitForSelector('#browserLayerViewId');
        const onControl = async (msg) => {
            switch (msg) {
                case 'next':
                    await appFrame.evaluate(() => { document.querySelector('#buttonNextSlide').click() }).catch(async () => {
                        event.removeListener('control', onControl);
                        await browser.close();
                    });
                    break;
                case 'prev':
                    await appFrame.evaluate(() => { document.querySelector('#buttonPrevSlide').click() }).catch((e) => { console.error(e) });
                    break;
                case 'play':
                    const slide = (() => { let res = null; for (let [key, value] of appFrame._childFrames.entries()) { res = value }; let resb = null; for (let [key, value] of res._childFrames.entries()) { resb = value }; return resb; })();
                    // console.log(slide);
                    await slide.evaluate(() => {
                        document.querySelector('video').play();
                    });
                    break;
                case 'back':
                    event.removeListener('control', onControl);
                    await browser.close();
                    break;
            }
            await appFrame.waitFor(100);
            io.emit('note', await slideFrame.evaluate(() => {
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
