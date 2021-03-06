const puppeteer = require('puppeteer-core');
const datas = require('./data.js');
const execSync = require('child_process').execSync;
const EventEmitter = require('events').EventEmitter;
const qrcode = require('qrcode');
const uuid = require('uuid');
const os = require('os');
const app = require('express')();
const http = require('http').createServer(app);
const io = require('socket.io')(http);

const event = new EventEmitter;
event.setMaxListeners(10);

class Otp {
    constructor() {
        this.key = '';
        /**
         * @type {SocketIO.Socket}
         */
        this._authorized_client = null;
    }
    newKey() {
        if (this._authorized_client) {
            this._authorized_client.disconnect();
        }
        this.key = uuid.v1();
        qrcode.toFile('./qr.png', 'http://' + addresses[0] + ':3000/' + this.key);
    }
    /**
     * 
     * @param {SocketIO.Socket} client 
     */
    setOuthorizedClient(client) {
        this._authorized_client = client;
    }
    getOuthorizedClient() {
        return this._authorized_client;
    }
}
const otp = new Otp();

const interfacess = os.networkInterfaces()
const addresses = [];
for (const itf in interfacess)
    for (const interface of interfacess[itf])
        if (interface.family === 'IPv4' && !interface.internal)
            addresses.push(interface.address);

app.get('/:uuid', function (req, res) {
    if (req.params.uuid == 'qr.png') {
        console.log(req.ip);
        if (req.ip.includes('::1') || req.ip.includes('127.0.0.1'))
            res.sendFile(__dirname + '/qr.png');
        else
            res.sendStatus(404);

    } else if (req.params.uuid == otp.key) {
        res.sendFile(__dirname + '/index.html');
    } else {
        res.sendStatus(404)
    }
});

io.on('connection', function (socket) {
    console.log('a user connected');
    socket.on('authorize', (uuid) => {
        if (uuid == otp.key) {
            otp.setOuthorizedClient(socket);
            socket.on('message', (msg) => {
                console.log(msg);
                event.emit('control', msg);
            });
            event.emit('connect');
        }
    })
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

    {
        const disps = execSync('python close_imbotbar.py').toString();
        console.log(disps);
        const json = JSON.parse(disps);
        await page.setViewport({ width: json.width, height: json.height });
    }

    
    await page.goto('https://office.com');

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
         * @returns {Array<{name:string,url:string}>}
         */
        async (text) => {
            await page.goto('https://onedrive.live.com');
            await page.waitForSelector('div[data-automationid="ドキュメント"]');
            await page.click('div[data-automationid="ドキュメント"]');
            await page.waitForSelector(`div.ms-Tile`);
            return await page.evaluate(() => {
                const files = [];
                for (const elem of document.querySelectorAll('div.ms-Tile')) {
                    files.push({ name: elem.getAttribute('data-automationid'), url: elem.querySelector('a').href });
                }
                return files;
            });
        })();

    console.log(ppURLs);
    // await new Promise((resolve,reject)=>{
    //     
    // });

    //myresize(browser._connection, windowId, datas.display.width, datas.display.height)

    for (const ppURL of ppURLs) {
        otp.newKey();
        await new Promise((resolve, reject) => {
            page.goto("http://localhost:3000/qr.png");
            event.once('connect', () => {
                resolve();
            });
        });

        /**
         * @type {puppeteer.Frame}
         */
        const appFrame = await new Promise(async (resolve, reject) => {
                await page.goto(ppURL.url);
                await page.waitForSelector('iframe[name="wac_frame"]');
                const appFrame = await page.frames().find(f => f.name() === 'wac_frame');
                await appFrame.waitForSelector(`button[data-automation-id="View"]`);
                const loop = (() => {
                        appFrame.click(`button[data-automation-id="View"]`).then(() => {
                            return appFrame.waitForSelector(`button[data-automation-id="PlayFromBeginning"]`, { timeout: 2000 });
                        }).then(() => {
                            return appFrame.click(`button[data-automation-id="PlayFromBeginning"]`);
                        }).then(()=>{
                            resolve(appFrame);
                        }).catch(() => {
                            loop();
                        });
                    });
                loop();
            });

        await new Promise(async (resolve, reject) => {   /* スライド */
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
                            resolve();
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
                        resolve();
                        break;
                }
                await appFrame.waitFor(100);
                otp.getOuthorizedClient().emit('note', await slideFrame.evaluate(() => {
                    var result = '';
                    document.querySelectorAll('#OutlineView ul > *').forEach((elements) => {
                        result += elements.innerText + '\n';
                    });
                    return result;
                }).catch(() => { }));
            };
            event.addListener('control', onControl);
        });
    }
    await browser.close();
    io.close();
    http.close();
})();
