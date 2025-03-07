const { dirname, join, extname } = require('path');
const { promises: fs, existsSync, statSync } = require('fs');
const dotenv = require('dotenv');
const puppeteer = require('puppeteer-core');
const { exec } = require('child_process');
const net = require('net');
const fetch = require('node-fetch');
const ini = require('ini');

// 動態載入 .env
const args = process.argv.slice(2);
const envPath = args.includes('--env') 
    ? args[args.indexOf('--env') + 1] 
    : join(__dirname, '.env');
console.log(`載入 .env 從: ${envPath}`);
if (!existsSync(envPath)) {
    console.log(`警告: .env 文件不存在於 ${envPath}，將使用預設值`);
}
dotenv.config({ path: envPath });
console.log('Loaded environment variables:', process.env);

// 從命令列參數獲取 platform
const platformArg = args.find(arg => arg.startsWith('--platform='));
const platformFromArgs = platformArg ? platformArg.split('=')[1] : null;
const platform = platformFromArgs || process.env.PLATFORM || (process.platform === 'win32' ? 'win' : 'mac');

// 環境變數與預設值
const desktopPath = process.env.DESKTOP_PATH || join(process.cwd(), 'Meeting');
const downloadPath = process.env.DOWNLOAD_PATH || join(process.cwd(), 'Downloads');
const loginWaitTime = parseInt(process.env.LOGIN_WAIT_TIME, 10) || 30000;
const logFilePath = process.env.LOG_FILE_PATH || join(process.cwd(), 'crawler_log.txt');
const chromePath = process.env.CHROME_PATH || (platform === 'win' 
    ? 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe' 
    : '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome');
const basePort = parseInt(process.env.REMOTE_DEBUGGING_PORT, 10) || 9222;
const downloadType = process.env.DOWNLOAD_TYPE || 'WORD';
const downloadTimeout = parseInt(process.env.DOWNLOAD_TIMEOUT, 10) || 300000; // 增加到 5 分鐘

const wait = ms => new Promise(resolve => setTimeout(resolve, ms));

async function log(message) {
    const timestamp = new Date().toISOString();
    const logEntry = `[${timestamp}] ${message}\n`;
    try {
        await fs.appendFile(logFilePath, logEntry, 'utf-8');
        console.log(logEntry.trim());
    } catch (error) {
        console.error(`寫入日誌失敗: ${error.message}`);
    }
}

async function findFreePort(startPort) {
    let port = startPort;
    const maxAttempts = 100;
    for (let i = 0; i < maxAttempts; i++) {
        await log(`檢查端口: ${port}`);
        const available = await checkPort(port);
        if (available) {
            await log(`找到可用端口: ${port}`);
            return port;
        }
        port++;
    }
    throw new Error('無法找到可用端口');
}

function checkPort(port) {
    return new Promise((resolve) => {
        const server = net.createServer();
        server.once('error', (err) => {
            if (err.code === 'EADDRINUSE') resolve(false);
            else resolve(true);
        });
        server.once('listening', () => {
            server.close();
            resolve(true);
        });
        server.listen(port, '127.0.0.1');
    });
}

async function readUrlFile(filePath) {
    try {
        const content = await fs.readFile(filePath, 'utf-8');
        const parsed = ini.parse(content);
        const url = parsed.InternetShortcut.URL;
        if (!url || !url.startsWith('http')) {
            await log(`無效 URL: ${url} 在 ${filePath}`);
            return null;
        }
        return url;
    } catch (error) {
        await log(`讀取 ${filePath} 時發生錯誤: ${error.message}`);
        return null;
    }
}

async function waitForCookies(page, cookieNames, options = {}) {
    const {
        maxAttempts = 5,
        waitTime = loginWaitTime || 30000,
        timeout = 120000
    } = options;

    const startTime = Date.now();
    let attempts = 0;

    while (attempts < maxAttempts && (Date.now() - startTime) < timeout) {
        try {
            const cookies = await page.cookies();
            const hasAllCookies = cookieNames.every(name => 
                cookies.some(cookie => cookie.name === name)
            );

            if (hasAllCookies) {
                await log(`找到所有必要 Cookie: ${cookieNames.join(', ')}`);
                return true;
            }

            await log(`未找到所有必要 Cookie: ${cookieNames.join(', ')}，當前 Cookie: ${JSON.stringify(cookies.map(c => c.name))}`);
            attempts++;
            await log(`第 ${attempts} 次嘗試，剩餘 ${maxAttempts - attempts} 次，等待 ${waitTime / 1000} 秒...`);
            await wait(waitTime);
        } catch (error) {
            await log(`檢查 Cookie 時發生錯誤: ${error.message}`);
            attempts++;
            await wait(waitTime);
        }
    }

    await log(`未能在 ${timeout / 1000} 秒內找到所有必要 Cookie: ${cookieNames.join(', ')}`);
    return false;
}

async function waitForDownload(downloadPath, timeout = downloadTimeout, startTime = Date.now()) {
    const endTime = startTime + timeout;
    let downloadedFile = null;

    await log(`開始等待下載，超時時間: ${timeout / 1000} 秒`);
    while (Date.now() < endTime) {
        const files = await fs.readdir(downloadPath);
        const targetExt = downloadType.toUpperCase() === 'PDF' ? '.pdf' : '.docx';

        const fileStats = await Promise.all(
            files
                .filter(file => 
                    !file.endsWith('.crdownload') && 
                    !file.startsWith('.') && 
                    file.endsWith(targetExt)
                )
                .map(async file => {
                    const stats = await fs.stat(join(downloadPath, file));
                    return { file, mtime: stats.mtimeMs, size: stats.size };
                })
        );

        // 只選取下載開始後修改的文件
        const newFiles = fileStats.filter(f => f.mtime >= startTime);
        if (newFiles.length > 0) {
            downloadedFile = newFiles[0].file; // 選第一個新文件
            const filePath = join(downloadPath, downloadedFile);
            await log(`發現新下載檔案: ${filePath}`);

            let previousSize = -1;
            let stableCount = 0;
            const maxStableCount = 5;
            const checkInterval = 3000;

            while (Date.now() < endTime) {
                const stats = await fs.stat(filePath);
                const currentSize = stats.size;
                if (currentSize === previousSize && currentSize > 0) {
                    stableCount++;
                    if (stableCount >= maxStableCount) {
                        await wait(2000);
                        return filePath;
                    }
                } else {
                    stableCount = 0;
                }
                previousSize = currentSize;
                await wait(checkInterval);
            }
        }
        await wait(1000);
    }
    throw new Error('下載超時，未找到新下載檔案');
}

async function startChrome(port) {
    return new Promise((resolve, reject) => {
        const userDataDir = platform === 'win' ? `C:\\Temp\\chrome-remote-${port}` : `/tmp/chrome-remote-${port}`;
        let chromeCommand = platform === 'mac' 
            ? `"${chromePath}" --remote-debugging-port=${port} --no-first-run --no-default-browser-check --start-fullscreen --user-data-dir="${userDataDir}" --no-sandbox --disable-setuid-sandbox`
            : `"${chromePath}" --remote-debugging-port=${port} --no-first-run --no-default-browser-check --start-fullscreen --user-data-dir="${userDataDir}" --no-sandbox --disable-setuid-sandbox`;

        if (!existsSync(chromePath)) {
            log(`錯誤: Chrome 可執行文件不存在於 ${chromePath}`);
            reject(new Error(`Chrome 可執行文件不存在於 ${chromePath}`));
            return;
        }

        log(`啟動 Chrome: ${chromeCommand}`);
        const chromeProcess = exec(chromeCommand, { shell: true }, (error, stdout, stderr) => {
            if (error) {
                log(`Chrome 啟動失敗: ${error.message}`);
                reject(error);
            } else {
                log(`Chrome 啟動成功，PID: ${chromeProcess.pid}`);
            }
        });
        chromeProcess.on('error', reject);
        chromeProcess.unref();
        setTimeout(() => resolve(chromeProcess), 10000);
    });
}

async function crawlMeetingUrls() {
    let chromeProcess;
    let isLoggedIn = false;

    try {
        await fs.mkdir(downloadPath, { recursive: true });
        await fs.access(downloadPath, fs.constants.W_OK);
        await log(`下載路徑 ${downloadPath} 可寫`);

        const port = await findFreePort(basePort);
        await log(`使用遠端除錯端口: ${port}`);
        chromeProcess = await startChrome(port);

        const wsUrl = `http://127.0.0.1:${port}/json/version`;
        await log(`嘗試連接 WebSocket: ${wsUrl}`);

        let attempts = 0;
        const maxAttempts = 10;
        let wsData;
        while (attempts < maxAttempts) {
            try {
                const response = await fetch(wsUrl);
                if (!response.ok) throw new Error(`WebSocket 不可用: ${response.statusText}`);
                wsData = await response.json();
                break;
            } catch (error) {
                attempts++;
                await log(`第 ${attempts} 次嘗試連接失敗: ${error.message}`);
                if (attempts === maxAttempts) throw error;
                await wait(3000);
            }
        }
        await log(`WebSocket 連接成功: ${wsData.webSocketDebuggerUrl}`);

        const browser = await puppeteer.connect({
            browserURL: `http://127.0.0.1:${port}`,
            defaultViewport: null
        });
        await log('成功連接到 Chrome 實例');

        const files = await fs.readdir(desktopPath);
        const urlFiles = files.filter(file => extname(file).toLowerCase() === '.url');
        if (urlFiles.length === 0) {
            await log('未找到 .url 文件');
            return;
        }

        const requiredCookies = ['session', 'session_list'];

        for (const file of urlFiles) {
            const filePath = join(desktopPath, file);
            const url = await readUrlFile(filePath);

            if (url) {
                await log(`正在處理: ${file} - ${url}`);
                const page = await browser.newPage();
                const client = await page.target().createCDPSession();

                await client.send('Page.setDownloadBehavior', {
                    behavior: 'allow',
                    downloadPath: downloadPath
                });
                await log(`下載路徑已設置為: ${downloadPath}`);

                try {
                    await page.setViewport({ width: 1920, height: 1080 });
                    await log(`開始導航到: ${url}`);
                    await page.goto(url, { waitUntil: 'networkidle0', timeout: 60000 });

                    if (!isLoggedIn) {
                        await log(`檢查必要 Cookie: ${requiredCookies.join(', ')}`);
                        const cookiesReady = await waitForCookies(page, requiredCookies, {
                            maxAttempts: 5,
                            waitTime: loginWaitTime,
                            timeout: 120000
                        });
                        if (!cookiesReady) {
                            await log('未能在指定時間內找到必要 Cookie，程式中止');
                            await page.close();
                            return;
                        }
                        isLoggedIn = true;
                        await log('已登入，後續 URL 將直接使用 Cookie');
                    } else {
                        await log('已登入，使用現有 Cookie 繼續導航');
                    }

                    await log(`導航完成: ${url}`);

                    // 點擊「suite-more-menu」按鈕
                    const menuButton = await page.waitForSelector('div.suite-more-menu > button', { visible: true, timeout: 15000 });
                    if (menuButton) {
                        await menuButton.click();
                        await log('已點擊 div.suite-more-menu > button');
                        await wait(1000);
                    } else {
                        await log('未找到 div.suite-more-menu > button');
                        await page.screenshot({ path: `debug_${file}_menu.png` });
                        continue;
                    }

                    // 移動到「下載為」元素
                    const downloadElement = await page.evaluateHandle(() => {
                        const spans = Array.from(document.querySelectorAll('span'));
                        return spans.find(span => span.textContent.trim().includes('下載為'));
                    });
                    if (downloadElement) {
                        const box = await downloadElement.boundingBox();
                        if (box) {
                            await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
                            await log(`已移動到「下載為」元素中心`);
                            await wait(2000);
                        } else {
                            await log('無法獲取「下載為」元素的boundingBox');
                            await page.screenshot({ path: `debug_${file}_download.png` });
                            continue;
                        }
                    } else {
                        await log('未找到「下載為」元素');
                        await page.screenshot({ path: `debug_${file}_download.png` });
                        continue;
                    }

                    // 選擇下載格式（PDF 或 Word）
                    const targetText = downloadType.toUpperCase() === 'PDF' ? 'PDF' : 'Word';
                    const formatElement = await page.evaluateHandle((text) => {
                        const spans = Array.from(document.querySelectorAll('span'));
                        return spans.find(span => span.textContent.trim().includes(text));
                    }, targetText);
                    if (formatElement) {
                        const box = await formatElement.boundingBox();
                        if (box) {
                            await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
                            await log(`已點擊「${targetText}」元素`);
                            await wait(2000);
                        } else {
                            await log(`無法獲取「${targetText}」元素的boundingBox`);
                            await page.screenshot({ path: `debug_${file}_format.png` });
                            continue;
                        }
                    } else {
                        await log(`未找到「${targetText}」元素`);
                        await page.screenshot({ path: `debug_${file}_format.png` });
                        continue;
                    }

                    // 點擊「匯出」按鈕並等待下載
                    const exportButtonHandle = await page.evaluateHandle(() => {
                        const buttons = Array.from(document.querySelectorAll('button'));
                        return buttons.find(btn => btn.textContent.trim().includes('匯出'));
                    });
                    const exportButton = exportButtonHandle.asElement();
                    const downloadStartTime = Date.now(); // 記錄下載開始時間
                    if (exportButton) {
                        await exportButton.click();
                        await log('已點擊「匯出」按鈕');
                        await page.screenshot({ path: `debug_${file}_after_export.png` });

                        // 處理可能的確認對話框
                        await page.evaluate(() => {
                            const confirmButton = document.querySelector('button[class*="confirm"], button[class*="ok"]');
                            if (confirmButton) confirmButton.click();
                        });
                        await log('已檢查並處理可能的確認對話框');

                        const downloadedFile = await waitForDownload(downloadPath, downloadTimeout, downloadStartTime);                     
                        await log(`檔案已下載至: ${downloadedFile}`);
                        
                        const newFileName = `${file.replace('.url', '')}_${Date.now()}${extname(downloadedFile)}`;
                        const newFilePath = join(downloadPath, newFileName);
                        await fs.rename(downloadedFile, newFilePath);
                        await log(`檔案已重新命名為: ${newFilePath}`);
                    } else {
                        await log('未找到「匯出」按鈕');
                        await page.screenshot({ path: `debug_${file}_export.png` });
                    }
                } catch (error) {
                    await log(`處理 ${url} 時發生錯誤: ${error.message}`);
                    if (error.name === 'TimeoutError') {
                        await log(`頁面導航或元素等待超時，跳過此 URL: ${url}`);
                    } else if (error.message.includes('net::ERR')) {
                        await log(`網路錯誤，跳過此 URL: ${url}`);
                    } else {
                        await log(`其他錯誤，詳情: ${error.stack}`);
                    }
                } finally {
                    try {
                        await page.close();
                        await log(`頁面已關閉: ${url}`);
                    } catch (closeError) {
                        await log(`關閉頁面時發生錯誤: ${closeError.message}`);
                    }
                }
            }
        }

        await browser.disconnect();
        await log('所有網頁處理完成！');
    } catch (error) {
        await log(`程式執行發生錯誤: ${error.message}`);
        console.error('程式執行錯誤:', error);
    } finally {
        if (chromeProcess) {
            chromeProcess.kill();
            await log('Chrome 進程已清理');
        }
    }
}

crawlMeetingUrls().catch(console.error);