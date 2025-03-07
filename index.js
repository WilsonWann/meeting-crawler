const { dirname, join, extname } = require('path');
const { promises: fs, existsSync } = require('fs');
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

// 從命令列參數獲取 platform（若有）
const platformArg = args.find(arg => arg.startsWith('--platform='));
const platformFromArgs = platformArg ? platformArg.split('=')[1] : null;

// 決定平台：命令列 > .env > 預設
const platform = platformFromArgs || process.env.PLATFORM || (process.platform === 'win32' ? 'win' : 'mac');

// 環境變數與預設值
const desktopPath = process.env.DESKTOP_PATH || join(process.cwd(), 'Meeting');
const downloadPath = process.env.DOWNLOAD_PATH || join(process.cwd(), 'Downloads');
const loginWaitTime = parseInt(process.env.LOGIN_WAIT_TIME, 10) || 30000; // 預設 30 秒
console.log('Desktop Path:', desktopPath);
console.log('Download Path:', downloadPath);
console.log('Login Wait Time:', loginWaitTime);

const logFilePath = process.env.LOG_FILE_PATH || join(process.cwd(), 'crawler_log.txt');
const chromePath = process.env.CHROME_PATH || (platform === 'win' 
    ? 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe' 
    : '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome');
const basePort = parseInt(process.env.REMOTE_DEBUGGING_PORT, 10) || 9222;
const downloadType = process.env.DOWNLOAD_TYPE || 'WORD';
console.log("🚀 ~ downloadType:", downloadType)
const downloadTimeout = parseInt(process.env.DOWNLOAD_TIMEOUT, 10) || 120000;

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
            if (err.code === 'EADDRINUSE') {
                resolve(false);
            } else {
                resolve(true);
            }
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

async function waitForDownload(downloadPath, timeout = downloadTimeout) {
    const startTime = Date.now();
    let downloadedFile = null;

    // 等待檔案出現
    while (Date.now() - startTime < timeout) {
        const files = await fs.readdir(downloadPath);
        downloadedFile = files.find(file => !file.endsWith('.crdownload')); // 排除 .crdownload 檔案
        if (downloadedFile) {
            break;
        }
        await wait(1000);
    }

    if (!downloadedFile) {
        throw new Error('下載超時，未找到下載檔案');
    }

    const filePath = join(downloadPath, downloadedFile);
    await log(`發現下載檔案: ${filePath}`);

    // 檢查檔案大小是否穩定
    let previousSize = -1;
    let stableCount = 0;
    const maxStableCount = 3; // 連續 3 次大小不變認為下載完成
    const checkInterval = 2000; // 每 2 秒檢查一次

    while (Date.now() - startTime < timeout) {
        const stats = statSync(filePath);
        const currentSize = stats.size;

        await log(`檢查檔案大小: ${currentSize} bytes`);

        if (currentSize === previousSize) {
            stableCount++;
            if (stableCount >= maxStableCount) {
                await log(`檔案大小穩定，下載完成: ${filePath}`);
                return filePath;
            }
        } else {
            stableCount = 0;
        }

        previousSize = currentSize;
        await wait(checkInterval);
    }

    throw new Error('下載超時，檔案大小未穩定');
}

async function startChrome(port) {
    return new Promise((resolve, reject) => {
        const userDataDir = platform === 'win' ? `C:\\Temp\\chrome-remote-${port}` : `/tmp/chrome-remote-${port}`;
        const isHeadless = args.includes('--headless=false') ? false : true; // 預設 headless: true，可通過 --headless=false 切換
        let chromeCommand;
        if (platform === 'mac') {
            chromeCommand = `"${chromePath}" --remote-debugging-port=${port} --no-first-run --no-default-browser-check --start-fullscreen --user-data-dir="${userDataDir}" --no-sandbox --disable-setuid-sandbox --headless=${isHeadless}`;
        } else if (platform === 'win') {
            chromeCommand = `"${chromePath}" --remote-debugging-port=${port} --no-first-run --no-default-browser-check --start-fullscreen --user-data-dir="${userDataDir}" --no-sandbox --disable-setuid-sandbox --headless=${isHeadless}`;
        } else {
            reject(new Error(`不支持的平台: ${platform}`));
            return;
        }
        
        if (!existsSync(chromePath)) {
            log(`錯誤: Chrome 可執行文件不存在於 ${chromePath}`);
            reject(new Error(`Chrome 可執行文件不存在於 ${chromePath}`));
            return;
        }
        
        log(`啟動 Chrome: ${chromeCommand}`);
        const chromeProcess = exec(
            chromeCommand,
            { shell: true },
            (error, stdout, stderr) => {
                if (error) {
                    log(`Chrome 啟動失敗: ${error.message}`);
                    log(`腳本輸出: ${stdout}`);
                    log(`錯誤輸出: ${stderr}`);
                    reject(error);
                } else {
                    log(`Chrome 啟動成功，PID: ${chromeProcess.pid}`);
                    log(`腳本輸出: ${stdout}`);
                }
            }
        );
        chromeProcess.on('error', (err) => {
            log(`Chrome 進程錯誤: ${err.message}`);
            reject(err);
        });
        chromeProcess.unref();
        setTimeout(() => resolve(chromeProcess), 10000);
    });
}

async function waitForCookies(page, cookieNames) {
    let attemptsLeft = 3; // 最多重試 3 次
    while (attemptsLeft > 0) {
        const cookies = await page.cookies();
        const hasAllCookies = cookieNames.every(name => cookies.some(cookie => cookie.name === name));
        if (hasAllCookies) {
            await log(`找到所有必要 Cookie: ${cookieNames.join(', ')}`);
            return true;
        }
        await log(`未找到所有必要 Cookie: ${cookieNames.join(', ')}，剩餘 ${attemptsLeft} 次重試，等待 ${loginWaitTime / 1000} 秒...`);
        await wait(loginWaitTime);
        attemptsLeft--;
    }
    await log(`未能在指定時間內找到所有必要 Cookie: ${cookieNames.join(', ')}`);
    return false;
}

async function crawlMeetingUrls() {
    let chromeProcess;
    let isLoggedIn = false; // 記錄是否已登入
    try {
        await fs.mkdir(downloadPath, { recursive: true });

        const port = await findFreePort(basePort);
        await log(`使用遠端除錯端口: ${port}`);
        chromeProcess = await startChrome(port);

        const wsUrl = `http://127.0.0.1:${port}/json/version`;
        await log(`嘗試連接 WebSocket: ${wsUrl}`);
        
        let attempts = 0;
        const maxAttempts = 10;
        let response, wsData;
        while (attempts < maxAttempts) {
            try {
                response = await fetch(wsUrl);
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

                await page._client().send('Page.setDownloadBehavior', {
                    behavior: 'allow',
                    downloadPath: downloadPath
                });

                try {
                    await page.setViewport({ width: 1920, height: 1080 });
                    await log(`開始導航到: ${url}`);
                    await page.goto(url, { waitUntil: 'networkidle0', timeout: 60000 }); // 等待網絡空閒

                    // 若尚未登入，檢查 Cookie
                    if (!isLoggedIn) {
                        await log(`檢查必要 Cookie: ${requiredCookies.join(', ')}`);
                        const cookiesReady = await waitForCookies(page, requiredCookies);
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

                    // 1. 點擊 div.suite-more-menu > button
                    const menuButton = await page.waitForSelector('div.suite-more-menu > button', { visible: true, timeout: 15000 });
                    if (menuButton) {
                        await menuButton.click();
                        await log('已點擊 div.suite-more-menu > button');
                        await wait(1000); // 額外等待 1 秒，確保下拉菜單出現
                    } else {
                        await log('未找到可見的 div.suite-more-menu > button');
                        await page.screenshot({ path: `debug_${file}_menu.png` });
                        await log(`已生成調試截圖: debug_${file}_menu.png`);
                        continue; // 跳過此 URL
                    }

                    // 2. 移動滑鼠到包含「下載為」的元素中心
                    const downloadElement = await page.evaluateHandle(() => {
                        const spans = Array.from(document.querySelectorAll('span'));
                        return spans.find(span => span.textContent.trim().includes('下載為'));
                    });
                    if (downloadElement) {
                        const box = await downloadElement.boundingBox();
                        if (box) {
                            const centerX = box.x + box.width / 2;
                            const centerY = box.y + box.height / 2;
                            await page.mouse.move(centerX, centerY);
                            await log(`已移動到包含「下載為」的元素中心座標 (${centerX}, ${centerY})`);
                        } else {
                            await log('無法獲取包含「下載為」的元素的boundingBox');
                            await page.screenshot({ path: `debug_${file}_download.png` });
                            await log(`已生成調試截圖: debug_${file}_download.png`);
                            continue;
                        }
                    } else {
                        await log('未找到包含「下載為」的元素');
                        await page.screenshot({ path: `debug_${file}_download.png` });
                        await log(`已生成調試截圖: debug_${file}_download.png`);
                        continue;
                    }
                    await wait(2000);

                    // 3. 根據 DOWNLOAD_TYPE 查找「PDF」或「Word」元素，移動滑鼠並點擊
                    const targetText = downloadType.toUpperCase() === 'PDF' ? 'PDF' : 'Word';
                    const formatElement = await page.evaluateHandle((text) => {
                        const spans = Array.from(document.querySelectorAll('span'));
                        return spans.find(span => span.textContent.trim().includes(text));
                    }, targetText);
                    if (formatElement) {
                        const box = await formatElement.boundingBox();
                        if (box) {
                            const centerX = box.x + box.width / 2;
                            const centerY = box.y + box.height / 2;
                            await page.mouse.move(centerX, centerY);
                            await log(`已移動到「${targetText}」元素中心座標 (${centerX}, ${centerY})`);
                            await page.mouse.click(centerX, centerY);
                            await log(`已點擊「${targetText}」元素`);
                        } else {
                            await log(`無法獲取「${targetText}」元素的boundingBox`);
                            await page.screenshot({ path: `debug_${file}_format.png` });
                            await log(`已生成調試截圖: debug_${file}_format.png`);
                            continue;
                        }
                    } else {
                        await log(`未找到包含「${targetText}」的元素`);
                        await page.screenshot({ path: `debug_${file}_format.png` });
                        await log(`已生成調試截圖: debug_${file}_format.png`);
                        continue;
                    }
                    await wait(2000);

                    // 4. 點擊包含「匯出」的按鈕
                    await page.waitForSelector('button', { visible: true, timeout: 15000 }); // 確保有按鈕可見
                    const exportButtonHandle = await page.evaluateHandle(() => {
                        const buttons = Array.from(document.querySelectorAll('button'));
                        return buttons.find(btn => btn.textContent.trim().includes('匯出'));
                    });

                    // 檢查 exportButtonHandle 是否有效
                    const exportButton = exportButtonHandle.asElement();
                    if (exportButton) {
                        await exportButton.click();
                        await log('已點擊「匯出」按鈕');

                        const downloadedFile = await waitForDownload(downloadPath);
                        await log(`檔案已下載至: ${downloadedFile}`);

                        const newFileName = `${file.replace('.url', '')}_${Date.now()}${extname(downloadedFile)}`;
                        const newFilePath = join(downloadPath, newFileName);
                        await fs.rename(downloadedFile, newFilePath);
                        await log(`檔案已重新命名為: ${newFilePath}`);
                    } else {
                        await log('未找到包含「匯出」的按鈕');
                        await page.screenshot({ path: `debug_${file}_export.png` });
                        await log(`已生成調試截圖: debug_${file}_export.png`);
                    }

                    await wait(3000);
                } catch (error) {
                    await log(`訪問 ${url} 時發生錯誤: ${error.message}`);
                    console.error(`訪問 ${url} 時發生錯誤:`, error);
                    if (error.name === 'TimeoutError') {
                        await log(`頁面導航或元素等待超時，跳過此 URL: ${url}`);
                    } else if (error.message.includes('net::ERR_NAME_NOT_RESOLVED')) {
                        await log(`無法解析域名，跳過此 URL: ${url}`);
                    } else if (error.message.includes('Target closed')) {
                        await log(`頁面已關閉，可能是 Chrome 進程提前終止，跳過此 URL: ${url}`);
                    } else {
                        await log(`其他錯誤，詳情: ${error.stack}`);
                    }
                } 
            }
        }

        await browser.disconnect();
        await log('所有網頁處理完成！');
    } catch (error) {
        await log(`程式執行發生錯誤: ${error.message}`);
        console.error('程式執行發生錯誤:', error);
    } finally {
        if (chromeProcess) {
            chromeProcess.kill();
            await log('Chrome 進程已清理');
        }
    }
}

crawlMeetingUrls().catch(console.error);