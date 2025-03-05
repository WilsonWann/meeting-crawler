const path = require('path');
const fs = require('fs').promises;
const dotenv = require('dotenv');
const puppeteer = require('puppeteer-core');
const { execFile } = require('child_process');
const ini = require('ini');

// 動態確定 .env 路徑
const envPath = path.join('./', '.env');
dotenv.config({ path: envPath });

// 環境變數與預設值
const desktopPath = process.env.DESKTOP_PATH || path.join(process.env.USERPROFILE, 'Desktop', 'Meeting');
const downloadPath = process.env.DOWNLOAD_PATH || path.join(process.env.USERPROFILE, 'Desktop', 'Downloads');
const logFilePath = process.env.LOG_FILE_PATH || path.join(path.dirname(process.execPath), 'crawler_log.txt');
const remoteDebuggingPort = process.env.REMOTE_DEBUGGING_PORT || '9222';
const chromePath = process.env.CHROME_PATH || 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const downloadType = process.env.DOWNLOAD_TYPE || 'WORD';
const downloadTimeout = process.env.DOWNLOAD_TIMEOUT || 15000;

const wait = ms => new Promise(resolve => setTimeout(resolve, ms));

async function log(message) {
    const timestamp = new Date().toISOString();
    const logEntry = `[${timestamp}] ${message}\n`;
    await fs.appendFile(logFilePath, logEntry, 'utf-8');
}

async function startChrome() {
    return new Promise((resolve, reject) => {
        const chromeProcess = execFile(
            chromePath,
            [
                `--remote-debugging-port=${remoteDebuggingPort}`,
                '--no-first-run',
                '--no-default-browser-check'
            ], // 移除 --user-data-dir
            (error) => {
                if (error) reject(error);
            }
        );
        setTimeout(() => resolve(chromeProcess), 3000);
    });
}

async function waitForDownload(downloadPath, startTime, timeout = 15000) {
    const initialFiles = await fs.readdir(downloadPath);
    let downloadedFile;

    while (Date.now() - startTime < timeout) {
        const currentFiles = await fs.readdir(downloadPath);
        const newFiles = currentFiles.filter(file => !initialFiles.includes(file));

        if (newFiles.length > 0) {
            const tempFiles = newFiles.filter(file => file.endsWith('.crdownload'));
            if (tempFiles.length > 0) {
                await log(`檢測到臨時檔案: ${tempFiles}`);
                downloadedFile = tempFiles[0].replace('.crdownload', '');
            } else {
                downloadedFile = newFiles[0];
            }

            const filePath = path.join(downloadPath, downloadedFile);
            try {
                const stats = await fs.stat(filePath);
                if (!downloadedFile.endsWith('.crdownload')) {
                    await log(`檢測到新檔案: ${filePath}`);
                    const initialStats = await fs.stat(filePath);
                    await wait(100);
                    const finalStats = await fs.stat(filePath);
                    if (initialStats.size === finalStats.size) {
                        return filePath;
                    }
                }
            } catch (error) {
                await log(`檢查檔案 ${filePath} 時發生錯誤: ${error.message}`);
            }
        }
        await wait(100);
    }
    throw new Error('下載超時');
}

async function crawlMeetingUrls() {
    let chromeProcess;
    try {
        await log('程式啟動');
        await log(`使用 .env 路徑: ${envPath}`);
        chromeProcess = await startChrome();
        await fs.mkdir(downloadPath, { recursive: true });

        const browser = await puppeteer.connect({
            browserURL: `http://localhost:${remoteDebuggingPort}`,
            defaultViewport: null,
            executablePath: chromePath
        });
        await log('成功連接到 Chrome 遠端除錯實例');

        const files = await fs.readdir(desktopPath);
        const urlFiles = files.filter(file => path.extname(file).toLowerCase() === '.url');

        for (const file of urlFiles) {
            const filePath = path.join(desktopPath, file);
            const url = await readUrlFile(filePath);

            if (url) {
                await log(`開始處理: ${file} - ${url}`);
                const page = await browser.newPage();

                await page._client().send('Page.setDownloadBehavior', {
                    behavior: 'allow',
                    downloadPath: downloadPath
                });

                try {
                    await page.setViewport({ width: 1920, height: 1080 });
                    await page.goto(url, { waitUntil: 'networkidle0', timeout: 30000 });
                    await log(`成功載入頁面: ${url}`);

                    const moreMenuButton = await page.$('div.suite-more-menu > button');
                    if (moreMenuButton) {
                        await moreMenuButton.click();
                        await log('已點擊 div.suite-more-menu > button');
                        await wait(500);
                    } else {
                        await log('未找到 div.suite-more-menu > button');
                    }

                    await page.mouse.move(1572, 436);
                    await log('已移動到座標 (1572, 436)');
                    await wait(500);

                    const downloadCoords = downloadType.toUpperCase() === 'PDF' ? 
                        { x: 1772, y: 456 } : 
                        { x: 1772, y: 436 };
                    await page.mouse.move(downloadCoords.x, downloadCoords.y);
                    await log(`已移動到座標 (${downloadCoords.x}, ${downloadCoords.y}) - ${downloadType}`);
                    await page.mouse.click(downloadCoords.x, downloadCoords.y);
                    await log(`已在座標 (${downloadCoords.x}, ${downloadCoords.y}) 執行點擊`);
                    await wait(500);

                    const exportButton = await page.evaluateHandle(() => {
                        const buttons = Array.from(document.querySelectorAll('button, a, [role="button"]'));
                        return buttons.find(btn => btn.textContent.trim().includes('匯出'));
                    });

                    if (exportButton) {
                        const downloadStartTime = Date.now();
                        await exportButton.click();
                        await log('已點擊「匯出」按鈕');

                        const downloadedFile = await waitForDownload(downloadPath, downloadStartTime, downloadTimeout);
                        await log(`檔案已下載至: ${downloadedFile}`);

                        const fileBaseName = file.replace('.url', '');
                        const newFileName = `${fileBaseName}_${Date.now()}${path.extname(downloadedFile)}`;
                        const newFilePath = path.join(downloadPath, newFileName);
                        await fs.rename(downloadedFile, newFilePath);
                        await log(`檔案已重新命名為: ${newFilePath}`);
                    } else {
                        await log('未找到「匯出」按鈕');
                    }

                    await page.close();
                    await log(`頁面已關閉: ${url}`);
                } catch (error) {
                    await log(`訪問 ${url} 時發生錯誤: ${error.message}`);
                    console.error(`訪問 ${url} 時發生錯誤:`, error);
                    await page.close();
                }
            } else {
                await log(`無法解析 URL 檔案: ${file}`);
            }
        }

        await browser.disconnect();
        await log('所有網頁處理完成！');
        console.log('所有網頁處理完成！');
    } catch (error) {
        await log(`程式執行發生錯誤: ${error.message}`);
        console.error('程式執行發生錯誤:', error);
    } finally {
        if (chromeProcess) chromeProcess.kill();
        await log('Chrome 進程已清理');
    }
}

async function readUrlFile(filePath) {
    try {
        const content = await fs.readFile(filePath, 'utf-8');
        const parsed = ini.parse(content);
        return parsed.InternetShortcut.URL;
    } catch (error) {
        console.error(`讀取 ${filePath} 時發生錯誤:`, error);
        return null;
    }
}

(async () => {
    await crawlMeetingUrls();
})();