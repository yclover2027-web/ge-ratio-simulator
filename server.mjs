import http from 'node:http';
import https from 'node:https';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const currentDir = path.dirname(fileURLToPath(import.meta.url));
const port = Number(process.env.PORT || 8000);
const mhlwStartPageUrl = 'https://www.mhlw.go.jp/topics/2020/04/tp20200401-01.html';

const contentTypes = {
    '.html': 'text/html; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.csv': 'text/csv; charset=utf-8',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

const server = http.createServer(async (request, response) => {
    try {
        const requestUrl = new URL(request.url, `http://${request.headers.host}`);

        if (requestUrl.pathname === '/api/latest-mhlw-master') {
            await handleLatestMhlwMaster(response);
            return;
        }

        serveStaticFile(requestUrl.pathname, response);
    } catch (error) {
        console.error(error);
        response.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
        response.end(`サーバー内でエラーが発生しました: ${error.message}`);
    }
});

server.listen(port, '127.0.0.1', () => {
    console.log(`ジェネリックシミュレーター: http://127.0.0.1:${port}/`);
});

async function handleLatestMhlwMaster(response) {
    const latestPageUrl = await findLatestMhlwPageUrl(mhlwStartPageUrl);
    const latestPageHtml = await downloadText(latestPageUrl);
    const latestExcelUrl = findOtherMasterExcelUrl(latestPageHtml, latestPageUrl);
    const excelBuffer = await downloadBuffer(latestExcelUrl);
    const fileName = path.basename(new URL(latestExcelUrl).pathname);

    response.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Cache-Control': 'no-store',
        'X-Source-Url': latestExcelUrl,
        'X-File-Name': decodeURIComponent(fileName)
    });
    response.end(excelBuffer);
}

async function findLatestMhlwPageUrl(startUrl) {
    let currentUrl = startUrl;

    // 厚労省ページは古い年度から「最新の情報はこちら」で新しいページへつながっています。
    // リンクが見つからなくなるまで進めることで、今後の年度更新にも追随しやすくします。
    for (let i = 0; i < 12; i++) {
        const html = await downloadText(currentUrl);
        const nextUrl = findNextLatestPageUrl(html, currentUrl);
        if (!nextUrl) return currentUrl;
        currentUrl = nextUrl;
    }

    throw new Error('最新版ページの探索回数が上限を超えました。');
}

function findNextLatestPageUrl(html, baseUrl) {
    const anchorPattern = /<a\b[^>]*href=["']([^"']+)["'][^>]*>\s*こちら\s*<\/a>/gi;
    let match;

    while ((match = anchorPattern.exec(html)) !== null) {
        const beforeAnchor = html.slice(Math.max(0, match.index - 180), match.index);
        if (beforeAnchor.includes('最新の情報')) {
            return new URL(match[1], baseUrl).href;
        }
    }

    const baseYear = Number(new URL(baseUrl).pathname.match(/\/topics\/(\d{4})\//)?.[1] || 0);
    const pageLinks = Array.from(html.matchAll(/<a\b[^>]*href=["']([^"']+)["'][^>]*>/gi))
        .map(linkMatch => new URL(linkMatch[1], baseUrl).href)
        .map(href => {
            const year = Number(new URL(href).pathname.match(/\/topics\/(\d{4})\/04\/tp\d{8}-01\.html$/)?.[1] || 0);
            return { href, year };
        })
        .filter(link => link.year > baseYear)
        .sort((a, b) => a.year - b.year);

    return pageLinks[0]?.href || '';
}

function findOtherMasterExcelUrl(html, pageUrl) {
    const excelLinks = Array.from(html.matchAll(/<a\b[^>]*href=["']([^"']+\.(?:xlsx?|xls))["'][^>]*>/gi))
        .map(match => match[1]);

    // 厚労省の薬価基準ページでは、5番「その他」のExcelファイル名が *_05.xlsx 形式です。
    // 見出しの表記ゆれに左右されないよう、まずファイル名で5番のExcelを選びます。
    const otherExcelLink = excelLinks.find(href => /_05\.(?:xlsx?|xls)$/i.test(href));

    if (!otherExcelLink) {
        throw new Error('「5．その他」のExcelリンクが見つかりませんでした。');
    }

    return new URL(otherExcelLink, pageUrl).href;
}

function serveStaticFile(pathname, response) {
    const requestedPath = pathname === '/' ? '/index.html' : pathname;
    const filePath = path.resolve(currentDir, `.${decodeURIComponent(requestedPath)}`);

    // URLに ../ が混ざっても、プロジェクトフォルダ外のファイルを返さないようにします。
    if (!filePath.startsWith(currentDir)) {
        response.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
        response.end('Forbidden');
        return;
    }

    fs.readFile(filePath, (error, data) => {
        if (error) {
            response.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
            response.end('Not found');
            return;
        }

        response.writeHead(200, {
            'Content-Type': contentTypes[path.extname(filePath)] || 'application/octet-stream',
            'Cache-Control': 'no-store'
        });
        response.end(data);
    });
}

function downloadText(url) {
    return downloadBuffer(url).then(buffer => buffer.toString('utf8'));
}

function downloadBuffer(url) {
    return new Promise((resolve, reject) => {
        https.get(url, response => {
            if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
                response.resume();
                downloadBuffer(new URL(response.headers.location, url).href).then(resolve, reject);
                return;
            }

            if (response.statusCode !== 200) {
                response.resume();
                reject(new Error(`取得に失敗しました: ${response.statusCode} ${url}`));
                return;
            }

            const chunks = [];
            response.on('data', chunk => chunks.push(chunk));
            response.on('end', () => resolve(Buffer.concat(chunks)));
        }).on('error', reject);
    });
}
