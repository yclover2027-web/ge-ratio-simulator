const https = require('node:https');
const path = require('node:path');

const mhlwStartPageUrl = 'https://www.mhlw.go.jp/topics/2020/04/tp20200401-01.html';

module.exports = async function handler(request, response) {
    if (request.method !== 'GET' && request.method !== 'HEAD') {
        response.setHeader('Allow', 'GET, HEAD');
        response.status(405).send('Method Not Allowed');
        return;
    }

    try {
        const latestPageUrl = await findLatestMhlwPageUrl(mhlwStartPageUrl);
        const latestPageHtml = await downloadText(latestPageUrl);
        const latestExcelUrl = findOtherMasterExcelUrl(latestPageHtml, latestPageUrl);
        const excelBuffer = await downloadBuffer(latestExcelUrl);
        const fileName = path.basename(new URL(latestExcelUrl).pathname);

        response.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        response.setHeader('Cache-Control', 'no-store');
        response.setHeader('X-Source-Url', latestExcelUrl);
        response.setHeader('X-File-Name', decodeURIComponent(fileName));
        response.status(200).send(excelBuffer);
    } catch (error) {
        console.error(error);
        response.status(500).send(`最新版Excelの取得に失敗しました: ${error.message}`);
    }
};

async function findLatestMhlwPageUrl(startUrl) {
    let currentUrl = startUrl;

    // 厚労省ページは古い年度から新しい年度のページへリンクされています。
    // 毎年URLが変わるため、次年度リンクがなくなるまで進めて最新版ページを見つけます。
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

    // 薬価基準ページでは、5番「その他」のExcelファイル名が *_05.xlsx 形式です。
    // 見出しの表記ゆれに左右されないよう、ファイル名で5番のExcelを選びます。
    const otherExcelLink = excelLinks.find(href => /_05\.(?:xlsx?|xls)$/i.test(href));

    if (!otherExcelLink) {
        throw new Error('「5．その他」のExcelリンクが見つかりませんでした。');
    }

    return new URL(otherExcelLink, pageUrl).href;
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
