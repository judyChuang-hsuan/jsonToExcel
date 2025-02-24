import fs from 'fs';
import { Parser } from '@json2csv/plainjs';
import XLSX from 'xlsx-js-style';
import {format} from 'date-fns';

function decode(text) {
    if (!text) return "";

    let charArr = [];
    for (let i = 0; i < text.length; i++) {
        charArr.push(text.charCodeAt(i));
    }

    return new TextDecoder().decode(new Uint8Array(charArr));
}

const headerStyle = {
    font: { bold: true, color: { rgb: "FFFFFF" } },
    fill: { fgColor: { rgb: "4F81BD" } }, // 背景顏色
    alignment: { horizontal: "center", vertical: "center" },
    border: {
        top: { style: "thin", color: { rgb: "000000" } },
        bottom: { style: "thin", color: { rgb: "000000" } },
        left: { style: "thin", color: { rgb: "000000" } },
        right: { style: "thin", color: { rgb: "000000" } }
    }
};

const cellStyle = {
    alignment: { horizontal: "left", vertical: "center" },
    border: {
        top: { style: "thin", color: { rgb: "000000" } },
        bottom: { style: "thin", color: { rgb: "000000" } },
        left: { style: "thin", color: { rgb: "000000" } },
        right: { style: "thin", color: { rgb: "000000" } }
    }
};

// 讀取 JSON 檔案
fs.readFile('../assets/message1.json', 'utf8', (err, data) => {
    if (err) {
        console.error('讀取 JSON 失敗:', err);
        return;
    }

    try {
        const jsonData = JSON.parse(data);

        // 取出 messages 陣列並解碼 Unicode 字符
        const messages = jsonData.messages.map(msg => ({
            sender_name: decode(msg.sender_name),
            timestamp: format(new Date(msg.timestamp_ms), 'yyyy-MM-dd HH:mm:ss'), // 轉換時間格式
            content: msg.content ? decode(msg.content) : '', // 解碼內容
            photos: msg.photos ? msg.photos.map(photo => decode(photo.uri)).join('; ') : '',
        }));

        // 檢查是否有資料
        if (messages.length === 0) {
            throw new Error("No messages found in the data");
        }

        // 設定 CSV 欄位
        const opts = { fields: ['sender_name', 'timestamp', 'content', 'photos'] };
        const parser = new Parser(opts);
        const csv = parser.parse(messages);

        // 將 CSV 存檔
        fs.writeFile('messages.csv', csv, (err) => {
            if (err) {
                console.error('寫入 CSV 失敗:', err);
            } else {
                console.log('成功轉換 JSON 至 CSV，檔案已儲存為 messages.csv');
                console.log(XLSX,'xlsx')
                // 轉換 JSON 資料為 Excel
                const wb = XLSX.utils.book_new(); // 創建新的工作簿

                const ws = XLSX.utils.json_to_sheet(messages, { header: ['sender_name', 'timestamp', 'content', 'photos'] }); // 轉換 JSON 資料為工作表

                // 設定標題行樣式
                const range = XLSX.utils.decode_range(ws['!ref']);  // 獲取工作表範圍
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const cell = ws[XLSX.utils.encode_cell({ r: 0, c: col })]; // 獲取標題行
                    if (cell) {
                        cell.s = headerStyle; // 應用樣式
                    }
                }

                // 設定資料行樣式
                for (let row = range.s.r + 1; row <= range.e.r; row++) {
                    for (let col = range.s.c; col <= range.e.c; col++) {
                        const cell = ws[XLSX.utils.encode_cell({ r: row, c: col })]; // 獲取每個資料行的單元格
                        if (cell) {
                            cell.s = cellStyle; // 應用樣式
                        }
                    }
                }

                // 設定列寬
                ws['!cols'] = [
                    { width: 20 },  // sender_name
                    { width: 30 },  // timestamp
                    { width: 100 },  // content
                    { width: 30 },  // photos
                ];

                // 設定行高
                ws['!rows'] = [
                    { hpt: 20 },  // 設定標題行高度
                    ...new Array(messages.length).fill({ hpt: 40 }) // 設定資料行高度
                ];

                // 將工作表添加到工作簿
                XLSX.utils.book_append_sheet(wb, ws, 'Messages');

                // 將 Excel 檔案寫入磁碟
                XLSX.writeFile(wb, 'messages_with_style.xlsx');
                console.log('成功將 CSV 轉換為帶有樣式的 Excel，檔案已儲存為 messages_with_style.xlsx');
            }
        });

    } catch (err) {
        console.error('解析 JSON 失敗:', err);
    }
});
