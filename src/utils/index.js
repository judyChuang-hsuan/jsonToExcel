import * as XLSX from "xlsx-js-style";
import {format} from "date-fns";

const MAX_SHEET_NAME_LENGTH = 20
// 標題樣式
const headerStyle = {
    font: {bold: true, color: {rgb: "FFFFFF"}},
    fill: {fgColor: {rgb: "4F81BD"}},
    alignment: {horizontal: "center", vertical: "center"},
    border: {top: {style: "thin"}, bottom: {style: "thin"}, left: {style: "thin"}, right: {style: "thin"}},
};

// 單元格樣式
const cellStyle = {
    alignment: {horizontal: "left", vertical: "center"},
    border: {top: {style: "thin"}, bottom: {style: "thin"}, left: {style: "thin"}, right: {style: "thin"}},
};

const truncateSheetName = (name) => {
    return name.length > MAX_SHEET_NAME_LENGTH
        ? name.substring(0, MAX_SHEET_NAME_LENGTH)
        : name;
};

// 生成唯一的 Sheet 名稱
const getUniqueSheetName = (wb, baseName) => {
    let sheetName = truncateSheetName(baseName);  // 先截斷名稱長度
    let counter = 1;

    // 檢查是否已經有同名的 Sheet
    while (wb.SheetNames.includes(sheetName)) {
        const truncatedBaseName = truncateSheetName(baseName); // 重新截斷基礎名稱
        // 控制數字編號，避免超過最大字數限制
        const counterStr = `_${counter}`;
        sheetName = `${truncatedBaseName.substring(0, 31 - counterStr.length)}${counterStr}`;
        counter++;
    }

    return sheetName;
};


const decode = (text) => {
    if (!text) return "";

    let charArr = [];
    for (let i = 0; i < text.length; i++) {
        charArr.push(text.charCodeAt(i));
    }

    return new TextDecoder().decode(new Uint8Array(charArr));
}

// 解析 JSON 並轉換為格式化資料
export const parseJSON = (jsonString) => {
    if (!jsonString) throw new Error("JSON 檔案內容為空");
    let jsonData;
    try {
        jsonData = JSON.parse(jsonString);
        if (!jsonData.messages) throw new Error("無 messages 陣列");
        const messages = jsonData.messages.map((msg) => ({
            sender_name: decode(msg.sender_name),
            timestamp: format(new Date(msg.timestamp_ms), "yyyy-MM-dd HH:mm:ss"),
            content: decode(msg.content) || "",
        }));
        return {
            sender: decode(jsonData.participants[0].name),
            messages
        }
    } catch (error) {
        throw new Error("JSON 解析失敗：" + error.message);
    }
};

// 轉換資料並下載 Excel
export const downloadExcel = (jsonFiles) => {
    console.log('download')

    if (jsonFiles.length === 0) return;

    const wb = XLSX.utils.book_new();
    jsonFiles.forEach(({fileName, messages, sender}) => {
        const ws = XLSX.utils.json_to_sheet(messages, {header: ["sender_name", "timestamp", "content"]});
        // 設定標題樣式
        const range = XLSX.utils.decode_range(ws["!ref"]);
        // 自適應欄寬
        const colWidths = [{width: 10}, {width: 10}, {width: 10}]; // 初始欄寬
        for (let row = range.s.r; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cell = ws[XLSX.utils.encode_cell({r: row, c: col})];
                if (cell && cell.v) {
                    const contentLength = cell.v.toString().length;
                    if (contentLength > colWidths[col].width) {
                        colWidths[col].width = contentLength + 5; // 增加寬度緩衝
                    }
                }
            }
        }
        ws["!cols"] = colWidths;

        // 設定內容樣式
        // 自適應行高
        const rowHeights = [{hpt: 20}]; // 標題行高度
        for (let row = range.s.r + 1; row <= range.e.r; row++) {
            let maxLines = 1;
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cell = ws[XLSX.utils.encode_cell({r: row, c: col})];
                if (cell && cell.v) {
                    const lines = cell.v.toString().split("\n").length; // 計算換行數
                    if (lines > maxLines) maxLines = lines;
                }
            }
            rowHeights.push({hpt: maxLines * 20}); // 每行約 20px 高度
        }
        ws["!rows"] = rowHeights;

        // 設定標題樣式
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cell = ws[XLSX.utils.encode_cell({r: 0, c: col})];
            if (cell) cell.s = headerStyle;
        }

        // 設定內容樣式
        for (let row = range.s.r + 1; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cell = ws[XLSX.utils.encode_cell({r: row, c: col})];
                if (cell) cell.s = cellStyle;
            }
        }

        const sheetName = getUniqueSheetName(wb, sender)

        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    })


    const excelBuffer = XLSX.write(wb, {bookType: "xlsx", type: "array"});
    const blob = new Blob([excelBuffer], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "instagram_messages.xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
};
