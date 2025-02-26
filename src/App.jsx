import React, { useState } from "react";
import { Upload, Button, message } from "antd";
import { InboxOutlined } from "@ant-design/icons";
import { parseJSON, downloadExcel } from "./utils"; // 引入 utils.js
import "./App.css"

const { Dragger } = Upload;

const ExcelConverter = () => {
    const [jsonFiles, setJsonFiles] = useState([]);

    const handleUpload = (file) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const result = e.target.result;
                if (!result) throw new Error("檔案內容為空");

                const parsedMessages = parseJSON(result);
                setJsonFiles((prev) => [...prev, { fileName: file.name, messages: parsedMessages.messages, sender: parsedMessages.sender }]);
                message.success(`${file.name} 解析成功！`);
            } catch (error) {
                message.error(`${file.name} 解析失敗：${error.message}`);
            }
        };

        reader.onerror = () => {
            message.error("讀取檔案失敗");
        };

        reader.readAsText(file);
        return false; // 阻止自動上傳
    };
    const onChange = (info) => {
        const { status } = info.file;
        if (status !== 'uploading') {
            console.log(info.file, info.fileList);
        }
        if (status === 'done') {
            message.success(`${info.file.name} file uploaded successfully.`);
        } else if (status === 'error') {
            message.error(`${info.file.name} file upload failed.`);
        }
    }

    return (
        <div className={'upload-section'} style={{ padding: "20px" }}>
            <Dragger beforeUpload={handleUpload} showUploadList accept=".json" multiple onChange={onChange}>
                <p className="ant-upload-drag-icon">
                    <InboxOutlined />
                </p>
                <p className="ant-upload-text">拖曳或點擊上傳 JSON</p>
                <p className="ant-upload-hint">支援 .json 檔案</p>
            </Dragger>

            <Button type="primary" onClick={() => downloadExcel(jsonFiles)} style={{ marginTop: 20 }} disabled={jsonFiles.length === 0}>
                下載 Excel
            </Button>
        </div>
    );
};

export default ExcelConverter;
