import React, {use, useState} from "react";
import {Upload, Button, message, Spin} from "antd";
import {InboxOutlined} from "@ant-design/icons";
import {parseJSON, downloadExcel} from "./utils"; // 引入 utils.js
import "./App.css"

const {Dragger} = Upload;

const ExcelConverter = () => {
    const [jsonFiles, setJsonFiles] = useState([]);
    const [loading, setLoading] = useState(false);
    const [fileList, setFileList] = useState([]);

    const handleUpload = () => {
        return false
    }

    const onChange = (info) => {
        setLoading(true)

        if (!info.fileList.length) {
            message.error("未偵測到 JSON 檔案！");
            return;
        }

        setFileList(info.fileList);

        const newJsonFiles = [];

        info.fileList.forEach((fileWrapper, index) => {
            const file = fileWrapper.originFileObj || fileWrapper; // 確保是 File 類型

            if (!(file instanceof Blob)) {
                console.error("檔案格式錯誤，無法讀取", file);
                return;
            }

            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const result = e.target.result;

                    const parsedMessages = parseJSON(result);
                    newJsonFiles.push({
                        fileName: file.name,
                        messages: parsedMessages.messages,
                        sender: parsedMessages.sender,
                    });

                    if (newJsonFiles.length === info.fileList.length) {
                        console.log("所有 JSON 解析完成，更新 jsonFiles", newJsonFiles);
                        setJsonFiles(newJsonFiles);
                        message.success("所有檔案解析成功！");
                        setLoading(false)
                    }
                } catch (error) {
                    message.error(`${file.name} 解析失敗：${error.message}`);
                    setLoading(false)
                }
            };

            reader.onerror = () => {
                message.error(`讀取 ${file.name} 失敗`);
                setLoading(false)
            };

            reader.readAsText(file);
        });
    };

    const onRemoveFiles = () => {
        setJsonFiles([])
        setFileList([])
    }


    return (
        <div className={'upload-section'} style={{padding: "20px"}}>
            <Dragger fileList={fileList}
                     beforeUpload={handleUpload}
                     showUploadList
                     accept=".json"
                     multiple
                     directory
                     onChange={onChange}>
                <p className="ant-upload-drag-icon">
                    <InboxOutlined/>
                </p>
                <p className="ant-upload-text">拖曳或點擊上傳 JSON</p>
                <p className="ant-upload-hint">支援 .json 檔案</p>
            </Dragger>

            <Button type="primary" onClick={() => downloadExcel(jsonFiles)} style={{marginTop: 20}}
                    disabled={jsonFiles.length === 0 && loading}>
                下載 Excel
            </Button>
            <Button type="default" onClick={onRemoveFiles} style={{marginTop: 20, marginLeft: 10}}>
                清空上傳資料
            </Button>
            <Spin spinning={loading} fullscreen></Spin>
        </div>
    );
};

export default ExcelConverter;
