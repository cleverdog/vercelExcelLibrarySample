import { IgrSpreadsheet } from "igniteui-react-spreadsheet";
import React, { useRef } from "react";
import {
  IgnSpreadsheet,
  saveWorkbook,
  loadFromUpload,
} from "../hooks/useIgnSpreadsheet";

const Spreadsheet = () => {
  const spreadsheetRef = useRef<IgrSpreadsheet>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const onFileInputChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files && loadFromUpload) {
      loadFromUpload(event.target.files[0]).then((w) => {
        if (spreadsheetRef.current != null) {
          spreadsheetRef.current.workbook = w;
        }
      });
    }
  };

  const fileUpload = () => {
    inputRef.current?.click();
  };

  const handleSave = () => {
    if (saveWorkbook && spreadsheetRef.current) {
      saveWorkbook(spreadsheetRef.current.workbook, "作成したファイル");
    }
  };

  return (
    <div style={{ display: "flex" }}>
      <div style={{ flexGrow: 1 }}>
        <button type="button" style={{ margin: "10px" }} onClick={fileUpload}>
          ファイルアップロード
        </button>
        <input
          hidden
          ref={inputRef}
          type="file"
          accept=".xlsx"
          onChange={onFileInputChange}
        />
      </div>
      <div style={{ flexGrow: 3, height: "700px" }}>
        <IgnSpreadsheet
          spreadsheetRef={spreadsheetRef}
          height="100%"
          isFormulaBarVisible
        />
        <button type="button" onClick={handleSave}>
          保存
        </button>
      </div>
    </div>
  );
};

export default Spreadsheet;
