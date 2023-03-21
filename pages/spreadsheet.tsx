/* eslint-disable @next/next/no-img-element */
import { IgrSpreadsheet, IgrSpreadsheetEditModeEnteringEventArgs } from "igniteui-react-spreadsheet";
import React from "react";
import { IgnSpreadsheet, loadFromUrl, saveWorkbook, fillOut } from "../hooks/useIgnSpreadsheet";

const Spreadsheet = () => {
  const [spreadsheet, setSpreadsheet] = React.useState<IgrSpreadsheet>();
  const [isLoading, setIsLoading] = React.useState(true);

  const onEditModeEntering = (
    s: IgrSpreadsheet,
    e: IgrSpreadsheetEditModeEnteringEventArgs
  ) => {
    e.cancel = true;
  };

  // 初期読み込みと同時にstateにも設定する
  const spreadsheetRef = React.useCallback((ss:IgrSpreadsheet) => {
    if (ss != null && loadFromUrl != null) {
      setSpreadsheet(ss);
      const url = "/product_standard_template_with_start_and_end_image_cell.xlsx"; // 画像埋め込みセルを開始・終了セルで指定
      // const url = "/product_standard_template_with_merged_image_cell.xlsx"; // 画像埋め込みセルをマージされた単一セルで指定
      loadFromUrl(url).then((w) => {
        /* eslint-disable no-param-reassign */
        (ss as IgrSpreadsheet).workbook = w;
        (ss as IgrSpreadsheet).workbook.protect(false,false);
        (ss as IgrSpreadsheet).areHeadersVisible = false;
        (ss as IgrSpreadsheet).isFormulaBarVisible = false;
        (ss as IgrSpreadsheet).areGridlinesVisible = false;
        (ss as IgrSpreadsheet).editModeEntering = onEditModeEntering;
        (ss as IgrSpreadsheet).zoomLevel = 100; // set number between 10 and 400, 100 is default.
        /* eslint-enable no-param-reassign */
        if (fillOut) {
          fillOut((ss as IgrSpreadsheet).workbook);
        }
        setIsLoading(false);
      });
    }
  }, []);

  const handleSave = () => {
    if (spreadsheet && saveWorkbook) {
      saveWorkbook(spreadsheet.workbook, "作成したファイル");
    }
  };

  return (
    <div style={{ padding: "1em" }}>
          <div style={{ marginBottom: "1em" }}>
            <button type="button" onClick={handleSave}>
              保存
            </button>
            <div style={{ overflow: "hidden", width:0, height:0 }}>
              <img style={{visibility:"hidden"}} id="imgdesu" src="https://erp-7bejsjlum-canbright.vercel.app/_next/image?url=https%3A%2F%2Fstg.api-canbright.jp%2Frails%2Factive_storage%2Frepresentations%2Fredirect%2FeyJfcmFpbHMiOnsibWVzc2FnZSI6IkJBaHBBaTBQIiwiZXhwIjpudWxsLCJwdXIiOiJibG9iX2lkIn19--447212be2e08cfda0d6638045c62f66af206472e%2FeyJfcmFpbHMiOnsibWVzc2FnZSI6IkJBaDdCem9MWm05eWJXRjBTU0lKYW5CbFp3WTZCa1ZVT2hSeVpYTnBlbVZmZEc5ZmJHbHRhWFJiQjJrQ1BBRnBBaFFCIiwiZXhwIjpudWxsLCJwdXIiOiJ2YXJpYXRpb24ifX0%3D--c182fed9d502785aebdb98e37d8483134204658d%2Fg8XFG.jpeg&amp;w=128&amp;q=75" alt="" />
            </div>
          </div>
        <div className="spread-container" style={{ height: "calc(100vh - 5em)", width: "100%", border: "1px solid #999", padding: "1em", boxSizing: "border-box" }}>
        {isLoading
          ?
          <div className="loader-wrap">
            <div className="loader">Loading...</div>
          </div>
          :
          null
        }
            <IgnSpreadsheet
              spreadsheetRef={spreadsheetRef}
              height="100%"
            />
        </div>
    </div>
  );
}

export default Spreadsheet;
