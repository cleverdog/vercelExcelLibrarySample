import { IgrSpreadsheet } from "igniteui-react-spreadsheet";
import { useRef, useState } from "react";
import { MultiSelect } from "react-multi-select-component";
import { IgnSpreadsheet, saveCsv } from "../hooks/useCsvCreator";
import sampleData from "../sampleData.json";

const CsvCreator = () => {
  const spreadsheetRef = useRef<IgrSpreadsheet>(null);
  const [selected, setSelected] = useState<{ label: string; value: string }[]>(
    []
  );

  const handleSave = () => {
    const headers = selected.map((s) => s.label);
    const keys = selected.map((s) => s.value);
    const rows = (sampleData.map((e) => keys.map((key) => e[key as keyof Params])) as string[][]);
    if (saveCsv) {
      saveCsv(headers, rows);
    }
  };

  interface Params {
    Id: string;
    CompanyName: string;
    ContactName: string;
    ContactTitle: string;
    Address: string;
    City: string;
    Region: null;
    PostalCode: number;
    Country: string;
    Phone: string;
    Fax: string;
  }

  const options = [
    { label: "会社名", value: "CompanyName" },
    { label: "住所", value: "Address" },
    { label: "市区町村", value: "City" },
    { label: "郵便番号", value: "PostalCode" },
    { label: "国", value: "Country" },
    { label: "電話番号", value: "Phone" },
    { label: "FAX", value: "Fax" },
  ];
  return (
    <>
      <div style={{ display: "none" }}>
        <IgnSpreadsheet spreadsheetRef={spreadsheetRef} />
      </div>
      <MultiSelect
        options={options}
        value={selected}
        onChange={setSelected}
        labelledBy="Select"
      />
      <div style={{ display: "flex" }}>
        <button type="button" onClick={handleSave}>
          保存
        </button>
      </div>
    </>
  );
};

export default CsvCreator;
