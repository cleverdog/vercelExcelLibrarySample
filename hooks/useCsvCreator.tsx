import { saveAs } from "file-saver";
import dynamic from "next/dynamic";
import { IgrSpreadsheet, IIgrSpreadsheetProps } from "igniteui-react-spreadsheet";
import React, { LegacyRef } from "react";

/* eslint-disable  import/no-mutable-exports */
export let saveCsv:
  | ((header: string[], rows: string[][]) => Promise<null>)
  | null = null;

interface IIgnSpreadsheetProps extends IIgrSpreadsheetProps {
  spreadsheetRef: LegacyRef<IgrSpreadsheet>;
}

export const IgnSpreadsheet = dynamic(
  async () => {
    // eslint-disable-next-line @typescript-eslint/no-shadow
    const { IgrSpreadsheetModule, IgrSpreadsheet } = await import(
      "igniteui-react-spreadsheet"
    );
    const { Workbook, WorkbookSaveOptions, WorkbookFormat } = await import(
      "igniteui-react-excel"
    );

    IgrSpreadsheetModule.register();

    saveCsv = async (
      headers: string[],
      rows: string[][]
    ): Promise<null> => {
      return new Promise<null>((resolve, reject) => {
        const workbook = new Workbook(WorkbookFormat.Excel2007);
        const sheet = workbook.worksheets().add("Sheet1");
        headers.forEach((header, i) => {
          sheet.rows(0).cells(i).value = header;
        });

        rows.forEach((row, i) => {
          row.forEach((cell, j) => {
            sheet.rows(1 + i).cells(j).value = cell;
          });
        });

        const opt = new WorkbookSaveOptions();
        opt.type = "blob";
        workbook.save(
          opt,
          (d) => {
            const fileExt = ".xlsx";
            const fileName = `作成したファイル(CSV)${fileExt}`;
            saveAs(d as Blob, fileName);
            resolve(null);
          },
          (e) => {
            reject(e);
          }
        );
      });
    };

    // refを設定するためにIgnSpreadsheetというコンポーネントを作成
    const IgnSpreadsheetComponent = ({
      spreadsheetRef,
      ...props
    }: IIgnSpreadsheetProps) => {
      return <IgrSpreadsheet ref={spreadsheetRef} {...props} />;
    };

    return IgnSpreadsheetComponent;
  },
  { ssr: false }
);
