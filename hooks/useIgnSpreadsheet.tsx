import { saveAs } from "file-saver";
import { IgRect } from "igniteui-react-core";
import {
  Workbook,
  Worksheet,
  WorksheetCell,
  WorksheetCellCollection,
  WorksheetMergedCellsRegion,
  WorksheetRowCollection,
} from "igniteui-react-excel";
import {
  IgrSpreadsheet,
  IIgrSpreadsheetProps,
} from "igniteui-react-spreadsheet";
import dynamic from "next/dynamic";
import { LegacyRef } from "react";

// 置き換えサンプルデータ
const targetItem: TargetItem = {
  itemName: "リンゴ",
  itemCode: "I1234",
  testvalue1: "テスト値1",
  testvalue2: "テスト値2",
  start_at: "2022-09-01T11:16:52.000+09:00",
  itemImageMergedCell: "",
  itemImageStartCell: "",
  itemImageEndCell: "",
  allergy: {
    shrimp: false,
    crab: true,
  },
};

// 置き換え文字列
interface TargetItem extends ItemImage {
  itemName: string;
  itemCode: string;
  testvalue1: string;
  testvalue2: string;
  start_at: string;
  allergy: Allergy;
}

// 画像の配置セルは別処理が必要なため定義を分けておく
interface ItemImage {
  itemImageMergedCell?: string;
  itemImageStartCell?: string;
  itemImageEndCell?: string;
}

// 階層化したデータを取り扱う場合（このサンプルで各種アレルギーの真偽）プレースホルダーの表現は {{parent.child}} とする
interface Allergy {
  shrimp: boolean;
  crab: boolean;
}

// {{ }} で囲まれているかどうか
const isTemplate = (str: string) => {
  return str.startsWith("{{") && str.endsWith("}}");
};

// {{ }} の削除
const trimBrackets = (str: string) => {
  return str.slice(2).slice(0, -2);
};

// 画像を取得
export const getImageData = (url: string) => {
  return new Promise((resolve) => {
    const req = new XMLHttpRequest();
    req.open("GET", url);
    req.responseType = "blob";
    req.onload = () => {
      const reader = new FileReader();
      reader.onloadend = () => {
        resolve(reader.result);
      };
      reader.readAsDataURL(req.response);
    };
    req.send();
  });
}

// 画像を取得（別パターン）
export const createGetImageData = () => {
  return new Promise((resolve) => {
    // 画像エレメントを作成
    const img = document.getElementById("imgdesu") as HTMLImageElement;
    const newImg = new Image();
    newImg.src = img.src;
    newImg.crossOrigin = "anonymous";
    // Canvasを作成
    const canvas = document.createElement('canvas');
    // eslint-disable-next-line no-param-reassign
    img.crossOrigin = "anonymous";
    canvas.width = img.width;
    canvas.height = img.height;
    newImg.onload = () => {
      const context = canvas.getContext('2d');
      if (context !== null) {
        context.drawImage(newImg, 0, 0);
        const dataURL = canvas.toDataURL();
        resolve(dataURL);
      }
    }
  });
};

// 画像の縦横比を取得するためにHTML画像インスタンスを作成する
export const getActualImage = (src: string) => {
  return new Promise<HTMLImageElement>((resolve, reject) => {
    const img: HTMLImageElement = new Image();
    img.onload = () => resolve(img);
    img.onerror = (e) => reject(e);
    img.src = src;
  });
};

export const readFileAsUint8Array = (file: File): Promise<Uint8Array> => {
  return new Promise<Uint8Array>((resolve, reject) => {
    const fr = new FileReader();
    fr.onerror = () => {
      reject(fr.error);
    };

    if (fr.readAsBinaryString) {
      fr.onload = () => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const rs = (fr as any).resultString;
        const str: string = rs != null ? rs : fr.result;
        const result = new Uint8Array(str.length);
        for (let i = 0; i < str.length; i++) {
          result[i] = str.charCodeAt(i);
        }
        resolve(result);
      };
      fr.readAsBinaryString(file);
    } else {
      fr.onload = () => {
        resolve(new Uint8Array(fr.result as ArrayBuffer));
      };
      fr.readAsArrayBuffer(file);
    }
  });
}

/* eslint-disable  import/no-mutable-exports */
//  dynamic import内で関数を設定する
export let fillOut: ((workbook: Workbook) => Promise<string>) | null = null;
export let loadFromUrl: ((url: string) => Promise<Workbook>) | null = null;
export let loadFromUpload: ((file: File) => Promise<Workbook>) | null = null;

export let saveWorkbook:
  | ((workbook: Workbook, fileNameWithoutExtension: string) => Promise<string>)
  | null = null;

interface IIgnSpreadsheetProps extends IIgrSpreadsheetProps {
  spreadsheetRef: LegacyRef<IgrSpreadsheet>;
}
export const IgnSpreadsheet = dynamic(
  async () => {
    // eslint-disable-next-line @typescript-eslint/no-shadow
    const { IgrSpreadsheet, IgrSpreadsheetModule } = await import(
      "igniteui-react-spreadsheet"
    );

    const {
      // dynamic importの外と中で同じ型を使わないといけない
      // eslint-disable-next-line @typescript-eslint/no-shadow
      Workbook,
      WorkbookLoadOptions,
      WorkbookSaveOptions,
      CellReferenceMode,
      // eslint-disable-next-line @typescript-eslint/no-shadow
      WorksheetCell,
      WorksheetImage,
    } = await import("igniteui-react-excel");

    IgrSpreadsheetModule.register();

    // URLからエクセルファイルを読み込み
    loadFromUrl = (url: string): Promise<Workbook> => {
      return new Promise<Workbook>((resolve, reject) => {
        const req = new XMLHttpRequest();
        req.open("GET", url, true);
        req.responseType = "arraybuffer";
        req.onload = (): void => {
          const data = new Uint8Array(req.response);
          Workbook.load(
            data,
            new WorkbookLoadOptions(),
            (w) => {
              resolve(w);
            },
            (e) => {
              reject(e);
            }
          );
        };
        req.send();
      });
    };

    // アップロードされたエクセルファイルを読み込み
    loadFromUpload = (file: File) => {
      return new Promise<Workbook>((resolve, reject) => {
        readFileAsUint8Array(file).then(
          (a) => {
            Workbook.load(
              a,
              new WorkbookLoadOptions(),
              (w) => {
                resolve(w);
              },
              (e) => {
                reject(e);
              }
            );
          },
          (e) => {
            reject(e);
          }
        );
      });
    };

    // プレースホルダーに置き換え用のデータを埋め込む
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    fillOut = (workbook: Workbook): Promise<string> => {
      return new Promise<string>( () => {
        const itemImageCells : ItemImage = {
          itemImageMergedCell : "",
          itemImageStartCell: "",
          itemImageEndCell: ""
        }
        const wsrc: WorksheetRowCollection = (
          workbook.sheets(0) as Worksheet
        ).rows();
        // 行・列の2次元配列をまわす
        for (let i = 0; i < wsrc.count; i++) {
          const wscc: WorksheetCellCollection = wsrc.item(i).cells();
          let rowCount = 0;

          for (let j = 0; j < wscc.maxCount; j++) {
            // wscc.maxCount => 値が設定されているセルの総数
            const wsc: WorksheetCell = wscc.item(j);
            if (wsc.value !== null && isTemplate(wsc.value)) {
              const key: keyof TargetItem = trimBrackets(wsc.value) as keyof TargetItem;
              if (
                Object.keys(targetItem).includes(key) ||
                Object.keys(targetItem).includes(key.split(".")[0])
              ) {
                const address = WorksheetCell.getCellAddressString(
                  wsrc.item(i),
                  wsc.columnIndex,
                  CellReferenceMode.A1,
                  false
                );
                if (Object.keys(itemImageCells).includes(key)) {
                  // 画像配置セルの場合は一旦セルの位置情報を取得しておく
                  itemImageCells[key as keyof ItemImage] = address;
                } else if (key === "start_at") {
                  // 日付データとして埋め込みたい場合、エクセルで設定した日付用セル書式設定を反映させるためにDateオブジェクトに変換
                  /* eslint-disable no-param-reassign */
                  workbook.worksheets(0).getCell(address).value = new Date(
                    targetItem[key]
                  );
                } else if (
                  key.split(".")[0] === "allergy" &&
                  targetItem.allergy &&
                  key.split(".")[1]
                ) {
                  const childItem = (targetItem.allergy as Allergy)[key.split(".")[1] as keyof Allergy];
                  /* eslint-disable no-param-reassign */
                  workbook.worksheets(0).getCell(address).value = typeof childItem === "boolean" ? Number(childItem) : childItem;
                  // ▲ boolean型のデータの場合、エクセルで設定した真偽用セル書式設定を反映させるために数値(0 or 1)に変換
                } else {
                  // 通常のフォーマットのセルに対する処理、該当のセルに値を入れる
                  /* eslint-disable no-param-reassign */
                  workbook.worksheets(0).getCell(address).value =
                    targetItem[key];
                }
              }
              rowCount++;
              if (rowCount === wscc.count) break;
            }
          }
        }

        // 画像の埋め込み処理
        // 縦長画像テスト用 => "/vertical.png";
        // 横長画像テスト用 => "/horizontal.jpg"
        createGetImageData().then((result: string | unknown) => {
          const image = new WorksheetImage(result);

          // ▼▼▼ 画像の配置ポジション指定処理 ココカラ ▼▼▼
          image.topLeftCornerPosition = { x: 0, y: 0 };
          image.bottomRightCornerPosition = { x: 100, y: 100 };

          // 画像の埋め込み先がマージされたセル {{itemImageMergedCell}} の場合
          if (itemImageCells.itemImageMergedCell) {
            let imageCell: WorksheetMergedCellsRegion | unknown;
            const mergedCells = workbook.worksheets(0).mergedCellsRegions(); // マージされたセルにアプローチするためにmergedCellsRegions()を利用
            for (let k = 0; k < mergedCells.count; k++) {
              const mergedCellRegion = mergedCells.item(k);
              if (mergedCellRegion.value === "{{itemImageMergedCell}}") {
                imageCell = mergedCellRegion; // mergedCellRegionにはfirstRow,firstColumn,lastRow,lastColumnのインデックス番号が含まれる
              }
            }
            image.topLeftCornerCell = workbook
              .worksheets(0)
              .getCell(
                `R${(imageCell as WorksheetMergedCellsRegion).firstRow + 1}C${
                  (imageCell as WorksheetMergedCellsRegion).firstColumn + 1
                }`,
                CellReferenceMode.R1C1
              );
            image.bottomRightCornerCell = workbook
              .worksheets(0)
              .getCell(
                `R${(imageCell as WorksheetMergedCellsRegion).lastRow + 1}C${
                  (imageCell as WorksheetMergedCellsRegion).lastColumn + 1
                }`,
                CellReferenceMode.R1C1
              );
          }
          // 開始・終了セル {{itemImageStartCell}} {{itemImageEndCell}} がそれぞれ指定されたセル範囲に画像を配置する場合
          if (itemImageCells.itemImageStartCell && itemImageCells.itemImageEndCell) {
            image.topLeftCornerCell = workbook
              .worksheets(0)
              .getCell(itemImageCells.itemImageStartCell, CellReferenceMode.A1);
            image.bottomRightCornerCell = workbook
              .worksheets(0)
              .getCell(itemImageCells.itemImageEndCell, CellReferenceMode.A1);
          }
          // ▲▲▲ 画像の配置ポジション指定処理 ココマデ ▲▲▲

          getActualImage(result as string).then((actualImage: HTMLImageElement) => {

            // ▼▼▼ 画像の縦横サイズ指定処理 ココカラ ▼▼▼
            const myBoundsInTwips: IgRect = image.getBoundsInTwips();
            const containerRatio = myBoundsInTwips.width / myBoundsInTwips.height;
            const imageRatio = actualImage.width / actualImage.height;

            if (imageRatio < containerRatio) { // height を基準にする
              myBoundsInTwips.width = myBoundsInTwips.height * (actualImage.width / actualImage.height);
            } else { // width を基準にする
              myBoundsInTwips.height = myBoundsInTwips.width * (actualImage.height / actualImage.width);
            }
            image.setBoundsInTwips(workbook.worksheets(0), myBoundsInTwips);
            // ▲▲▲ 画像の縦横サイズ指定処理 ココマデ ▲▲▲

            workbook.worksheets(0).shapes().add(image); // 作成した画像データの追加

            // 画像用プレースホルダーの削除
            if (itemImageCells.itemImageMergedCell) {
              workbook.worksheets(0).getCell(itemImageCells.itemImageMergedCell).value = "";
            }
            if (itemImageCells.itemImageStartCell && itemImageCells.itemImageEndCell) {
              workbook.worksheets(0).getCell(itemImageCells.itemImageStartCell, CellReferenceMode.A1).value = "";
              workbook.worksheets(0).getCell(itemImageCells.itemImageEndCell, CellReferenceMode.A1).value = "";
            }

          });
        });
      });
    };

    // Workbookをxlsxファイル形式で保存
    saveWorkbook = (
      workbook: Workbook,
      fileNameWithoutExtension: string
    ): Promise<string> => {
      return new Promise<string>( (resolve, reject) => {
        const opt = new WorkbookSaveOptions();
        opt.type = "blob";

        workbook.save(
          opt,
          (d) => {
            const fileExt = ".xlsx";
            const fileName = fileNameWithoutExtension + fileExt;
            saveAs(d as Blob, fileName);
            resolve(fileName);
          },
          (e) => {
            reject(e);
          }
        );
      });
    };

    // refを設定するためにIgnSpreadsheetというコンポーネントを作成
    const IgnSpreadsheetComponents = ({
      spreadsheetRef,
      ...props
    }: IIgnSpreadsheetProps) => {
      return <IgrSpreadsheet ref={spreadsheetRef} {...props} />;
    };

    return IgnSpreadsheetComponents;
  },
  { ssr: false }
);
