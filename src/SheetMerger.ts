/**
 * Copyright 2025 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * 設定シートの構造
 */
interface MergeConfig {
  foreignKeyColumn: string;
  searchRange: number;
  dataSources: DataSource[];
}

interface DataSource {
  sheetName: string;
  columns: string[];
}

/**
 * シート統合クラス
 * 外部キーで各シートに分割されているデータを一つのシートに統合する
 */
export class SheetMerger {
  private spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  private configSheetName: string;
  private outputSheetName: string;

  /**
   * @param configSheetName 設定シート名(デフォルト: "sm.settings")
   * @param outputSheetName 出力シート名(デフォルト: "sm.integrated_data")
   */
  constructor(
    configSheetName = 'sm.settings',
    outputSheetName = 'sm.integrated_data'
  ) {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    this.configSheetName = configSheetName;
    this.outputSheetName = outputSheetName;
  }

  /**
   * メイン処理: 設定を読み込んでデータを統合
   */
  mergeSheets(): void {
    try {
      const config = this.readConfig();
      const mergedData = this.collectData(config);
      this.writeOutput(mergedData, config);
      SpreadsheetApp.getUi().alert('データの統合が完了しました。');
    } catch (error) {
      SpreadsheetApp.getUi().alert(
        `エラーが発生しました: ${(error as Error).message}`
      );
      throw error;
    }
  }

  /**
   * 設定シートから設定を読み込む
   */
  private readConfig(): MergeConfig {
    const configSheet = this.spreadsheet.getSheetByName(this.configSheetName);
    if (!configSheet) {
      throw new Error(
        `設定シート "${this.configSheetName}" が見つかりません。`
      );
    }

    const data = configSheet.getDataRange().getValues();

    // 外部キーカラム名を取得(B1セル)
    const foreignKeyColumn = data[0]?.[1] as string;
    if (!foreignKeyColumn) {
      throw new Error('外部キーカラム名が設定されていません。');
    }

    // 探索範囲を取得(B2セル)、デフォルトは10行
    const searchRangeValue = data[1]?.[1];
    const searchRange =
      searchRangeValue && !isNaN(Number(searchRangeValue))
        ? Number(searchRangeValue)
        : 10;

    // データソース設定を読み込む(4行目以降、インデックス4から)
    const dataSourceMap = new Map<string, string[]>();
    for (let i = 4; i < data.length; i++) {
      const sheetName = data[i][0] as string;
      const columnName = data[i][1] as string;

      if (sheetName && columnName) {
        if (!dataSourceMap.has(sheetName)) {
          dataSourceMap.set(sheetName, []);
        }
        dataSourceMap.get(sheetName)!.push(columnName.trim());
      }
    }

    const dataSources: DataSource[] = [];
    for (const [sheetName, columns] of dataSourceMap) {
      dataSources.push({ sheetName, columns });
    }

    if (dataSources.length === 0) {
      throw new Error('データソースが設定されていません。');
    }

    return { foreignKeyColumn, searchRange, dataSources };
  }

  /**
   * 各シートからデータを収集して統合
   */
  private collectData(
    config: MergeConfig
  ): Map<string, Map<string, string | number>> {
    const mergedData = new Map<string, Map<string, string | number>>();

    for (const source of config.dataSources) {
      const sheet = this.spreadsheet.getSheetByName(source.sheetName);
      if (!sheet) {
        Logger.log(
          `警告: シート "${source.sheetName}" が見つかりません。スキップします。`
        );
        continue;
      }

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) {
        Logger.log(
          `警告: シート "${source.sheetName}" にデータがありません。スキップします。`
        );
        continue;
      }

      // 外部キーカラムを探索範囲内で探索
      let headerRowIndex = -1;
      let foreignKeyIndex = -1;
      const searchLimit = Math.min(config.searchRange, data.length);

      for (let i = 0; i < searchLimit; i++) {
        const row = data[i] as string[];
        const index = row.findIndex(
          col => String(col).trim() === config.foreignKeyColumn
        );
        if (index !== -1) {
          headerRowIndex = i;
          foreignKeyIndex = index;
          break;
        }
      }

      if (headerRowIndex === -1 || foreignKeyIndex === -1) {
        Logger.log(
          `警告: シート "${source.sheetName}" の探索範囲(${config.searchRange}行)内に外部キーカラム "${config.foreignKeyColumn}" が見つかりません。`
        );
        continue;
      }

      const headers = data[headerRowIndex] as string[];

      // 対象カラムのインデックスを取得
      const columnIndices = source.columns
        .map(col => {
          const index = headers.indexOf(col);
          if (index === -1) {
            Logger.log(
              `警告: シート "${source.sheetName}" にカラム "${col}" が見つかりません。`
            );
            return null;
          }
          return { name: col, index };
        })
        .filter(
          (item): item is { name: string; index: number } => item !== null
        );

      // データ行を処理(ヘッダーの次の行から)
      for (let i = headerRowIndex + 1; i < data.length; i++) {
        const row = data[i];
        const foreignKey = String(row[foreignKeyIndex]);

        if (!foreignKey) {
          continue;
        }

        // 外部キーごとのデータマップを取得または作成
        if (!mergedData.has(foreignKey)) {
          mergedData.set(foreignKey, new Map());
        }
        const rowData = mergedData.get(foreignKey)!;

        // 外部キーを保存
        rowData.set(config.foreignKeyColumn, foreignKey);

        // 各カラムのデータを保存
        for (const { name, index } of columnIndices) {
          const columnKey = `${source.sheetName}.${name}`;
          rowData.set(columnKey, row[index]);
        }
      }
    }

    return mergedData;
  }

  /**
   * 統合データを出力シートに書き込む
   */
  private writeOutput(
    mergedData: Map<string, Map<string, string | number>>,
    config: MergeConfig
  ): void {
    // 出力シートを取得または作成
    let outputSheet = this.spreadsheet.getSheetByName(this.outputSheetName);
    if (outputSheet) {
      outputSheet.clear();
    } else {
      outputSheet = this.spreadsheet.insertSheet(this.outputSheetName);
    }

    // 1行目: シート名行を作成
    const sheetNameRow: string[] = [config.foreignKeyColumn];
    for (const source of config.dataSources) {
      for (let i = 0; i < source.columns.length; i++) {
        sheetNameRow.push(source.sheetName);
      }
    }

    // 2行目: カラム名行を作成
    const columnNameRow: string[] = [config.foreignKeyColumn];
    for (const source of config.dataSources) {
      for (const column of source.columns) {
        columnNameRow.push(column);
      }
    }

    // ヘッダーキーを作成(データ取得用)
    const headerKeys: string[] = [config.foreignKeyColumn];
    for (const source of config.dataSources) {
      for (const column of source.columns) {
        headerKeys.push(`${source.sheetName}.${column}`);
      }
    }

    // データ行を作成
    const outputData: (string | number)[][] = [sheetNameRow, columnNameRow];
    for (const rowData of mergedData.values()) {
      const row: (string | number)[] = [];
      for (const headerKey of headerKeys) {
        row.push(rowData.get(headerKey) ?? '');
      }
      outputData.push(row);
    }

    // シートに書き込み
    if (outputData.length > 0) {
      outputSheet
        .getRange(1, 1, outputData.length, outputData[0].length)
        .setValues(outputData);

      // 1行目と2行目を太字にする
      outputSheet.getRange(1, 1, 2, outputData[0].length).setFontWeight('bold');
    }
  }

  /**
   * 設定シートのテンプレートを作成
   */
  createConfigTemplate(): void {
    let configSheet = this.spreadsheet.getSheetByName(this.configSheetName);

    if (!configSheet) {
      configSheet = this.spreadsheet.insertSheet(this.configSheetName);
    } else {
      configSheet.clear();
    }

    const template = [
      ['外部カラム', 'id'],
      ['探索範囲(行)', '10'],
      ['', ''],
      ['シート名', 'カラム名'],
      ['ユーザー', 'name'],
      ['ユーザー', 'email'],
      ['ユーザー', 'department'],
      ['注文', 'order_date'],
      ['注文', 'amount'],
      ['注文', 'status'],
    ];

    configSheet.getRange(1, 1, template.length, 2).setValues(template);
    configSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    configSheet.getRange(3, 1, 1, 2).setFontWeight('bold');
    configSheet.setColumnWidth(1, 200);
    configSheet.setColumnWidth(2, 400);

    SpreadsheetApp.getUi().alert('設定シートのテンプレートを作成しました。');
  }
}
