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

import { SheetMerger } from '../src/SheetMerger';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type MockType = any;

describe('SheetMerger', () => {
  let mockSpreadsheet: MockType;
  let mockUi: MockType;

  beforeEach(() => {
    // Mock SpreadsheetApp
    mockUi = {
      alert: jest.fn(),
    };

    mockSpreadsheet = {
      getSheetByName: jest.fn(),
      insertSheet: jest.fn(),
    };

    global.SpreadsheetApp = {
      getActiveSpreadsheet: jest.fn(() => mockSpreadsheet),
      getUi: jest.fn(() => mockUi),
    } as MockType;

    global.Logger = {
      log: jest.fn(),
    } as MockType;
  });

  describe('createConfigTemplate', () => {
    it('設定シートが存在しない場合、新しいシートを作成する', () => {
      const mockSheet = {
        getRange: jest.fn().mockReturnThis(),
        setValues: jest.fn().mockReturnThis(),
        setFontWeight: jest.fn().mockReturnThis(),
        setColumnWidth: jest.fn(),
      };

      mockSpreadsheet.getSheetByName.mockReturnValue(null);
      mockSpreadsheet.insertSheet.mockReturnValue(mockSheet);

      const merger = new SheetMerger();
      merger.createConfigTemplate();

      expect(mockSpreadsheet.insertSheet).toHaveBeenCalledWith('sm.settings');
      expect(mockSheet.setValues).toHaveBeenCalled();
      expect(mockUi.alert).toHaveBeenCalledWith(
        '設定シートのテンプレートを作成しました。'
      );
    });

    it('設定シートが既に存在する場合、クリアして再作成する', () => {
      const mockSheet = {
        clear: jest.fn(),
        getRange: jest.fn().mockReturnThis(),
        setValues: jest.fn().mockReturnThis(),
        setFontWeight: jest.fn().mockReturnThis(),
        setColumnWidth: jest.fn(),
      };

      mockSpreadsheet.getSheetByName.mockReturnValue(mockSheet);

      const merger = new SheetMerger();
      merger.createConfigTemplate();

      expect(mockSheet.clear).toHaveBeenCalled();
      expect(mockSheet.setValues).toHaveBeenCalled();
    });
  });

  describe('mergeSheets', () => {
    it('設定シートが存在しない場合、エラーをスローする', () => {
      mockSpreadsheet.getSheetByName.mockReturnValue(null);

      const merger = new SheetMerger();

      expect(() => merger.mergeSheets()).toThrow(
        '設定シート "sm.settings" が見つかりません。'
      );
    });

    it('外部キーカラム名が設定されていない場合、エラーをスローする', () => {
      const mockConfigSheet = {
        getDataRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([['外部カラム', '']]),
        }),
      };

      mockSpreadsheet.getSheetByName.mockReturnValue(mockConfigSheet);

      const merger = new SheetMerger();

      expect(() => merger.mergeSheets()).toThrow(
        '外部キーカラム名が設定されていません。'
      );
    });

    it('データソースが設定されていない場合、エラーをスローする', () => {
      const mockConfigSheet = {
        getDataRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([
            ['外部カラム', 'id'],
            ['探索範囲(行)', '10'],
            ['', ''],
            ['シート名', 'カラム名'],
            // データソースの行がない
          ]),
        }),
      };

      mockSpreadsheet.getSheetByName.mockImplementation((name: string) => {
        if (name === 'sm.settings') return mockConfigSheet;
        return null;
      });

      const merger = new SheetMerger();

      // readConfigでエラーがスローされることを確認
      expect(() => merger.mergeSheets()).toThrow(
        'データソースが設定されていません。'
      );
    });

    it('正常なデータで統合が成功する', () => {
      const mockConfigSheet = {
        getDataRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([
            ['外部カラム', 'id'],
            ['探索範囲(行)', '10'],
            ['', ''],
            ['シート名', 'カラム名'],
            ['ユーザー', 'name'],
            ['ユーザー', 'email'],
            ['注文', 'amount'],
          ]),
        }),
      };

      const mockUserSheet = {
        getDataRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([
            ['id', 'name', 'email'],
            [1, '山田', 'yamada@example.com'],
            [2, '佐藤', 'sato@example.com'],
          ]),
        }),
      };

      const mockOrderSheet = {
        getDataRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([
            ['id', 'amount'],
            [1, 10000],
            [2, 25000],
          ]),
        }),
      };

      const mockOutputSheet = {
        clear: jest.fn(),
        getRange: jest.fn().mockReturnThis(),
        setValues: jest.fn(),
        setFontWeight: jest.fn(),
      };

      mockSpreadsheet.getSheetByName.mockImplementation((name: string) => {
        if (name === 'sm.settings') return mockConfigSheet;
        if (name === 'ユーザー') return mockUserSheet;
        if (name === '注文') return mockOrderSheet;
        if (name === 'sm.integrated_data') return mockOutputSheet;
        return null;
      });

      const merger = new SheetMerger();
      merger.mergeSheets();

      expect(mockOutputSheet.setValues).toHaveBeenCalled();
      expect(mockUi.alert).toHaveBeenCalledWith('データの統合が完了しました。');
    });

    it('ヘッダー行が途中から始まる場合も正しく統合できる', () => {
      const mockConfigSheet = {
        getDataRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([
            ['外部カラム', 'id'],
            ['探索範囲(行)', '10'],
            ['', ''],
            ['シート名', 'カラム名'],
            ['ユーザー', 'name'],
          ]),
        }),
      };

      const mockUserSheet = {
        getDataRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([
            ['', '', ''], // 空行
            ['メモ', 'これはテストデータです', ''], // メモ行
            ['id', 'name', 'age'], // ヘッダー行が3行目
            [1, '山田', 30],
            [2, '佐藤', 25],
          ]),
        }),
      };

      const mockOutputSheet = {
        clear: jest.fn(),
        getRange: jest.fn().mockReturnThis(),
        setValues: jest.fn(),
        setFontWeight: jest.fn(),
      };

      mockSpreadsheet.getSheetByName.mockImplementation((name: string) => {
        if (name === 'sm.settings') return mockConfigSheet;
        if (name === 'ユーザー') return mockUserSheet;
        if (name === 'sm.integrated_data') return mockOutputSheet;
        return null;
      });

      const merger = new SheetMerger();
      merger.mergeSheets();

      expect(mockOutputSheet.setValues).toHaveBeenCalled();
      const callArgs = mockOutputSheet.setValues.mock.calls[0][0];
      // 1行目: シート名
      expect(callArgs[0][0]).toEqual('id');
      expect(callArgs[0]).toContain('ユーザー');
      // 2行目: カラム名
      expect(callArgs[1][0]).toEqual('id');
      expect(callArgs[1]).toContain('name');
      // 3行目: データ行1
      expect(callArgs[2][0]).toEqual('1');
      expect(callArgs[2]).toContain('山田');
      // 4行目: データ行2
      expect(callArgs[3][0]).toEqual('2');
      expect(callArgs[3]).toContain('佐藤');
    });
  });

  describe('exportToJson', () => {
    it('JSON形式でダイアログに出力する', () => {
      const mockConfigSheet = {
        getDataRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([
            ['外部カラム', 'id'],
            ['探索範囲(行)', '10'],
            ['', ''],
            ['シート名', 'カラム名'],
            ['ユーザー', 'name'],
            ['ユーザー', 'email'],
          ]),
        }),
      };

      const mockUserSheet = {
        getDataRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([
            ['id', 'name', 'email'],
            [1, '山田', 'yamada@example.com'],
            [2, '佐藤', 'sato@example.com'],
          ]),
        }),
      };

      const mockHtmlOutput = {
        setWidth: jest.fn().mockReturnThis(),
        setHeight: jest.fn().mockReturnThis(),
      };

      mockUi.showModalDialog = jest.fn();

      global.HtmlService = {
        createHtmlOutput: jest.fn().mockReturnValue(mockHtmlOutput),
      } as MockType;

      mockSpreadsheet.getSheetByName.mockImplementation((name: string) => {
        if (name === 'sm.settings') return mockConfigSheet;
        if (name === 'ユーザー') return mockUserSheet;
        return null;
      });

      const merger = new SheetMerger();
      merger.exportToJson();

      expect(global.HtmlService.createHtmlOutput).toHaveBeenCalled();
      expect(mockHtmlOutput.setWidth).toHaveBeenCalledWith(700);
      expect(mockHtmlOutput.setHeight).toHaveBeenCalledWith(600);
      expect(mockUi.showModalDialog).toHaveBeenCalledWith(
        mockHtmlOutput,
        'JSON出力'
      );

      // HTMLに含まれるJSONを検証
      const htmlContent = (global.HtmlService.createHtmlOutput as jest.Mock)
        .mock.calls[0][0] as string;
      expect(htmlContent).toContain('<textarea');
      expect(htmlContent).toContain('readonly');

      // HTMLエンティティをデコードしてJSON検証
      const match = htmlContent.match(/<textarea[^>]*>([\s\S]*?)<\/textarea>/);
      expect(match).toBeTruthy();

      if (match) {
        const jsonString = match[1]
          .replace(/&#39;/g, "'")
          .replace(/&quot;/g, '"')
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .replace(/&amp;/g, '&');
        const parsedJson = JSON.parse(jsonString);
        expect(Array.isArray(parsedJson)).toBe(true);
        expect(parsedJson.length).toBe(2);
        expect(parsedJson[0]).toHaveProperty('id');
        expect(parsedJson[0]).toHaveProperty('ユーザー');
        expect(parsedJson[0].ユーザー).toHaveProperty('name');
        expect(parsedJson[0].ユーザー).toHaveProperty('email');
      }
    });
  });

  describe('exportSelectedToJson', () => {
    it('選択された主キーのみをJSON形式で出力する', () => {
      const mockConfigSheet = {
        getDataRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([
            ['外部カラム', 'id'],
            ['探索範囲(行)', '10'],
            ['', ''],
            ['シート名', 'カラム名'],
            ['ユーザー', 'name'],
            ['ユーザー', 'email'],
          ]),
        }),
      };

      const mockUserSheet = {
        getDataRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([
            ['id', 'name', 'email'],
            [1, '山田', 'yamada@example.com'],
            [2, '佐藤', 'sato@example.com'],
            [3, '鈴木', 'suzuki@example.com'],
          ]),
        }),
      };

      const mockActiveSheet = {
        getActiveRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([[1], [3]]), // idが1と3を選択
        }),
      };

      const mockHtmlOutput = {
        setWidth: jest.fn().mockReturnThis(),
        setHeight: jest.fn().mockReturnThis(),
      };

      mockUi.showModalDialog = jest.fn();

      global.HtmlService = {
        createHtmlOutput: jest.fn().mockReturnValue(mockHtmlOutput),
      } as MockType;

      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn(() => mockSpreadsheet),
        getUi: jest.fn(() => mockUi),
        getActiveSheet: jest.fn(() => mockActiveSheet),
      } as MockType;

      mockSpreadsheet.getSheetByName.mockImplementation((name: string) => {
        if (name === 'sm.settings') return mockConfigSheet;
        if (name === 'ユーザー') return mockUserSheet;
        return null;
      });

      const merger = new SheetMerger();
      merger.exportSelectedToJson();

      expect(mockUi.showModalDialog).toHaveBeenCalled();

      // HTMLに含まれるJSONを検証
      const htmlContent = (global.HtmlService.createHtmlOutput as jest.Mock)
        .mock.calls[0][0] as string;
      const match = htmlContent.match(/<textarea[^>]*>([\s\S]*?)<\/textarea>/);
      expect(match).toBeTruthy();

      if (match) {
        const jsonString = match[1]
          .replace(/&#39;/g, "'")
          .replace(/&quot;/g, '"')
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .replace(/&amp;/g, '&');
        const parsedJson = JSON.parse(jsonString);

        // 選択された2つのIDのみがJSON出力される
        expect(Array.isArray(parsedJson)).toBe(true);
        expect(parsedJson.length).toBe(2);
        expect(parsedJson[0].id).toBe('1');
        expect(parsedJson[0].ユーザー.name).toBe('山田');
        expect(parsedJson[1].id).toBe('3');
        expect(parsedJson[1].ユーザー.name).toBe('鈴木');
      }
    });

    it('選択範囲がない場合、エラーを表示する', () => {
      const mockConfigSheet = {
        getDataRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([
            ['外部カラム', 'id'],
            ['探索範囲(行)', '10'],
            ['', ''],
            ['シート名', 'カラム名'],
            ['ユーザー', 'name'],
          ]),
        }),
      };

      const mockUserSheet = {
        getDataRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([
            ['id', 'name'],
            [1, '山田'],
          ]),
        }),
      };

      const mockActiveSheet = {
        getActiveRange: jest.fn().mockReturnValue({
          getValues: jest.fn().mockReturnValue([['']]), // 空の選択
        }),
      };

      mockUi.ButtonSet = { OK: 'OK' };

      global.SpreadsheetApp = {
        getActiveSpreadsheet: jest.fn(() => mockSpreadsheet),
        getUi: jest.fn(() => mockUi),
        getActiveSheet: jest.fn(() => mockActiveSheet),
      } as MockType;

      mockSpreadsheet.getSheetByName.mockImplementation((name: string) => {
        if (name === 'sm.settings') return mockConfigSheet;
        if (name === 'ユーザー') return mockUserSheet;
        return null;
      });

      const merger = new SheetMerger();
      merger.exportSelectedToJson();

      expect(mockUi.alert).toHaveBeenCalledWith(
        '選択エラー',
        '主キー「id」を含むセルを選択してください。',
        'OK'
      );
    });
  });
});
