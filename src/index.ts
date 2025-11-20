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

import { SheetMerger } from './SheetMerger';

/**
 * メニューに表示される関数: データを統合
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function mergeSheets(): void {
  const merger = new SheetMerger();
  merger.mergeSheets();
}

/**
 * メニューに表示される関数: 設定シートのテンプレートを作成
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function createConfigTemplate(): void {
  const merger = new SheetMerger();
  merger.createConfigTemplate();
}

/**
 * メニューに表示される関数: JSON形式でデータを出力
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function exportToJson(): void {
  const merger = new SheetMerger();
  merger.exportToJson();
}

/**
 * スプレッドシートを開いたときに実行される関数
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('シート統合ツール')
    .addItem('設定テンプレート作成', 'createConfigTemplate')
    .addItem('データ統合実行', 'mergeSheets')
    .addItem('JSON出力', 'exportToJson')
    .addToUi();
}
