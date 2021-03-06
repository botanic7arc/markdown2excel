# MarkdownExcel変換ソフト試験書 担当者:xx

## 関数試験:GetInputFileName

### mdファイル確認処理

#### 正常

##### 状態:mdファイルが1つ存在->O:4/11@

1. mdファイルを1つ試験用ディレクトリに配置する。
2. ビルドした実行ファイルを同ディレクトリに配置する。
3. 関数を実行する。
- {配置したmdファイル名}が表示されていること。

##### 状態:mdファイルが3つ存在->O:4/11

1. mdファイルを3つ試験用ディレクトリに配置する。
2. ビルドした実行ファイルを同ディレクトリに配置する。
3. 関数を実行する。
- {配置したmdファイル名1}が表示されていること。
- {配置したmdファイル名2}が表示されていること。
- {配置したmdファイル名3}が表示されていること。

#### 異常

##### 状態:mdファイルが存在しない->O:4/11@4/10時点では使用者にとってわかりにくいエラーでした

1. 試験用ディレクトリからmdファイルを削除する。
2. ビルドした実行ファイルを同ディレクトリに配置する。
3. 関数を実行する。
- " mdファイルが同ディレクトリに存在しません(改行)実行ファイルをmdファイルが存在しているディレクトリに配置してください"が表示されていること。

### 入力確認処理

#### 正常

##### 入力:3->O:4/11

1. mdファイルを5個試験用ディレクトリに配置する。
2. 関数を実行する。
3. ファイルに対応する数字入力画面で"3"を入力する。
- 選択したファイル名が返されていること。

#### 異常

##### 入力:範囲外の数値->X:4/11

1. mdファイルを5個試験用ディレクトリに配置する。
2. 関数を実行する。
3. ファイルに対応する数字入力画面で"20"を入力する。
- "getInputFileName Error : GetInputFileName(Out of range value):20"が表示されていること。

##### 入力:数値以外->O:4/11

1. mdファイルを5個試験用ディレクトリに配置する。
2. 関数を実行する。
3. ファイルに対応する数字入力画面で"b"を入力する。
- "GetInputFileName(Input is not a number):b"が表示されていること。

## 関数試験:SetConstContent

### シート設定処理

#### 正常

##### 入力:A(試験概要)->O:4/11

1. mdファイルを5個試験用ディレクトリに配置する。
2. 関数を実行する。
3. ファイルに対応する数字入力画面で"1"を入力する。
- 作成されたExcelファイルのシートが1つで名前が"試験書"となっていること。

### 固定値セル設定処理

#### 正常

##### 入力:正しいセル書式->O:4/11

1. mdファイルを5個試験用ディレクトリに配置する。
2. const.goでvar constContent = [][]string{{"F1","テスト"}} //{{"CELL","VALUE"}} と設定する。
3. 関数を実行する。
4. ファイルに対応する数字入力画面で"1"を入力する。
- F1に"テスト"と設定されていること。

#### 異常

##### 入力:不正なセル書式->O:4/11

1. mdファイルを5個試験用ディレクトリに配置する。
2. const.goでvar constContent = [][]string{{"1F","テスト"}} //{{"CELL","VALUE"}} と設定する。
3. 関数を実行する。
4. ファイルに対応する数字入力画面で"1"を入力する。
- "SetConstContent : cannot convert cell "1F" to coordinates: invalid cell name "1F""と表示されていること。
- 変換は完了していること。

## 動作試験

### Excel作成処理

#### 正常

##### 入力:形式を満たしたmdファイル->O:4/11

1. mdファイルを5個試験用ディレクトリに配置する。
2. 本ソフトウェアを実行する。
3. ファイルに対応する数字入力画面で"3"を入力する。
- mdファイルを元にしたExcelファイルが作成されていること。

##### 入力:空のmdファイル->O:4/11

1. 空のmdファイルを試験用ディレクトリに配置する。
2. 本ソフトウェアを実行する。
3. ファイルに対応する数字入力画面で空のmdファイルに対応する数字を入力する。
- 空のExcelファイルが作成されていること。

#### 異常

##### 状態:変換後の名前と同名のExcelファイルを開いている状態->O:4/11

1. mdファイルを5個試験用ディレクトリに配置する。
2. 本ソフトウェアを実行する。
3. ファイルに対応する数字入力画面で"3"を入力する。
4. 作成されたExcelファイルを開く。
5. 本ソフトウェアを実行する。
6. ファイルに対応する数字入力画面で"3"を入力する。
- "open [変換後Excelファイル名]: permission denied"と表示されていること。
- Excelファイルが更新されていないこと。