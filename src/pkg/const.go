package pkg

//Markdownの内容に依らない固定値群
var rowWidth = []float64{5, 25, 15, 10, 18, 56, 50, 7, 9, 30}
var tableHeader = []string{"No", "大項目", "中項目", "正常/異常", "試験概要", "試験方法", "確認内容", "結果", "実施日", "備考"}
var constContent = [][]string{{}} //{{"CELL","VALUE"}}

const sheetName = "試験書"
const parseChara = "->"       //試験項目と確認内容を分割する文字
const startHeaderColumn = 'A' //見出し開始列
const startHeaderRow = 3      //見出し開始行(2以上)

//コンソール表示内容
const lang = 1 //0:English, 1:Japanese
var dialog = [][]string{{"===select input file num===", "input file num:", "Conversion completed,", "output File:"}, {"===変換するファイルに対応する数字を入力してください===", "変換ファイル番号:", "変換完了", "出力ファイル名:"}}
