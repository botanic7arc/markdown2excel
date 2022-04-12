package pkg

type SheetData struct {
	sheetName           string
	rowWidth            []float64
	testItemNum         int  //試験項目数
	startHeaderColumn   rune //見出し開始列
	startHeaderRow      int  //見出し開始行(2以上)
	tableHeader         []string
	constContent        [][]string
	parseChara          string //試験項目と確認内容を分割する文字
	bigItemBorderRow    []int
	middleItemBorderRow []int
	smallItemBorderRow  []int
	procedureStr        string
	confirmStr          string
	procedureCell       string
	confirmCell         string
}

func NewSheetData() *SheetData {
	s := new(SheetData)
	s.sheetName = sheetName
	s.rowWidth = append(s.rowWidth, rowWidth...)
	s.testItemNum = 0
	s.startHeaderColumn = startHeaderColumn
	s.startHeaderRow = startHeaderRow
	s.tableHeader = append(s.tableHeader, tableHeader...)
	if 0 < len(constContent) {
		s.constContent = append(s.constContent, constContent...)
	}
	s.parseChara = parseChara
	return s
}
