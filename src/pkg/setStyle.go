package pkg

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

//列の項目設定
func SetTableHeader(excel *excelize.File) {
	c := startHeaderColumn
	for _, v := range tableHeader {
		cell := fmt.Sprintf("%s%d", string(c), startHeaderRow)
		err := excel.SetCellValue(sheetName, cell, v)
		if err != nil {
			log.Printf("SetTableHeader Error : %s", err)
		}
		c++ //'A'->'B', 'B'->'C'...
	}
}

//列の幅を設定
func SetRowWidth(excel *excelize.File) {
	c := startHeaderColumn
	for _, v := range rowWidth {
		row := string(c)
		err := excel.SetColWidth(sheetName, row, row, v)
		if err != nil {
			log.Printf("SetRowWidth Error : %s", err)
		}
		c++ //'A'->'B', 'B'->'C'...
	}
}

//セルのスタイル設定
func SetCellStyle(excel *excelize.File, s *SheetData) {
	//見出し部分の色と罫線
	header := &excelize.Style{
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#B0C4DE"},
			Pattern: 1},
	}
	headerCellStart := fmt.Sprintf("%s%d", string(startHeaderColumn), startHeaderRow)
	headerCellEndColumn := startHeaderColumn + len(tableHeader) - 1
	headerCellEnd := fmt.Sprintf("%s%d", string(rune(headerCellEndColumn)), startHeaderRow)
	StyleHeader, _ := excel.NewStyle(header)
	excel.SetCellStyle(sheetName, headerCellStart, headerCellEnd, StyleHeader)
	//No・試験概要・試験項目・結果・実施日・備考の網罫線
	item := &excelize.Style{
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "top",
			WrapText:   true,
		},
	}
	styleItem, _ := excel.NewStyle(item)
	noCellStart := fmt.Sprintf("%s%d", string(startHeaderColumn), startHeaderRow+1)
	noCellEnd := fmt.Sprintf("%s%d", string(startHeaderColumn), s.testItemNum+startHeaderRow)
	excel.SetCellStyle(sheetName, noCellStart, noCellEnd, styleItem)
	testCellStart := fmt.Sprintf("%s%d", string(startHeaderColumn+3), 1+startHeaderRow)
	testCellEnd := fmt.Sprintf("%s%d", string(rune(headerCellEndColumn)), s.testItemNum+startHeaderRow)
	excel.SetCellStyle(sheetName, testCellStart, testCellEnd, styleItem)

	//大項目の区切り下線
	bottomBorder := &excelize.Style{
		Border: []excelize.Border{
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "top",
			WrapText:   true,
		},
	}
	styleBottomBorder, _ := excel.NewStyle(bottomBorder)

	//大項目のテキスト整形
	textAlign := &excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "top",
			WrapText:   true,
		},
	}
	styleTextAlign, _ := excel.NewStyle(textAlign)

	for i, item := range s.bigItemBorderRow {
		cell := fmt.Sprintf("%s%d", string(startHeaderColumn+1), startHeaderRow+item)
		excel.SetCellStyle(sheetName, cell, cell, styleBottomBorder)
		var textCell string
		if i == 0 {
			textCell = fmt.Sprintf("%s%d", string(startHeaderColumn+1), startHeaderRow+1)
		} else if 0 < i && i < len(s.bigItemBorderRow) {
			textCell = fmt.Sprintf("%s%d", string(startHeaderColumn+1), startHeaderRow+s.bigItemBorderRow[i-1]+1)
		}
		excel.SetCellStyle(sheetName, textCell, textCell, styleTextAlign)

	}

	//中項目・小項目の区切り左線または下と左線
	bottomLeftBorder := &excelize.Style{
		Border: []excelize.Border{
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "top",
			WrapText:   true,
		},
	}
	styleBottomLeftBorder, _ := excel.NewStyle(bottomLeftBorder)

	leftBorder := &excelize.Style{
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "top",
			WrapText:   true,
		},
	}
	styleLeftBorder, _ := excel.NewStyle(leftBorder)

	j := 0
	k := 0
	for i := 0; i < s.testItemNum; i++ {
		cell1 := fmt.Sprintf("%s%d", string(startHeaderColumn+2), i+startHeaderRow+1)
		//現在のセルと項目の区切りが一致する場合は適用させるスタイルを変更する
		if j < len(s.middleItemBorderRow) && i+1 == s.middleItemBorderRow[j] {
			excel.SetCellStyle(sheetName, cell1, cell1, styleBottomLeftBorder)
			j++
		} else {
			excel.SetCellStyle(sheetName, cell1, cell1, styleLeftBorder)
		}
		cell2 := fmt.Sprintf("%s%d", string(startHeaderColumn+3), i+startHeaderRow+1)
		//現在のセルと項目の区切りが一致する場合は適用させるスタイルを変更する
		if k < len(s.smallItemBorderRow) && i+1 == s.smallItemBorderRow[k] {
			excel.SetCellStyle(sheetName, cell2, cell2, styleBottomLeftBorder)
			k++
		} else {
			excel.SetCellStyle(sheetName, cell2, cell2, styleLeftBorder)
		}
	}

}
