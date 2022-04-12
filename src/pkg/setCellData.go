package pkg

import (
	"bufio"
	"fmt"
	"io"
	"log"
	"os"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

func ReadFile(excel *excelize.File, fileName string, s *SheetData) *excelize.File {
	//ファイル読み込み
	file, err := os.Open(fileName)
	if err != nil {
		log.Printf("ReadFile(Open) Error: %s", err)
	}
	reader := bufio.NewReader(file)
	//1行ずつ処理
	for {
		line, err := reader.ReadString('\n')
		line = strings.TrimRight(line, "\n")
		SetCellPosition(excel, s, line)
		if err != nil {
			if err == io.EOF {
				break
			}
			log.Printf("ReadFile(ReadString) Error : %s", err)
		}
	}
	//最終行
	if s.confirmCell != "" && s.confirmStr != "" {
		excel.SetCellValue(sheetName, s.confirmCell, s.confirmStr)
		s.confirmCell = ""
		s.confirmStr = ""
		s.testItemNum++
	}
	s.bigItemBorderRow = append(s.bigItemBorderRow, s.testItemNum)
	s.middleItemBorderRow = append(s.middleItemBorderRow, s.testItemNum)
	s.smallItemBorderRow = append(s.smallItemBorderRow, s.testItemNum)
	//No(項番)追加
	for i := 0; i < s.testItemNum; i++ {
		cell := fmt.Sprintf("%s%d", string(startHeaderColumn), i+startHeaderRow+1)
		excel.SetCellValue(sheetName, cell, i+1)
	}
	return excel
}

func SetCellPosition(excel *excelize.File, s *SheetData, line string) {
	mark := strings.Split(line, " ")
	if 1 < len(mark[0]) && mark[0][1] == '.' {
		mark[0] = "."
	}
	switch mark[0] {
	case "#": //# -> 左上タイトル
		excel.SetCellValue(sheetName, "A1", mark[1])
		excel.SetCellValue(sheetName, "C1", mark[2])
	case "##": //## -> 大項目
		SetCellData(excel, s, mark[1], 1)
	case "###": //### -> 中項目
		SetCellData(excel, s, mark[1], 2)
	case "####": //#### -> 小項目
		SetCellData(excel, s, mark[1], 3)
	case "#####": //##### -> 小項目
		SetCellData(excel, s, mark[1], 4)
	case ".": //- ->試験内容
		SetCellData(excel, s, line, 5)
	case "-": //- 確認項目
		SetCellData(excel, s, line, 6)
	default:
	}
}

func SetCellData(excel *excelize.File, s *SheetData, line string, column int) {
	//試験内容行(.md)
	if column == 5 {
		if len(s.procedureStr) != 0 {
			s.procedureStr = fmt.Sprintf("%s\n", s.procedureStr)
		}
		s.procedureStr += line
		if len(s.procedureCell) == 0 {
			s.procedureCell = fmt.Sprintf("%s%d", string(startHeaderColumn+rune(column)), startHeaderRow+1+s.testItemNum)
		}
		return
	}
	//確認内容行(.md)
	if column == 6 {
		excel.SetCellValue(sheetName, s.procedureCell, s.procedureStr)
		s.procedureCell = ""
		s.procedureStr = ""
		if len(s.confirmStr) != 0 {
			s.confirmStr = fmt.Sprintf("%s\n", s.confirmStr)
		}
		s.confirmStr += line
		if len(s.confirmCell) == 0 {
			s.confirmCell = fmt.Sprintf("%s%d", string(startHeaderColumn+rune(column)), startHeaderRow+1+s.testItemNum)
		}
		return
	}
	//確認内容追加
	if s.confirmCell != "" && s.confirmStr != "" {
		excel.SetCellValue(sheetName, s.confirmCell, s.confirmStr)
		s.confirmCell = ""
		s.confirmStr = ""
		s.testItemNum++
	}

	cell := fmt.Sprintf("%s%d", string(startHeaderColumn+rune(column)), startHeaderRow+1+s.testItemNum)
	var outline []string
	if column == 4 {
		outline = strings.Split(line, "->")
		excel.SetCellValue(sheetName, cell, outline[0])
	} else {
		excel.SetCellValue(sheetName, cell, line)
	}

	switch column {
	case 1:
		if s.testItemNum != 0 {
			s.bigItemBorderRow = append(s.bigItemBorderRow, s.testItemNum)
		}
	case 2:
		if s.testItemNum != 0 {
			s.middleItemBorderRow = append(s.middleItemBorderRow, s.testItemNum)
		}
	case 3:
		if s.testItemNum != 0 {
			s.smallItemBorderRow = append(s.smallItemBorderRow, s.testItemNum)
		}
	case 4:
		if len(outline) < 2 {
			log.Printf("SetCellData (the format is incorrect):[%s]", outline[0])
			return
		}
		result := regexp.MustCompile("[:@]").Split(outline[1], -1)
		for i, v := range result {
			cell := fmt.Sprintf("%s%d", string(startHeaderColumn+rune(column+i)+3), startHeaderRow+1+s.testItemNum)
			excel.SetCellValue(sheetName, cell, v)
		}
	default:
	}
}

func SetConstContent(excel *excelize.File, s *SheetData) {
	//デフォルトで作成されるSheet1を削除
	index := excel.NewSheet(sheetName)
	excel.SetActiveSheet(index)
	excel.DeleteSheet("Sheet1")
	if 0 < len(s.constContent[0]) {
		for _, v := range constContent {
			err := excel.SetCellValue(s.sheetName, v[0], v[1])
			if err != nil {
				log.Printf("SetConstContent : %s", err)
			}
		}
	}
}
