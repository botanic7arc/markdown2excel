//go get -u github.com/xuri/excelize/v2
package main

import (
	"log"
	"os"
	"time"

	"github.com/botanic7arc/markdown2excel/pkg"

	"github.com/xuri/excelize/v2"
)

func main() {
	inputFileName, err := pkg.GetInputFileName()
	if err != nil {
		log.Printf("getInputFileName Error : %s", err)
		duration := time.Duration(15) * time.Second
		time.Sleep(duration)
		os.Exit(1)
	}
	f := excelize.NewFile()
	s := pkg.NewSheetData()
	pkg.SetConstContent(f, s) //mdファイルの内容に依存しないデータ追加等処理
	pkg.SetRowWidth(f)
	pkg.SetTableHeader(f)
	pkg.ReadFile(f, inputFileName, s) //セルへの値設定も行う
	pkg.SetCellStyle(f, s)
	pkg.OutputFile(f, inputFileName)

	duration := time.Duration(4) * time.Second
	time.Sleep(duration)
}
