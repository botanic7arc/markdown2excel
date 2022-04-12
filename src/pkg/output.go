package pkg

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

func OutputFile(excel *excelize.File, fileName string) {
	outputFileName := strings.TrimRight(fileName, ".md") + ".xlsx"
	fmt.Printf("%s\n", dialog[lang][2])
	fmt.Printf("%s%s\n", dialog[lang][3], outputFileName)
	if err := excel.SaveAs(outputFileName); err != nil {
		fmt.Println(err)
	}
}
