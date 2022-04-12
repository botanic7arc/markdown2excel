package pkg

import (
	"bufio"
	"fmt"
	"io/ioutil"
	"os"
	"strconv"
	"strings"
)

func GetInputFileName() (string, error) {
	fmt.Println(dialog[lang][0])
	//ディレクトリ内ファイル列挙
	files, _ := ioutil.ReadDir("./")
	j := 1
	mdFileList := []string{}
	for _, f := range files {
		if strings.HasSuffix(f.Name(), ".md") {
			fmt.Printf("%d:%s\n", j, f.Name())
			j++
			mdFileList = append(mdFileList, f.Name())
		}
	}
	if len(mdFileList) == 0 {
		return "", fmt.Errorf("md file does not exist in the same directory\nPlace the executable file in the directory where the md file exists")
	}

	//入力受け取り
	fmt.Print(dialog[lang][1])
	scanner := bufio.NewScanner(os.Stdin)
	scanner.Scan()
	in := scanner.Text()

	var i int
	i, err := strconv.Atoi(in)
	if err != nil {
		return "", fmt.Errorf("GetInputFileName(Input is not a number):%s", in)
	}
	if len(mdFileList) < i || i < 1 {
		return "", fmt.Errorf("GetInputFileName(Out of range value):%d", i)
	}
	return mdFileList[i-1], nil
}
