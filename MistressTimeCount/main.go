package main

import (
	"fmt"
	"github.com/spf13/pflag"
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
)

var source = pflag.StringP("source", "s", "", "Input Source File")
var pubMap = make(map[string]int64, 0)
var returnMap = make(map[string]int64, 0)

func main() {
	pflag.Parse()

	if *source == "" {
		fmt.Println("请使用 [go run main.go -s (文件名) -c (学科所在列)] 的方式指定源文件")
		return
	}
	process(*source)
}

func process(sourceName string) {
	//按行读取旧excel
	f, err := excelize.OpenFile(sourceName)
	rows, err := f.Rows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	rowInt := 1
	cellInt := 1

	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			fmt.Println(err)
			return
		}

		if rowInt == 1 {
			rowInt++
			continue
		}

		cellInt = 1
		for _, colCell := range row {
			if cellInt == 9 {
				str := strings.Split(colCell, "/")
				yearStr := str[len(str)-1]
				pubMap[yearStr]++
			}
			if cellInt == 10 {
				str := strings.Split(colCell, "/")
				yearStr := str[len(str)-1]
				returnMap[yearStr]++
			}
			cellInt++
		}
		rowInt++
	}

	//fmt.Println(pubMap, returnMap)
	//os.Exit(1)

	fnew := excelize.NewFile()
	index := fnew.NewSheet("Sheet1")

	// 根据指定路径保存文件
	fnew.SetCellValue("Sheet1", "A1", "年份")
	fnew.SetCellValue("Sheet1", "B1", "发表时间")
	fnew.SetCellValue("Sheet1", "C1", "撤稿时间")
	i := 2
	yearNum := 1950
	for yearNum < 2050 {
		tmpStr := strconv.Itoa(yearNum)
		numPub, ok1 := pubMap[tmpStr]
		numReturn, ok2 := returnMap[tmpStr]
		if ok1 || ok2 {
			fmt.Println(ok1, ok2, numPub, numReturn)
			fnew.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), yearNum)
			fnew.SetCellValue("Sheet1", fmt.Sprintf("B%d", i), numPub)
			fnew.SetCellValue("Sheet1", fmt.Sprintf("C%d", i), numReturn)
			i++
		}
		yearNum++
	}

	// 设置工作簿的默认工作表
	fnew.SetActiveSheet(index)
	err = fnew.SaveAs("年份统计_" + sourceName)

	if err != nil {
		fmt.Println(err)
	}
}
