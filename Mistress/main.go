package main

import (
	"fmt"
	"github.com/spf13/pflag"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"os"
	"strings"
)

var character2int = map[string]int{
	"A": 1,
	"B": 2,
	"C": 3,
	"D": 4,
	"E": 5,
	"F": 6,
	"G": 7,
	"H": 8,
	"I": 9,
	"J": 10,
	"K": 11,
	"L": 12,
	"M": 13,
	"N": 14,
	"O": 15,
}
var int2character = map[int]string{
	1:  "A",
	2:  "B",
	3:  "C",
	4:  "D",
	5:  "E",
	6:  "F",
	7:  "G",
	8:  "H",
	9:  "I",
	10: "J",
	11: "K",
	12: "L",
	13: "M",
	14: "N",
	15: "O",
}

var filterAllData []string
var filterSingleData []string
var AllExcelSource = make([]string, 0)

var singleSource = pflag.StringP("source", "s", "", "Input Source File")
var singleSourceCol = pflag.StringP("col", "c", "H", "Input Source File Col")

var isSingleFileMode = false

func main() {
	processParam()

	fmt.Println("excel操作工具 V2 终端命令版。")
	if isSingleFileMode {
		fmt.Printf("当前为单文件处理模式,只处理 [%s] 文件的 (%s)列\n", *singleSource, *singleSourceCol)
		if !checkFileExists(*singleSource) {
			fmt.Println("文件不存在!")
			os.Exit(1)
		}
	} else {
		fmt.Println("当前为批处理模式,当前目录有如下文件")
		getAllExcel()
		tipsAndShowSource()
	}
	//os.Exit(1)

	// 读取关键词
	getFilterWords()
	//fmt.Println(filterAllData)
	//fmt.Println(filterSingleData)
	//os.Exit(1)

	// 开始校验
	if isSingleFileMode {
		process(*singleSource, filterAllData)
		for _, word := range filterSingleData {
			process(*singleSource, []string{word})
		}
	} else {
		for _, name := range AllExcelSource {
			process(name, filterAllData)
			for _, word := range filterSingleData {
				process(name, []string{word})
			}
		}
	}

	fmt.Println("处理完毕")
}

func processParam() {
	pflag.Parse()
	if *singleSource != "" {
		isSingleFileMode = true
	}

}

func getAllExcel() {
	reader, err := ioutil.ReadDir("./")
	if err != nil {
		fmt.Println("程序内部错误 错误代码 [01]")
		os.Exit(1)
	}
	for _, fi := range reader {
		if fi.IsDir() {
			continue
		}
		fileName := fi.Name()
		if fileName == "filter.xlsx" {
			continue
		}

		if fileName[len(fi.Name())-5:] == ".xlsx" {
			AllExcelSource = append(AllExcelSource, fileName)
		}
	}
	if len(AllExcelSource) == 0 {
		fmt.Println("当前目录下未找到(.xlsx)结尾文件")
		os.Exit(1)
	}
}

func tipsAndShowSource() {
	for i, v := range AllExcelSource {
		fmt.Printf("序号:[%d] ==> %s\n", i, v)
	}
	//fmt.Println("请输入需要处理的源excel文件序号: ")
	//fmt.Scanln(&sourceFile)
	//fmt.Println("请输入源excel文件需要匹配的列,大小写均可 例如(C)")
	//fmt.Scanln(&sourceContent)
	//fmt.Println("请输入包含关键词的excel文件序号: ")
	//fmt.Scanln(&filterFile)
	//fmt.Println("请输入关键词excel的关键词所在列,大小写均可 例如(H)")
	//fmt.Scanln(&filterContent)
	//
	//sourceContent = strings.ToUpper(sourceContent)
	//filterContent = strings.ToUpper(filterContent)
}

func checkFileExists(sourceName string) bool {
	_, err := os.Stat(sourceName)
	if err == nil {
		return true
	}
	return false
}

func getFilterWords() {
	ff, err := excelize.OpenFile("filter.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	allCols, err := ff.GetCols("Sheet1")
	tmpInt := 1
	for _, col := range allCols {
		if tmpInt != character2int["B"] {
			tmpInt++
			continue
		}

		for _, rowCell := range col {
			//fmt.Println(rowCell)
			filterAllData = append(filterAllData, rowCell)
		}
		//fmt.Println(col)
		tmpInt++
	}

	if len(filterAllData) == 0 {
		fmt.Println("没有全部匹配的过滤关键词！")
		os.Exit(1)
	}

	singleCols, err := ff.GetCols("Sheet2")
	tmpInt = 1
	for _, col := range singleCols {
		if tmpInt != character2int["B"] {
			tmpInt++
			continue
		}

		for _, rowCell := range col {
			//fmt.Println(rowCell)
			filterSingleData = append(filterSingleData, rowCell)
		}
		//fmt.Println(col)
		tmpInt++
	}

	if len(filterSingleData) == 0 {
		fmt.Println("没有分别匹配的过滤关键词！")
		os.Exit(1)
	}
	//fmt.Println(cols)
}

func process(sourceName string, filterWords []string) {
	//SheetName
	//按行读取旧excel
	f, err := excelize.OpenFile(sourceName)
	rows, err := f.Rows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	checkCol := character2int[*singleSourceCol]

	fnew := excelize.NewFile()
	index := fnew.NewSheet("Sheet1")
	styleRed, err := fnew.NewStyle(`{"fill":{"type":"pattern","color":["#EEEE00"],"pattern":1}}`)
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
		cellInt = 1
		isHas := false
		for _, colCell := range row {
			cellName := fmt.Sprintf("%s%d", int2character[cellInt], rowInt)
			if cellInt == checkCol {
				for _, v := range filterWords {
					if strings.Contains(colCell, v) {
						isHas = true
						//fmt.Println(cellName)
						break
					}
				}
			}
			fnew.SetCellValue("Sheet1", cellName, colCell)
			cellInt++
		}
		if isHas {
			tmpInt := 1
			for _, _ = range row {
				cellName := fmt.Sprintf("%s%d", int2character[tmpInt], rowInt)
				fnew.SetCellStyle("Sheet1", cellName, cellName, styleRed)
				tmpInt++
			}
		}
		rowInt++
	}

	// 设置工作簿的默认工作表
	fnew.SetActiveSheet(index)
	// 根据指定路径保存文件
	if len(filterWords) == 1 {
		filterWord := strings.Replace(filterWords[0], "/", "-", -1)
		filterWord = strings.Replace(filterWord, " ", "-", -1)
		err = fnew.SaveAs(filterWord + "_" + sourceName)
	} else {
		err = fnew.SaveAs("allFilter_" + sourceName)
	}
	if err != nil {
		fmt.Println(err)
	}
}
