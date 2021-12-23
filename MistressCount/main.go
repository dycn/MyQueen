package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"os"
	"strings"
)


var AllExcelSource = make([]string, 0)
var retFile *excelize.File
var (
	retCountryRow int64 = 2
	retArtcleRow int64 = 2
	retYesNoRow int64 = 2
)

func main() {
	getAllExcel()
	initRetExcel()
	for _,name := range AllExcelSource {
		process(name)
	}
	closeRetExcel()
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

func initRetExcel(){
	fnew := excelize.NewFile()
	index := fnew.NewSheet("国家")

	// 根据指定路径保存文件
	fnew.SetCellValue("国家", "A1", "学科")
	fnew.SetCellValue("国家", "B1", "国家")
	fnew.SetCellValue("国家", "C1", "数量")

	fnew.NewSheet("文章类型")
	fnew.SetCellValue("文章类型", "A1", "学科")
	fnew.SetCellValue("文章类型", "B1", "文章类型")
	fnew.SetCellValue("文章类型", "C1", "数量")

	fnew.NewSheet("国家yes/no")
	fnew.SetCellValue("国家yes/no", "A1", "学科")
	fnew.SetCellValue("国家yes/no", "B1", "yes/no")
	fnew.SetCellValue("国家yes/no", "C1", "数量")

	fnew.SetActiveSheet(index) //默认页

	//居中样式
	style, err := fnew.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Vertical:"center"}})
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
	err = fnew.SetColStyle("国家", "A", style)
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
	err = fnew.SetColStyle("文章类型", "A", style)
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
	err = fnew.SetColStyle("国家yes/no", "A", style)
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}

	retFile = fnew
}

func closeRetExcel(){
	err := retFile.SaveAs("统计结果.xlsx")

	if err != nil {
		fmt.Println(err)
	}
}

func process(sourceName string) {
	subject := sourceName[:3]
	//fmt.Println(subject)
	//os.Exit(1)

	//按行读取旧excel
	f, err := excelize.OpenFile(sourceName)
	rows, err := f.Rows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	rowInt := 1
	cellInt := 1
	var countryMap = make(map[string]int64, 0)
	var artcleMap = make(map[string]int64, 0)
	var yesnoMap = make(map[string]int64, 0)

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
			if cellInt == 12 {
				artcleMap[colCell]++
			}
			if cellInt == 13 {
				tmp := strings.Split(colCell, " ")
				country := ""
				yesno := ""
				if len(tmp) == 1 {
					country = colCell
				}else{
					yesno = tmp[len(tmp)-1]
					country = strings.Join(tmp[:len(tmp)-1], " ")
				}

				countryMap[country]++
				yesnoMap[yesno]++
			}
			cellInt++
		}
		rowInt++
	}


	for country, count := range countryMap {
		retFile.SetCellValue("国家", fmt.Sprintf("A%d", retCountryRow), subject)
		retFile.SetCellValue("国家", fmt.Sprintf("B%d", retCountryRow), country)
		retFile.SetCellValue("国家", fmt.Sprintf("C%d", retCountryRow), count)
		retCountryRow++
	}

	for artcle, count := range artcleMap {
		retFile.SetCellValue("文章类型", fmt.Sprintf("A%d", retArtcleRow), subject)
		retFile.SetCellValue("文章类型", fmt.Sprintf("B%d", retArtcleRow), artcle)
		retFile.SetCellValue("文章类型", fmt.Sprintf("C%d", retArtcleRow), count)
		retArtcleRow++
	}

	for yesno, count := range yesnoMap {
		retFile.SetCellValue("国家yes/no", fmt.Sprintf("A%d", retYesNoRow), subject)
		retFile.SetCellValue("国家yes/no", fmt.Sprintf("B%d", retYesNoRow), yesno)
		retFile.SetCellValue("国家yes/no", fmt.Sprintf("C%d", retYesNoRow), count)
		retYesNoRow++
	}

	err1 := retFile.MergeCell("国家", fmt.Sprintf("A%d",int(retCountryRow) - len(countryMap)), fmt.Sprintf("A%d",retCountryRow-1))
	if err1 != nil{
		fmt.Println(int(retCountryRow) - len(countryMap), len(countryMap),err1)
		os.Exit(1)
	}
	err2 := retFile.MergeCell("文章类型", fmt.Sprintf("A%d",int(retArtcleRow) - len(artcleMap)), fmt.Sprintf("A%d",retArtcleRow-1))
	if err2 != nil {
		fmt.Println(err2)
		os.Exit(1)
	}
	err3 := retFile.MergeCell("国家yes/no", fmt.Sprintf("A%d",int(retYesNoRow) - len(yesnoMap)), fmt.Sprintf("A%d",retYesNoRow-1))
	if err3 != nil{
		fmt.Println(err3)
	}
}
