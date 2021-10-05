package main

import (
	"fmt"
	"github.com/spf13/pflag"
	"github.com/xuri/excelize/v2"
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

var source = pflag.StringP("source", "s", "", "Input Source File")
var sourceCol = pflag.StringP("sourceCol", "c", "", "Input Source File Col")
var fileCourseMap = make(map[string]*excelize.File, 0)
var sourceRowCourseMap = make([]map[string]string, 0)
var fileCourseRowInt = make(map[string]int64, 0)
var title = make([]string, 0)

func main() {
	pflag.Parse()

	if *source == "" || *sourceCol == "" {
		fmt.Println("请使用 [go run main.go -s (文件名) -c (学科所在列)] 的方式指定源文件")
		return
	}
	col := strings.ToUpper(*sourceCol)

	processNewCourse(*source, col)
	process(*source, col)
}

func processNewCourse(sourceName string, col string) {
	//按行读取旧excel
	f, err := excelize.OpenFile(sourceName)
	rows, err := f.Rows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	rowInt := 1
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			fmt.Println(err)
			return
		}
		if rowInt == 1 {
			for _, content := range row {
				title = append(title, content)
			}
			rowInt++
			sourceRowCourseMap = append(sourceRowCourseMap, map[string]string{})
			continue
		}

		courseMap := make(map[string]string, 0)
		for idx, colCell := range row {
			if int2character[idx+1] == col {
				tmpString := colCell
				for {
					ret1 := strings.SplitN(tmpString, "(", 2)
					if len(ret1) == 1 {
						break
					}
					ret2 := strings.SplitN(ret1[1], ")", 2)
					if len(ret2) == 1 {
						break
					}
					tmpString = ret2[1]
					course := ret2[0]
					//fmt.Println(course)
					if _, ok := fileCourseMap[course]; !ok {
						fileCourseMap[course] = makeNewFile()
					}
					courseMap[course] = course
				}
				break
			}
		}
		sourceRowCourseMap = append(sourceRowCourseMap, courseMap)
	}
}

func process(sourceName string, col string) {
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
			for _, f := range fileCourseMap {
				for idx, content := range title {
					f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", int2character[idx+1], 1), content)
				}
			}
			rowInt++
			continue
		}

		cellInt = 1
		for _, colCell := range row {
			courseMap := sourceRowCourseMap[rowInt-1]
			for _, v := range courseMap {
				rowNum := fileCourseRowInt[v]

				if f, ok := fileCourseMap[v]; ok {
					f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", int2character[cellInt], rowNum+2), colCell)
				}
			}
			cellInt++
		}

		courseMap := sourceRowCourseMap[rowInt-1]
		for _, v := range courseMap {
			fileCourseRowInt[v]++
		}
		rowInt++
	}

	//// 设置工作簿的默认工作表
	//fnew.SetActiveSheet(index)
	// 根据指定路径保存文件
	for course, f := range fileCourseMap {
		fmt.Println("正在保存", course, "学科")
		fileString := strings.Replace(course, "/", "-", -1)
		err = f.SaveAs(fileString + "_" + sourceName)
	}

	if err != nil {
		fmt.Println(err)
	}
}

func makeNewFile() *excelize.File {
	fnew := excelize.NewFile()
	_ = fnew.NewSheet("Sheet1")
	return fnew
}
