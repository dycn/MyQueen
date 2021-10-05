package main

import (
	"github.com/xuri/excelize/v2"
	"sync"
)

type Excel struct {
	lock     sync.Locker
	f        *excelize.File
	curSheet string
}

func NewExcel() *Excel {
	f := excelize.NewFile()
	excel := &Excel{f: f}
	return excel
}

func OpenExcel(fileName, sheetName string) *Excel {
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil
	}
	excel := &Excel{f: f, curSheet: sheetName}
	return excel
}

//func (e *Excel) Close() {
//	close(e.f)
//}

// 获取某个单元格的内容
func (e *Excel) getCell(cellName string) string {
	e.lock.Lock()
	ret, err := e.f.GetCellValue(e.curSheet, cellName)
	e.lock.Unlock()
	if err != nil {
		return ""
	}
	return ret
}

//

func getCell() {

}
