package main

import (
	"fmt"
	"path/filepath"
	"sync"

	"github.com/tealeg/xlsx/v3"
)

const STYLE_TYPE_TITLE = 0
const STYLE_TYPE_HEADER = 1
const STYLE_TYPE_NORMAL = 2
const STYLE_TYPE_NORMAL_GREY = 3
const STYLE_TYPE_TOTAL = 4

const TYPE_ROW_TITLE = 0
const TYPE_ROW_HEADER = 1
const TYPE_ROW_NORMAL = 2
const TYPE_ROW_NORMAL_GREY = 3
const TYPE_ROW_TOTAL = 4

func createSalaryXlsx(salaryMap map[string]map[string]Salary) error {
	wb := xlsx.NewFile()

	defer wb.Save(filepath.Join(mConf.OutputPath, mConf.FileName))

	var fcwg sync.WaitGroup

	fcwg.Add(len(salaryMap))

	for i := 0; i < len(salaryMap); i++ {
		go createSalarySheet()
	}

	fcwg.Wait()

	return nil
}

func createSalarySheet(wb *xlsx.File, sheetName string, salary map[string]Salary) error {
	sheet, found := checkSheet(wb, sheetName)

	return nil
}

func checkOrCreateSheet(wb *xlsx.File, name string) (sheet *xlsx.Sheet, ok bool) {
	sheet, ok = wb.Sheet[name]
	if !ok {
		fmt.Println("Sheet ", name, " not exist")
		sheet, err := wb.AddSheet(name)
		if err != nil {
			panic(err)
		}
		return 
	}
	fmt.Println("Max row in sheet: ", sheet.MaxRow)
	return
}

func createSheets(name string, wb *xlsx.File) (sheet *xlsx.Sheet, err error) {
	sheet, err = wb.AddSheet(name)
	fmt.Println(err)
	return
}

func copySheet(oldName string, newName string, wb *xlsx.File) (sheet *xlsx.Sheet, err error) {
	old, ok := wb.Sheet[oldName]
	if !ok {
		fmt.Println("Sheet ", oldName, " not exist")
		return
	}
	sheet, err = wb.AppendSheet(*old, newName)

	return
}

func getStyle(styleType int) (style *xlsx.Style) {
	style = xlsx.NewStyle()

	style.Alignment.Horizontal = "center"
	style.Alignment.Vertical = "center"
	style.Fill.FgColor = "0000000"
	style.Border = *xlsx.NewBorder("thin", "thin", "thin", "thin")

	switch styleType {
	case STYLE_TYPE_TITLE:
		style.Font.Size = 20
		style.Font.Bold = true
		style.ApplyAlignment = true
		style.Fill.BgColor = "5A5A5A00"
		style.Font.Name = "Microsoft YaHei"
	case STYLE_TYPE_HEADER:
		style.Fill.BgColor = "5A5A5A00"
		style.Font.Bold = true
		style.Font.Name = "Microsoft YaHei"
	case STYLE_TYPE_NORMAL:
		style.Fill.BgColor = "00000000"
		style.Font.Name = "宋体"
		style.Font.Size = 14
	case STYLE_TYPE_NORMAL_GREY:
		style.Fill.BgColor = "E7E6E600"
		style.Font.Name = "宋体"
		style.Font.Size = 14
	case STYLE_TYPE_TOTAL:
		style.Font.Name = "宋体"
		style.Font.Bold = true
	}

	return
}

// func addRowTyped(rowType int, sheet *xlsx.Sheet) (row *xlsx.Row) {

// }
