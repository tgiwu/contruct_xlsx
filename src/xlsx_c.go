package main

import (
	"fmt"
	"path/filepath"
	"reflect"

	// "sync"

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

var styleMap map[int]*xlsx.Style

type EmptyError struct {
	msg string
}

func (ee EmptyError) Error() string {
	return ee.msg
}

func createSalaryXlsx(salaryMap map[string]map[string]Salary) error {
	wb := xlsx.NewFile()

	// var fcwg sync.WaitGroup

	// fcwg.Add(len(salaryMap))

	for name, salaries := range salaryMap {
		createSalarySheet(wb, name, salaries)
	}

	// fcwg.Wait()

	delFileIfExist(mConf.OutputPath, mConf.FileName)

	err := wb.Save(filepath.Join(mConf.OutputPath, mConf.FileName))

	if err != nil {
		fmt.Println(err)
		return err
	}
	return nil
}

func createSalarySheet(wb *xlsx.File, sheetName string, salary map[string]Salary) error {
	sheet, _ := checkOrCreateSheet(wb, sheetName)

	list := make([]Salary, len(salary)+1)

	standardTotal := 0
	netpayTotal := 0
	accountTotal := 0
	for _, item := range salary {
		list[item.Id-1] = item
		standardTotal += item.Standard
		netpayTotal += item.NetPay
		accountTotal += item.Account
	}

	totalSalary := Salary{Name: "合计",
		Standard: standardTotal,
		NetPay:   netpayTotal,
		Account:  accountTotal}

	list[len(salary)] = totalSalary

	constructTable(sheet, len(mConf.Headers), len(salary)+3)

	err := writeTitle(sheet, getTitle(sheetName, mConf.Month, mConf.Year))

	if err != nil {
		return err
	}

	err = writeHeader(sheet, sheetName, mConf.Headers)

	if err != nil {
		return err
	}

	err = writeRow(sheet, sheetName, sortSalaryById(salaryMap[sheetName]))

	if err != nil {
		return nil
	}

	return nil
}

func constructTable(sheet *xlsx.Sheet, colCount int, rowCount int) {
	for i := 1; i < colCount+1; i++ {
		col := xlsx.NewColForRange(i, i)
		if i == colCount {
			col.SetWidth(17)
		} else {
			col.SetWidth(9.33)
		}
		sheet.SetColParameters(col)
	}

	for i := 0; i < rowCount+3; i++ {
		row := sheet.AddRow()
		if i == 0 {
			row.SetHeight(23)
		} else {
			row.SetHeight(18)
		}
	}
}

func writeRow(sheet *xlsx.Sheet, title string, salaries []Salary) error {
	if len(salaries) == 0 {
		return EmptyError{msg: "salary is empty " + title}
	}

	for i, salary := range salaries {
		row, _ := sheet.Row(i + 2)
		for j, s := range mConf.Headers {
			cell := row.GetCell(j)
			v := reflect.ValueOf(salary)
			if v.Kind() == reflect.Struct {
				value := v.FieldByName(mConf.HeadersMap[s])
				if value.Kind() == reflect.String {
					cell.SetValue(fmt.Sprint(value))
				} else {
					cell.SetInt64(value.Int())
				}
			}
			if salary.Id%2 == 0 {
				cell.SetStyle(getStyle(STYLE_TYPE_NORMAL_GREY))
			} else {
				cell.SetStyle(getStyle(STYLE_TYPE_NORMAL))
			}
		}

	}

	return nil
}

func writeHeader(sheet *xlsx.Sheet, title string, headers []string) error {

	if len(headers) == 0 {
		return EmptyError{msg: fmt.Sprintf("empty header %s \n", title)}
	}
	row, _ := sheet.Row(1)
	for i, header := range headers {

		cell := row.GetCell(i)
		cell.SetValue(header)
		cell.SetStyle(getStyle(TYPE_ROW_HEADER))
	}

	return nil
}

func writeTitle(sheet *xlsx.Sheet, title string) error {

	titleCell, _ := sheet.Cell(0, 0)

	titleCell.Merge(len(mConf.Headers)-1, 0)
	titleCell.SetValue(title)
	titleCell.SetStyle(getStyle(STYLE_TYPE_TITLE))

	return nil
}

func getTitle(sheetName string, month int, year int) string {
	return fmt.Sprintf("%s%d年%d月工资", sheetName, year, month)
}

func checkOrCreateSheet(wb *xlsx.File, name string) (sheet *xlsx.Sheet, ok bool) {
	sheet, ok = wb.Sheet[name]
	if !ok {
		fmt.Println("Sheet ", name, " not exist")
		sheet, err := wb.AddSheet(name)
		if err != nil {
			panic(err)
		}
		return sheet, true
	}
	fmt.Println("Max row in sheet: ", sheet.MaxRow)
	return
}

// func createSheets(name string, wb *xlsx.File) (sheet *xlsx.Sheet, err error) {
// 	sheet, err = wb.AddSheet(name)
// 	fmt.Println(err)
// 	return
// }

// func copySheet(oldName string, newName string, wb *xlsx.File) (sheet *xlsx.Sheet, err error) {
// 	old, ok := wb.Sheet[oldName]
// 	if !ok {
// 		fmt.Println("Sheet ", oldName, " not exist")
// 		return
// 	}
// 	sheet, err = wb.AppendSheet(*old, newName)

// 	return
// }

func getStyle(styleType int) (style *xlsx.Style) {

	if styleMap == nil {
		styleMap = make(map[int]*xlsx.Style)
	}

	if style, found := styleMap[styleType]; found {
		return style
	}

	style = xlsx.NewStyle()

	style.Alignment.Horizontal = "center"
	style.Alignment.Vertical = "center"
	style.Border = *xlsx.NewBorder("thin", "thin", "thin", "thin")

	switch styleType {
	case STYLE_TYPE_TITLE:
		style.Font.Size = 16
		style.Font.Bold = true
		style.Fill.BgColor = "ff5a5a5a"
		style.Font.Color = xlsx.RGB_White
		style.Font.Name = "Microsoft YaHei"
	case STYLE_TYPE_HEADER:
		style.Fill.BgColor = "ff5a5a5a"
		style.Fill.FgColor = "ff5a5a5a"
		style.Font.Color = xlsx.RGB_White
		style.Font.Size = 11
		style.Font.Bold = true
		style.Font.Name = "Microsoft YaHei"
	case STYLE_TYPE_NORMAL:
		style.Fill.BgColor = "FFFFFF00"
		style.Font.Name = "宋体"
		style.Font.Size = 14
		style.Font.Color = xlsx.RGB_Black
	case STYLE_TYPE_NORMAL_GREY:
		style.Fill.BgColor = "E7E6E600"
		style.Font.Name = "宋体"
		style.Font.Size = 14
		style.Font.Color = xlsx.RGB_Black
	case STYLE_TYPE_TOTAL:
		style.Font.Name = "宋体"
		style.Font.Bold = true
		style.Font.Color = xlsx.RGB_Black
	}

	style.ApplyFill = true
	style.ApplyFont = true
	style.ApplyAlignment = true
	style.ApplyBorder = true

	styleMap[styleType] = style

	return
}
