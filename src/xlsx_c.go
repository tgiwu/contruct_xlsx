package main

import (
	"fmt"
	"path/filepath"

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

	err := writeTitle(sheet, getTitle(sheetName, mConf.Month, mConf.Year))

	if err != nil {
		return err
	}

	err = writeHeader(sheet, sheetName, mConf.Headers)

	if err != nil {
		return err
	}
	return nil
}

func writeRow(sheet *xlsx.Sheet, title string, salaries []Salary) error {
	if len(salaries) == 0 {
		return EmptyError{msg: "salary is empty " + title}
	}

	// for i, salary := range salaries {
	// 	row, err := sheet.Row(salary.Id + 2)
	// 	if err != nil {
	// 		return err
	// 	}

	// cell := row.GetCell(i)
	// switch i {

	// }
	// }

	return nil
}

func writeHeader(sheet *xlsx.Sheet, title string, headers []string) error {

	if len(headers) == 0 {
		return EmptyError{msg: fmt.Sprintf("empty header %s \n", title)}
	}
	row, err := sheet.Row(1)

	if err != nil {
		return err
	}

	for i, header := range headers {
		cell := row.GetCell(i)
		cell.SetValue(header)
		cell.SetStyle(getStyle(TYPE_ROW_HEADER))
	}

	return nil
}

func writeTitle(sheet *xlsx.Sheet, title string) error {
	titleCell, err := sheet.Cell(0, 0)

	if err != nil {
		fmt.Println("write title error ", title, err.Error())
		return err
	}

	titleCell.Merge(11, 0)
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
