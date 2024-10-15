package main

import (
	"fmt"
	"path/filepath"
	"reflect"

	// "sync"

	"github.com/360EntSecGroup-Skylar/excelize"
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
var styleM map[int]int

type EmptyError struct {
	msg string
}

func (ee EmptyError) Error() string {
	return ee.msg
}

func constructXlsx(salaryMap map[string]map[string]Salary) error {
	excel := excelize.NewFile()

	for name, salaries := range salaryMap {
		constructSalarySheet(excel, name, salaries)
	}

	excel.DeleteSheet("Sheet1")

	delFileIfExist(mConf.OutputPath, mConf.FileName)
	excel.SaveAs(filepath.Join(mConf.OutputPath, mConf.FileName))
	return nil
}

func constructSalarySheet(excel *excelize.File, sheetName string, salary map[string]Salary) {
	excel.NewSheet(sheetName)

	list := make([]Salary, len(salary)+1)
	total := Salary{}
	calcTotal(&salary, &list, &total)

	fillTitle(excel, sheetName, getTitle(sheetName, mConf.Month, mConf.Year))
	fillHeader(excel, sheetName, mConf.Headers)
	fillRow(excel, sheetName, sortSalaryById(salaryMap[sheetName]))
	fillTotal(excel, sheetName, len(salary)+2, total)

}

func calcTotal(salary *map[string]Salary, list *[]Salary, total *Salary) {

	standardTotal := 0
	netpayTotal := 0
	accountTotal := 0
	for _, item := range *salary {
		(*list)[item.Id-1] = item
		standardTotal += item.Standard
		netpayTotal += item.NetPay
		accountTotal += item.Account
	}

	total.Name = "合计"
	total.Standard = standardTotal
	total.NetPay = netpayTotal
	total.Account = accountTotal
}

func fillTitle(excel *excelize.File, sheetName string, title string) {
	excel.MergeCell(sheetName, pos(0, 0), pos(0, len(mConf.Headers)-1))
	excel.SetCellValue(sheetName, pos(0, 0), title)
	excel.SetCellStyle(sheetName, pos(0, 0), pos(0, len(mConf.Headers)-1), cellStyle(excel, TYPE_ROW_TITLE))
}

func fillHeader(excel *excelize.File, sheetName string, headers []string) {
	for i, header := range headers {
		excel.SetCellValue(sheetName, pos(1, i), header)

		switch {
		case header == "备注":
			excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 27.75)
		case header == "序号":
			excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 7)
		case len(header) < 4:
			excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 8.33)
		default:
			excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 11.33)
		}
	}
	excel.SetCellStyle(sheetName, pos(1, 0), pos(1, len(mConf.Headers)-1), cellStyle(excel, TYPE_ROW_HEADER))
}

func fillRow(excel *excelize.File, sheetName string, salaries []Salary) {
	for i, salary := range salaries {
		for j, s := range mConf.Headers {
			v := reflect.ValueOf(salary)
			if v.Kind() == reflect.Struct {
				value := v.FieldByName(mConf.HeadersMap[s])
				if value.Kind() == reflect.String {
					excel.SetCellStr(sheetName, pos(i+2, j), value.String())
				} else {
					excel.SetCellInt(sheetName, pos(i+2, j), int(value.Int()))
				}
			}
			if i%2 == 0 {
				excel.SetCellStyle(sheetName, pos(i+2, j), pos(i+2, j), cellStyle(excel, STYLE_TYPE_NORMAL))
			} else {
				excel.SetCellStyle(sheetName, pos(i+2, j), pos(i+2, j), cellStyle(excel, STYLE_TYPE_NORMAL_GREY))
			}
		}

	}
}

func fillTotal(excel *excelize.File, sheetName string, row int, total Salary) {
	excel.MergeCell(sheetName, pos(row, 0), pos(row, 3))
	excel.SetCellValue(sheetName, pos(row, 0), total.Name)
	excel.SetCellStyle(sheetName, pos(row, 0), pos(row, 3), cellStyle(excel, TYPE_ROW_TOTAL))

	excel.SetCellValue(sheetName, pos(row, 4), total.Standard)
	excel.SetCellValue(sheetName, pos(row, 5), total.NetPay)
	excel.SetCellStyle(sheetName, pos(row, 4), pos(row, len(mConf.Headers)-3), cellStyle(excel, TYPE_ROW_NORMAL))

	excel.SetCellValue(sheetName, pos(row, len(mConf.Headers)-2), total.Account)
	excel.SetCellStyle(sheetName, pos(row, len(mConf.Headers)-2), pos(row, len(mConf.Headers)-2), cellStyle(excel, TYPE_ROW_TOTAL))

	excel.SetCellStyle(sheetName, pos(row, len(mConf.Headers)-1), pos(row, len(mConf.Headers)-1), cellStyle(excel, TYPE_ROW_TOTAL))
}

func cellStyle(excel *excelize.File, styleNo int) int {

	if styleM == nil {
		styleM = make(map[int]int)
	}

	if styleId, found := styleM[styleNo]; found {
		return styleId
	}
	styleStr := ""
	switch styleNo {
	case STYLE_TYPE_TITLE:
		styleStr = `{
		"border":[{"type":"left","color":"000000","style":1},
			{"type":"top","color":"000000","style":1},
			{"type":"right","color":"000000","style":1},
			{"type":"bottom","color":"000000","style":1}],
		"fill":{"type":"gradient","color":["#A5A5A5","#A5A5A5"], "shading":1},
		"alignment":{"horizontal":"center", "vertical":"center"},
		"font":{"bold":true, "italic":false, "family":"Microsoft YaHei", "size":16, "color":"#FFFFFF"}}`
	case STYLE_TYPE_HEADER:
		styleStr = `{
			"border":[{"type":"left","color":"000000","style":1},
				{"type":"top","color":"000000","style":1},
				{"type":"right","color":"000000","style":1},
				{"type":"bottom","color":"000000","style":1}],
			"fill":{"type":"gradient","color":["#A5A5A5","#A5A5A5"], "shading":1},
			"alignment":{"horizontal":"center", "vertical":"center"},
			"font":{"bold":true, "italic":false, "family":"Microsoft YaHei", "size":14, "color":"#FFFFFF"}}`
	case STYLE_TYPE_NORMAL:
		styleStr = `{
			"border":[{"type":"left","color":"000000","style":1},
				{"type":"top","color":"000000","style":1},
				{"type":"right","color":"000000","style":1},
				{"type":"bottom","color":"000000","style":1}],
			"fill":{"type":"gradient","color":["#FFFFFF","#FFFFFF"], "shading":1},
			"alignment":{"horizontal":"center", "vertical":"center"},
			"font":{"bold":true, "italic":false, "family":"宋体", "size":12, "color":"#000000"}}`
	case STYLE_TYPE_NORMAL_GREY:
		styleStr = `{
			"border":[{"type":"left","color":"000000","style":1},
				{"type":"top","color":"000000","style":1},
				{"type":"right","color":"000000","style":1},
				{"type":"bottom","color":"000000","style":1}],
			"fill":{"type":"gradient","color":["#E7E6E6","#E7E6E6"], "shading":1},
			"alignment":{"horizontal":"center", "vertical":"center"},
			"font":{"bold":true, "italic":false, "family":"宋体", "size":12, "color":"#000000"}}`
	case STYLE_TYPE_TOTAL:
		styleStr = `{
			"border":[{"type":"left","color":"000000","style":1},
				{"type":"top","color":"000000","style":1},
				{"type":"right","color":"000000","style":1},
				{"type":"bottom","color":"000000","style":1}],
			"fill":{"type":"gradient","color":["#FFFFFF","#FFFFFF"], "shading":1},
			"alignment":{"horizontal":"center", "vertical":"center"},
			"font":{"bold":true, "italic":false, "family":"宋体", "size":14, "color":"#000000"}}`
	}
	styleId, err := excel.NewStyle(styleStr)
	if err != nil {
		panic(err)
	}
	return styleId
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
