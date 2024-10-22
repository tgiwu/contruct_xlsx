package main

import (
	"bytes"
	"fmt"
	"path/filepath"
	"reflect"
	"slices"
	"strings"
	"text/template"

	// "sync"

	"github.com/360EntSecGroup-Skylar/excelize"
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

var styleM map[int]int

var overviewArr []Overview

type EmptyError struct {
	msg string
}

func (ee EmptyError) Error() string {
	return ee.msg
}

func constructXlsx(salaryMap map[string]map[string]Salary) error {
	excel := excelize.NewFile()

	keys := make([]string, 0, len(salaryMap))
	for k := range salaryMap {
		keys = append(keys, k)
	}

	slices.Sort(keys)

	for _, key := range keys {
		constructSalarySheet(excel, key, salaryMap[key])
	}

	constructOverviewSheet(excel, overviewArr)

	excel.DeleteSheet("Sheet1")

	delFileIfExist(mConf.OutputPath, mConf.FileName)
	excel.SaveAs(filepath.Join(mConf.OutputPath, mConf.FileName))
	return nil
}

func constructOverviewSheet(excel *excelize.File, overviews []Overview) {

	slices.SortFunc(overviews, func(a, b Overview) int {
		return strings.Compare(a.Area, b.Area)
	})

	excel.NewSheet("总览")

	excel.MergeCell("总览", pos(0, 0), pos(0, len(mConf.OverviewHeader)-1))
	excel.SetCellValue("总览", pos(0, 0), fmt.Sprintf("%d年%d月工资总览", mConf.Year, mConf.Month))
	excel.SetCellStyle("总览", pos(0, 0), pos(0, len(mConf.OverviewHeader)-1), cellStyle(excel, TYPE_ROW_TITLE))

	fillHeader(excel, "总览", mConf.OverviewHeader)

	numOfStaffTotal := 0
	account := 0
	for i, overview := range overviews {
		for j, s := range mConf.OverviewHeader {

			if s == "序号" {
				excel.SetCellInt("总览", pos(i+2, j), i+1)
			} else {
				v := reflect.ValueOf(overview)
				if v.Kind() == reflect.Struct {
					value := v.FieldByName(mConf.OverviewHeaderMap[s])

					kind := value.Kind()
					switch kind {
					case reflect.String:
						excel.SetCellStr("总览", pos(i+2, j), value.String())
					case reflect.Int, reflect.Int64, reflect.Int32, reflect.Int16, reflect.Int8:
						if s == "发放人数" {
							numOfStaffTotal += int(value.Int())
						} else if s == "总计费用" {
							account += int(value.Int())
						}
						excel.SetCellInt("总览", pos(i+2, j), int(value.Int()))
					default:
						excel.SetCellValue("总览", pos(i+2, j), value)
					}
				}
			}

			if i%2 == 0 {
				excel.SetCellStyle("总览", pos(i+2, j), pos(i+2, j), cellStyle(excel, STYLE_TYPE_NORMAL))
			} else {
				excel.SetCellStyle("总览", pos(i+2, j), pos(i+2, j), cellStyle(excel, STYLE_TYPE_NORMAL_GREY))
			}
		}
	}

	excel.MergeCell("总览", pos(len(overviews)+2, 0), pos(len(overviews)+2, 1))

	excel.SetCellStr("总览", pos(len(overviews)+2, 0), "合计")
	excel.SetCellStyle("总览", pos(len(overviews)+2, 0), pos(len(overviews)+2, 1), cellStyle(excel, TYPE_ROW_TOTAL))

	excel.SetCellInt("总览", pos(len(overviews)+2, slices.Index(mConf.OverviewHeader, "发放人数")), numOfStaffTotal)
	excel.SetCellStyle("总览", pos(len(overviews)+2, slices.Index(mConf.OverviewHeader, "发放人数")), pos(len(overviews)+2, slices.Index(mConf.OverviewHeader, "发放人数")), cellStyle(excel, STYLE_TYPE_NORMAL))

	excel.SetCellInt("总览", pos(len(overviews)+2, slices.Index(mConf.OverviewHeader, "总计费用")), account)
	excel.SetCellStyle("总览", pos(len(overviews)+2, slices.Index(mConf.OverviewHeader, "总计费用")), pos(len(overviews)+2, slices.Index(mConf.OverviewHeader, "备注")), cellStyle(excel, STYLE_TYPE_TOTAL))

}

func constructSalarySheet(excel *excelize.File, sheetName string, salary map[string]Salary) {
	excel.NewSheet(sheetName)

	list := make([]Salary, len(salary)+1)
	total := Salary{}
	overview := Overview{Area: sheetName}

	calcTotal(&salary, &list, &total, &overview)

	overviewArr = append(overviewArr, overview)

	fillTitle(excel, sheetName, getTitle(sheetName, mConf.Month, mConf.Year))
	fillHeader(excel, sheetName, mConf.Headers)
	fillRow(excel, sheetName, sortSalaryById(salaryMap[sheetName]))
	fillTotal(excel, sheetName, len(salary)+2, total)

}

func calcTotal(salary *map[string]Salary, list *[]Salary, total *Salary, overview *Overview) {

	standardTotal := 0
	netpayTotal := 0
	accountTotal := 0
	numOfStaff := 0
	for _, item := range *salary {
		(*list)[item.Id-1] = item
		standardTotal += item.Standard
		netpayTotal += item.NetPay
		accountTotal += item.Account
		numOfStaff++
	}

	total.Name = "合计"
	total.Standard = standardTotal
	total.NetPay = netpayTotal
	total.Account = accountTotal

	overview.AccountTotal = accountTotal
	overview.NumOfStaff = numOfStaff
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
			switch {
			case maxLenForBackupMap[sheetName] < 10:
				excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 27.75)
			case maxLenForBackupMap[sheetName] < 15:
				excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 32.75)
			case maxLenForBackupMap[sheetName] < 20:
				excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 37.75)
			default:
				excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 42.75)
			}

		case header == "序号":
			excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 7)
		case len(header) < 4:
			excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 8.33)
		default:
			excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 11.33)
		}
	}
	excel.SetCellStyle(sheetName, pos(1, 0), pos(1, len(headers)-1), cellStyle(excel, TYPE_ROW_HEADER))
}

func fillRow(excel *excelize.File, sheetName string, salaries []Salary) {
	for i, salary := range salaries {
		for j, s := range mConf.Headers {
			v := reflect.ValueOf(salary)
			if v.Kind() == reflect.Struct {
				value := v.FieldByName(mConf.HeadersMap[s])

				kind := value.Kind()
				switch kind {
				case reflect.String:
					excel.SetCellStr(sheetName, pos(i+2, j), value.String())
				case reflect.Int, reflect.Int64, reflect.Int32, reflect.Int16, reflect.Int8:
					excel.SetCellInt(sheetName, pos(i+2, j), int(value.Int()))
				default:
					excel.SetCellValue(sheetName, pos(i+2, j), value)
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

type stylePar struct {
	Color     string
	Bold      bool
	Font      string
	Size      int
	FontColor string
}

func cellStyle(excel *excelize.File, styleNo int) int {

	if styleM == nil {
		styleM = make(map[int]int)
	}

	if styleId, found := styleM[styleNo]; found {
		return styleId
	}
	styleTemp := `{
		"border":[{"type":"left","color":"000000","style":1},
			{"type":"top","color":"000000","style":1},
			{"type":"right","color":"000000","style":1},
			{"type":"bottom","color":"000000","style":1}],
		"fill":{"type":"gradient","color":["{{.Color}}","{{.Color}}"], "shading":1},
		"alignment":{"horizontal":"center", "vertical":"center"},
		"font":{"bold":{{.Bold}}, "italic":false, "family":"{{.Font}}", "size":{{.Size}}, "color":"{{.FontColor}}"}}`

	t := template.Must(template.New("style").Parse(styleTemp))
	var tempBytes bytes.Buffer
	stylePar := stylePar{
		Color:     "#A5A5A5",
		Bold:      true,
		Font:      "Microsoft YaHei",
		Size:      16,
		FontColor: "#FFFFFF",
	}
	switch styleNo {
	case STYLE_TYPE_TITLE:
		stylePar.Color = "#A5A5A5"
		stylePar.Bold = true
		stylePar.Font = "Microsoft YaHei"
		stylePar.Size = 16
		stylePar.FontColor = "#FFFFFF"

	case STYLE_TYPE_HEADER:
		stylePar.Color = "#A5A5A5"
		stylePar.Bold = true
		stylePar.Font = "Microsoft YaHei"
		stylePar.Size = 14
		stylePar.FontColor = "#FFFFFF"

	case STYLE_TYPE_NORMAL:
		stylePar.Color = "#FFFFFF"
		stylePar.Bold = false
		stylePar.Font = "宋体"
		stylePar.Size = 12
		stylePar.FontColor = "#000000"

	case STYLE_TYPE_NORMAL_GREY:
		stylePar.Color = "#E7E6E6"
		stylePar.Bold = false
		stylePar.Font = "宋体"
		stylePar.Size = 12
		stylePar.FontColor = "#000000"

	case STYLE_TYPE_TOTAL:
		stylePar.Color = "#FFFFFF"
		stylePar.Bold = true
		stylePar.Font = "宋体"
		stylePar.Size = 12
		stylePar.FontColor = "#000000"

	}

	t.Execute(&tempBytes, stylePar)
	styleStr := tempBytes.String()
	// fmt.Printf("stylestr : %s", styleStr)
	styleId, err := excel.NewStyle(styleStr)
	if err != nil {
		fmt.Printf("%d err str %s", styleNo, styleStr)
		panic(err)
	}
	return styleId
}

func getTitle(sheetName string, month int, year int) string {
	return fmt.Sprintf("%s%d年%d月工资", sheetName, year, month)
}
