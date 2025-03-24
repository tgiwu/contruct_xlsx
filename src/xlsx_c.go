package main

import (
	"bytes"
	"fmt"
	"path/filepath"
	"reflect"
	"slices"
	"strings"
	"text/template"

	"github.com/360EntSecGroup-Skylar/excelize"
)

const STYLE_TYPE_TITLE = 0
const STYLE_TYPE_HEADER = 1
const STYLE_TYPE_NORMAL = 2
const STYLE_TYPE_NORMAL_GREY = 3
const STYLE_TYPE_TOTAL = 4
const STYLE_TYPE_ERROR = 5

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

// 构建excel
func constructSalaryXlsx(salaryMap map[string]map[string]Salary, fileName string, finishChan chan string) error {
	fmt.Printf("construct xlsx %s start\n", fileName)
	excel := excelize.NewFile()

	keys := make([]string, 0, len(salaryMap))
	for k := range salaryMap {
		keys = append(keys, k)
	}

	slices.Sort(keys)
	// overviewArr will be used multi times, empty it before use
	overviewArr = make([]Overview, 0)

	//生成单个工作表
	for _, key := range keys {
		constructSalarySheet(excel, key, salaryMap[key])
	}
	//生成总览表
	constructOverviewSheet(excel, overviewArr)
	//删除默认工作表
	excel.DeleteSheet("Sheet1")

	delFileIfExist(mConf.OutputPath, fileName)
	excel.SaveAs(filepath.Join(mConf.OutputPath, fileName))
	finishChan <- fmt.Sprintf("%s finish !!", fileName)

	return nil
}

func constructOverviewSheet(excel *excelize.File, overviews []Overview) {

	slices.SortFunc(overviews, func(a, b Overview) int {
		return strings.Compare(a.Area, b.Area)
	})

	excel.NewSheet(SALARY_SHEET_NAME_OVERVIEW)

	excel.MergeCell(SALARY_SHEET_NAME_OVERVIEW, pos(0, 0), pos(0, len(mConf.OverviewHeader)-1))
	excel.SetCellValue(SALARY_SHEET_NAME_OVERVIEW, pos(0, 0), fmt.Sprintf("%d年%d月工资总览", mConf.Year, mConf.Month))
	excel.SetCellStyle(SALARY_SHEET_NAME_OVERVIEW, pos(0, 0), pos(0, len(mConf.OverviewHeader)-1), cellStyle(excel, TYPE_ROW_TITLE))

	fillHeader(excel, SALARY_SHEET_NAME_OVERVIEW, mConf.OverviewHeader)

	numOfStaffTotal := 0
	account := 0
	for i, overview := range overviews {
		for j, s := range mConf.OverviewHeader {

			if s == "序号" {
				excel.SetCellInt(SALARY_SHEET_NAME_OVERVIEW, pos(i+2, j), i+1)
			} else {
				v := reflect.ValueOf(overview)
				if v.Kind() == reflect.Struct {
					value := v.FieldByName(mConf.OverviewHeaderMap[s])

					kind := value.Kind()
					switch kind {
					case reflect.String:
						excel.SetCellStr(SALARY_SHEET_NAME_OVERVIEW, pos(i+2, j), value.String())
					case reflect.Int, reflect.Int64, reflect.Int32, reflect.Int16, reflect.Int8:
						if s == SALARY_OVERVIEW_COLUMN_NUMBER {
							numOfStaffTotal += int(value.Int())
						} else if s == SALARY_OVERVIEW_COLUMN_SALARY {
							account += int(value.Int())
						}
						excel.SetCellInt(SALARY_SHEET_NAME_OVERVIEW, pos(i+2, j), int(value.Int()))
					default:
						excel.SetCellValue(SALARY_SHEET_NAME_OVERVIEW, pos(i+2, j), value)
					}
				}
			}

			if i%2 == 0 {
				excel.SetCellStyle(SALARY_SHEET_NAME_OVERVIEW, pos(i+2, j), pos(i+2, j), cellStyle(excel, STYLE_TYPE_NORMAL))
			} else {
				excel.SetCellStyle(SALARY_SHEET_NAME_OVERVIEW, pos(i+2, j), pos(i+2, j), cellStyle(excel, STYLE_TYPE_NORMAL_GREY))
			}
		}
	}

	excel.MergeCell(SALARY_SHEET_NAME_OVERVIEW, pos(len(overviews)+2, 0), pos(len(overviews)+2, 1))

	excel.SetCellStr(SALARY_SHEET_NAME_OVERVIEW, pos(len(overviews)+2, 0), "合计")
	excel.SetCellStyle(SALARY_SHEET_NAME_OVERVIEW, pos(len(overviews)+2, 0), pos(len(overviews)+2, 1), cellStyle(excel, TYPE_ROW_TOTAL))

	excel.SetCellInt(SALARY_SHEET_NAME_OVERVIEW, pos(len(overviews)+2, slices.Index(mConf.OverviewHeader, SALARY_OVERVIEW_COLUMN_NUMBER)), numOfStaffTotal)
	excel.SetCellStyle(SALARY_SHEET_NAME_OVERVIEW, pos(len(overviews)+2, slices.Index(mConf.OverviewHeader, SALARY_OVERVIEW_COLUMN_NUMBER)), pos(len(overviews)+2, slices.Index(mConf.OverviewHeader, "发放人数")), cellStyle(excel, STYLE_TYPE_NORMAL))

	excel.SetCellInt(SALARY_SHEET_NAME_OVERVIEW, pos(len(overviews)+2, slices.Index(mConf.OverviewHeader, SALARY_OVERVIEW_COLUMN_SALARY)), account)
	excel.SetCellStyle(SALARY_SHEET_NAME_OVERVIEW, pos(len(overviews)+2, slices.Index(mConf.OverviewHeader, SALARY_OVERVIEW_COLUMN_SALARY)), pos(len(overviews)+2, slices.Index(mConf.OverviewHeader, "备注")), cellStyle(excel, STYLE_TYPE_TOTAL))

}

func constructSalarySheet(excel *excelize.File, sheetName string, salary map[string]Salary) {
	excel.NewSheet(sheetName)

	list := make([]Salary, len(salary)+1)
	total := Salary{}
	overview := Overview{Area: sheetName}

	calcTotal(&salary, &list, &total, &overview)

	overviewArr = append(overviewArr, overview)

	fillTitle(excel, sheetName, getTitle(sheetName, mConf.Month, mConf.Year))
	fillHeader(excel, sheetName, mConf.HeadersRisk)
	fillRow(excel, sheetName, sortSalaryById(salary))
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

	indexStandard, indexNetPay, indexAccount := 0, 0, 0

	for i, s := range mConf.Headers {
		switch s {
		case "应发工资":
			indexStandard = i
		case "实发工资":
			indexNetPay = i
		case "合计":
			indexAccount = i
		}
	}

	total.totalStandard = fmt.Sprintf("=SUM(%s:%s)", pos(2, indexStandard), pos(len(*list), indexStandard))
	total.totalNetPay = fmt.Sprintf("=SUM(%s:%s)", pos(2, indexNetPay), pos(len(*list), indexNetPay))
	total.totalAccount = fmt.Sprintf("=SUM(%s:%s)", pos(2, indexAccount), pos(len(*list), indexAccount))

	total.Name = "合计"

	overview.AccountTotal = accountTotal
	overview.NumOfStaff = numOfStaff
}

func fillTitle(excel *excelize.File, sheetName string, title string) {
	excel.MergeCell(sheetName, pos(0, 0), pos(0, len(mConf.HeadersRisk)-1))
	excel.SetCellValue(sheetName, pos(0, 0), title)
	excel.SetCellStyle(sheetName, pos(0, 0), pos(0, len(mConf.HeadersRisk)-1), cellStyle(excel, TYPE_ROW_TITLE))
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
				excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 45.75)
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
	var errCells = make(map[string]string, 0)
	for i, salary := range salaries {
		for j, s := range mConf.HeadersRisk {
			v := reflect.ValueOf(salary)
			if v.Kind() == reflect.Struct {
				value := v.FieldByName(mConf.HeadersRiskMap[s])

				kind := value.Kind()
				switch kind {
				case reflect.String:
					if strings.HasPrefix(value.String(), "=") {
						excel.SetCellFormula(sheetName, pos(i+2, j), value.String())
					} else {
						excel.SetCellStr(sheetName, pos(i+2, j), value.String())
					}
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

			if len(salary.ErrorMap) > 0 {
				fmt.Println("has error")
			}
			comment, f := salary.ErrorMap[s]
			if f {
				errCells[pos(i+2, j)] = comment
			}
		}
	}
	//mark error
	if len(errCells) > 0 {
		for p, comment := range errCells {
			excel.SetCellStyle(sheetName, p, p, cellStyle(excel, STYLE_TYPE_ERROR))
			excel.AddComment(sheetName, p, fmt.Sprintf(`{"author":"Robot: ","text":"%s"}`, comment))
		}
	}
}

func fillTotal(excel *excelize.File, sheetName string, row int, total Salary) {
	excel.MergeCell(sheetName, pos(row, 0), pos(row, 3))
	excel.SetCellValue(sheetName, pos(row, 0), total.Name)
	excel.SetCellStyle(sheetName, pos(row, 0), pos(row, 3), cellStyle(excel, TYPE_ROW_TOTAL))

	excel.SetCellFormula(sheetName, pos(row, 4), total.totalStandard)
	excel.SetCellFormula(sheetName, pos(row, 5), total.totalNetPay)

	indexOfTotal := -1
	for index, name := range mConf.HeadersRisk {
		if name == "合计" {
			indexOfTotal = index
		}
	}

	if indexOfTotal == -1 {
		panic("can not locate column total!")
	}

	excel.SetCellStyle(sheetName, pos(row, 4), pos(row, indexOfTotal), cellStyle(excel, TYPE_ROW_NORMAL))

	excel.SetCellFormula(sheetName, pos(row, indexOfTotal), total.totalAccount)
	excel.SetCellStyle(sheetName, pos(row, indexOfTotal), pos(row, indexOfTotal), cellStyle(excel, TYPE_ROW_TOTAL))
	//set style for columns behind total
	excel.SetCellStyle(sheetName, pos(row, indexOfTotal+1), pos(row, len(mConf.HeadersRisk)-1), cellStyle(excel, TYPE_ROW_TOTAL))
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
	case STYLE_TYPE_ERROR:
		stylePar.Color = "#FF0000"
		stylePar.Bold = true
		stylePar.Font = "宋体"
		stylePar.Size = 13
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

func constructTransferInfoXlsx(transferInfos *[]TransferInfo, fileName string, finishChan chan string) {
	excel := excelize.NewFile()

	constructTransferInfoSheet(excel, "transferInfo", transferInfos)

	//删除默认工作表
	excel.DeleteSheet("Sheet1")

	filePath := filepath.Join(mConf.OutputPath, fileName)
		
	delFileIfExist(mConf.OutputPath, fileName)
	excel.SaveAs(filePath)
	finishChan <- fmt.Sprintf("%s finish !!", fileName)

}

func constructTransferInfoSheet(excel *excelize.File, sheet string, transferInfos *[]TransferInfo) {
	excel.NewSheet(sheet)
	fillTransferInfoTitle(excel, sheet, "转账信息")
	fillTransferInfoHeader(excel, sheet, &TRANSFER_INFO_COLUNM)
	fillTransferinfoRows(excel, sheet, transferInfos)

}

func fillTransferInfoTitle(excel *excelize.File, sheet string, title string) {
	excel.MergeCell(sheet, pos(0, 0), pos(0, len(TRANSFER_INFO_COLUNM)-1))
	excel.SetCellValue(sheet, pos(0, 0), title)
	excel.SetCellStyle(sheet, pos(0, 0), pos(0, len(TRANSFER_INFO_COLUNM)-1), cellStyle(excel, TYPE_ROW_TITLE))
}

func fillTransferInfoHeader(excel *excelize.File, sheet string, colonms *[]string) {
	for i, header := range *colonms {
		excel.SetCellValue(sheet, pos(1, i), header)

		switch header {
		case TRANSFER_INFO_COLUNM[0], TRANSFER_INFO_COLUNM[1], TRANSFER_INFO_COLUNM[3]:
			excel.SetColWidth(sheet, pos(-1, i), pos(-1, i), 25.33)
		case TRANSFER_INFO_COLUNM[2], TRANSFER_INFO_COLUNM[4], TRANSFER_INFO_COLUNM[5]:
			excel.SetColWidth(sheet, pos(-1, i), pos(-1, i), 15.33)
		default:
			excel.SetColWidth(sheet, pos(-1, i), pos(-1, i), 20.83)
		}
	}
	excel.SetCellStyle(sheet, pos(1, 0), pos(1, len(*colonms)-1), cellStyle(excel, TYPE_ROW_HEADER))
}

func fillTransferinfoRows(excel *excelize.File, sheet string, transferInfos *[]TransferInfo) {
	for i, info := range *transferInfos {
		for j, tag := range TRANSFER_INFO_COLUNM_TAG {
			v := reflect.ValueOf(info)
			if v.Kind() == reflect.Struct {
				value := v.FieldByName(tag)

				kind := value.Kind()
				switch kind {
				case reflect.String:
					excel.SetCellStr(sheet, pos(i+2, j), value.String())
				case reflect.Int, reflect.Int64, reflect.Int32, reflect.Int16, reflect.Int8:
					excel.SetCellInt(sheet, pos(i+2, j), int(value.Int()))
				default:
					excel.SetCellValue(sheet, pos(i+2, j), value)
				}
			}
			if i%2 == 0 {
				excel.SetCellStyle(sheet, pos(i+2, j), pos(i+2, j), cellStyle(excel, STYLE_TYPE_NORMAL))
			} else {
				excel.SetCellStyle(sheet, pos(i+2, j), pos(i+2, j), cellStyle(excel, STYLE_TYPE_NORMAL_GREY))
			}
		}
	}
}
