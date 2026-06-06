package main

import (
	"fmt"
	"path/filepath"
	"reflect"
	"slices"
	"strings"

	"github.com/xuri/excelize/v2"

	log "github.com/sirupsen/logrus"
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

var (
	styleCellTitle  int
	styleCellHeader int
	styleCellNormal int
	styleCellGrey   int
	styleCellTotal  int
	styleCellErr    int
)

type EmptyError struct {
	msg string
}

func (ee EmptyError) Error() string {
	return ee.msg
}

func setUpCellStyle(excel *excelize.File) {
	styleCellTitle, _ = excel.NewStyle(&excelize.Style{
		Border: []excelize.Border{{Type: "left", Color: "#FFFFFF", Style: 1},
			{Type: "right", Color: "#FFFFFF", Style: 1},
			{Type: "top", Color: "#FFFFFF", Style: 1},
			{Type: "bottom", Color: "#FFFFFF", Style: 1}},
		Font:      &excelize.Font{Bold: true, Color: "#FFFFFF", Family: "Microsoft YaHei", Size: 16},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#A5A5A5"}, Pattern: 1},
	})

	styleCellHeader, _ = excel.NewStyle(&excelize.Style{
		Border: []excelize.Border{{Type: "left", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1}},
		Font:      &excelize.Font{Bold: true, Color: "#FFFFFF", Family: "Microsoft YaHei", Size: 14},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#A5A5A5"}, Pattern: 1},
	})

	styleCellNormal, _ = excel.NewStyle(&excelize.Style{
		Border: []excelize.Border{{Type: "left", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1}},
		Font:      &excelize.Font{Bold: false, Color: "#000000", Family: "宋体", Size: 14},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#FFFFFF"}, Pattern: 1},
	})

	styleCellGrey, _ = excel.NewStyle(&excelize.Style{
		Border: []excelize.Border{{Type: "left", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1}},
		Font:      &excelize.Font{Bold: false, Color: "#000000", Family: "宋体", Size: 14},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#E6E6E6"}, Pattern: 1},
	})

	styleCellTotal, _ = excel.NewStyle(&excelize.Style{
		Border: []excelize.Border{{Type: "left", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1}},
		Font:      &excelize.Font{Bold: true, Color: "#000000", Family: "宋体", Size: 14},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#FFFFFF"}, Pattern: 1},
	})

	styleCellErr, _ = excel.NewStyle(&excelize.Style{
		Border: []excelize.Border{{Type: "left", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1}},
		Font:      &excelize.Font{Bold: true, Color: "#000000", Family: "宋体", Size: 14},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#FF0000"}, Pattern: 1},
	})

}

// construct salary xlsx by area
func constructSalaryXlsx(salaryMap map[string][]Salary, fileName string, finishChan chan string) error {

	log.Infof("construct xlsx %s start\n", fileName)
	excel := excelize.NewFile()
	defer excel.Close()
	setUpCellStyle(excel)
	keys := make([]string, 0, len(salaryMap))
	for k := range salaryMap {
		keys = append(keys, k)
	}

	slices.Sort(keys)

	//contruct single table
	for _, key := range keys {
		constructSalarySheet(excel, key, salaryMap[key])
	}

	//删除默认工作表
	excel.DeleteSheet("Sheet1")

	delFileIfExist(mConf.OutputPath, fileName)
	err := excel.SaveAs(filepath.Join(mConf.OutputPath, fileName))
	if err != nil {
		panic(err)
	}

	finishChan <- fmt.Sprintf("%s finish !!", fileName)

	return nil
}

func constructSalaryRiskXlsx(salaryMap map[string][]Salary, fileName string, finishChan chan string) error {
	log.Infof("construct xlsx %s start\n", fileName)
	excel := excelize.NewFile()

	defer excel.Close()
	setUpCellStyle(excel)

	keys := make([]string, 0, len(salaryMap))
	for k := range salaryMap {
		keys = append(keys, k)
	}

	slices.SortFunc(keys, func(s1, s2 string) int {
		switch {
		case s1 == "总览":
			return -1
		case s2 == "总览":
			return 1
		default:
			return 0
		}
	})

	//contruct single table
	for _, key := range keys {
		if key == "总览" {
			constructOverviewSheet(excel, salaryRiskMap2[key])
		} else {
			constructSalarySheet(excel, key, salaryMap[key])
		}
	}

	//删除默认工作表
	excel.DeleteSheet("Sheet1")

	delFileIfExist(mConf.OutputPath, fileName)
	excel.SaveAs(filepath.Join(mConf.OutputPath, fileName))
	finishChan <- fmt.Sprintf("%s finish !!", fileName)

	return nil
}

func constructOverviewSheet(excel *excelize.File, items []Salary) {
	excel.NewSheet(SALARY_SHEET_NAME_OVERVIEW)

	excel.MergeCell(SALARY_SHEET_NAME_OVERVIEW, pos(0, 0), pos(0, len(mConf.OverviewHeader)-1))
	excel.SetCellValue(SALARY_SHEET_NAME_OVERVIEW, pos(0, 0), fmt.Sprintf("%d年%d月工资总览", mConf.Year, mConf.Month))
	excel.SetCellStyle(SALARY_SHEET_NAME_OVERVIEW, pos(0, 0), pos(0, len(mConf.OverviewHeader)-1), styleCellTitle)

	fillHeader(excel, SALARY_SHEET_NAME_OVERVIEW, mConf.OverviewHeader)
	excel.SetCellStyle(SALARY_SHEET_NAME_OVERVIEW, pos(1, 0), pos(1, len(mConf.OverviewHeader)-1), styleCellHeader)

	for i, item := range items {

		if item.StaffId == 999 {
			fillTotalRisk(excel, SALARY_SHEET_NAME_OVERVIEW, item, i+2)
			break
		}
		for j, s := range mConf.OverviewHeader {

			if s == "序号" {
				excel.SetCellInt(SALARY_SHEET_NAME_OVERVIEW, pos(i+2, j), int64(i+1))
			} else {
				v := reflect.ValueOf(item)
				if v.Kind() == reflect.Struct {
					value := v.FieldByName(mConf.OverviewHeaderMap[s])

					kind := value.Kind()
					switch kind {
					case reflect.String:
						if strings.HasPrefix(value.String(), "=") {
							excel.SetCellFormula(SALARY_SHEET_NAME_OVERVIEW, pos(i+2, j), value.String())
						} else {
							excel.SetCellStr(SALARY_SHEET_NAME_OVERVIEW, pos(i+2, j), value.String())
						}
					case reflect.Int, reflect.Int64, reflect.Int32, reflect.Int16, reflect.Int8:
						excel.SetCellInt(SALARY_SHEET_NAME_OVERVIEW, pos(i+2, j), value.Int())
					default:
						excel.SetCellValue(SALARY_SHEET_NAME_OVERVIEW, pos(i+2, j), value)
					}
				}
			}

			if i%2 == 0 {
				excel.SetCellStyle(SALARY_SHEET_NAME_OVERVIEW, pos(i+2, j), pos(i+2, j), styleCellNormal)
			} else {
				excel.SetCellStyle(SALARY_SHEET_NAME_OVERVIEW, pos(i+2, j), pos(i+2, j), styleCellGrey)
			}
		}
	}
}

func fillTotalRisk(excel *excelize.File, sheetName string, item Salary, row int) {
	excel.MergeCell(sheetName, pos(row, 0), pos(row, 1))
	excel.SetCellValue(sheetName, pos(row, 0), item.Area)
	excel.SetCellStyle(sheetName, pos(row, 0), pos(row, 1), styleCellTotal)

	personSumIndex, accountSumIndex := -1, -1
	for index, name := range mConf.OverviewHeader {
		switch name {
		case "发放人数":
			personSumIndex = index
		case "总计费用":
			accountSumIndex = index
		}
	}

	if personSumIndex == -1 || accountSumIndex == -1 {
		panic("can not locate column total!")
	}

	excel.SetCellFormula(sheetName, pos(row, personSumIndex), item.TotalPerson)
	excel.SetCellStyle(sheetName, pos(row, 0), pos(row, personSumIndex), styleCellTotal)

	excel.SetCellFormula(sheetName, pos(row, accountSumIndex), item.TotalAccount)
	excel.SetCellStyle(sheetName, pos(row, accountSumIndex), pos(row, accountSumIndex), styleCellTotal)
	//set style for columns behind total
	excel.SetCellStyle(sheetName, pos(row, accountSumIndex+1), pos(row, len(mConf.OverviewHeader)-1), styleCellNormal)
}

func constructSalarySheet(excel *excelize.File, sheetName string, salaries []Salary) {
	excel.NewSheet(sheetName)

	fillTitle(excel, sheetName, getTitle(sheetName, mConf.Month, mConf.Year))
	fillHeader(excel, sheetName, mConf.HeadersRisk)
	fillRow(excel, sheetName, salaries)
}

func fillTitle(excel *excelize.File, sheetName string, title string) {
	excel.MergeCell(sheetName, pos(0, 0), pos(0, len(mConf.HeadersRisk)-1))
	excel.SetRowHeight(sheetName, 1, 25)
	excel.SetCellValue(sheetName, pos(0, 0), title)
	excel.SetCellStyle(sheetName, pos(0, 0), pos(0, len(mConf.HeadersRisk)-1), styleCellTitle)
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
				excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 55.75)
			}

		case header == "序号":
			excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 7)
		case len(header) < 4:
			excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 9.33)
		default:
			excel.SetColWidth(sheetName, pos(-1, i), pos(-1, i), 12.83)
		}
	}
	excel.SetCellStyle(sheetName, pos(1, 0), pos(1, len(headers)-1), styleCellHeader)
}

func fillRow(excel *excelize.File, sheetName string, salaries []Salary) {
	var errCells = make(map[string]string, 0)
	for i, salary := range salaries {

		if salary.StaffId == 999 {
			fillTotal(excel, sheetName, len(salaries)+1, salary)
			break
		}

		for j, s := range mConf.HeadersRisk {
			v := reflect.ValueOf(salary)
			if v.Kind() == reflect.Struct {
				value := v.FieldByName(mConf.HeadersRiskMap[s])

				//
				if mConf.HeadersRiskMap[s] == "RowId" {
					excel.SetCellInt(sheetName, pos(i+2, j), int64(i+1))
					continue
				}
				kind := value.Kind()
				switch kind {
				case reflect.String:
					if strings.HasPrefix(value.String(), "=") {
						excel.SetCellFormula(sheetName, pos(i+2, j), value.String())
					} else {
						excel.SetCellStr(sheetName, pos(i+2, j), value.String())
					}
				case reflect.Int, reflect.Int64, reflect.Int32, reflect.Int16, reflect.Int8:
					excel.SetCellInt(sheetName, pos(i+2, j), value.Int())
				default:
					excel.SetCellValue(sheetName, pos(i+2, j), value)
				}
			}
			if i%2 == 0 {
				excel.SetCellStyle(sheetName, pos(i+2, j-1), pos(i+2, j), styleCellNormal)
			} else {
				excel.SetCellStyle(sheetName, pos(i+2, j-1), pos(i+2, j), styleCellGrey)
			}

			if len(salary.ErrorMap) > 0 {
				log.Infoln("has error")
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
			excel.SetCellStyle(sheetName, p, p, styleCellErr)
			excel.AddComment(sheetName, excelize.Comment{Author: "Robot:", Text: comment, Cell: p})
		}
	}
}

func fillTotal(excel *excelize.File, sheetName string, row int, total Salary) {
	excel.MergeCell(sheetName, pos(row, 0), pos(row, 3))
	excel.SetCellValue(sheetName, pos(row, 0), total.Name)
	excel.SetCellStyle(sheetName, pos(row, 0), pos(row, 3), styleCellTotal)

	excel.SetCellFormula(sheetName, pos(row, 4), total.TotalStandard)
	excel.SetCellFormula(sheetName, pos(row, 5), total.TotalNetPay)

	indexOfTotal := -1
	for index, name := range mConf.HeadersRisk {
		if name == "合计" {
			indexOfTotal = index
		}
	}

	if indexOfTotal == -1 {
		panic("can not locate column total!")
	}

	excel.SetCellStyle(sheetName, pos(row, 0), pos(row, indexOfTotal), styleCellTotal)

	excel.SetCellFormula(sheetName, pos(row, indexOfTotal), total.TotalAccount)
	excel.SetCellStyle(sheetName, pos(row, indexOfTotal), pos(row, indexOfTotal), styleCellTotal)
	//set style for columns behind total
	excel.SetCellStyle(sheetName, pos(row, indexOfTotal+1), pos(row, len(mConf.HeadersRisk)-1), styleCellNormal)
}

func getTitle(sheetName string, month int, year int) string {
	return fmt.Sprintf("%s%d年%d月工资", sheetName, year, month)
}

// construct transferInfo xlsx,ignore style
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
	// excel.SetCellStyle(sheet, pos(0, 0), pos(0, len(TRANSFER_INFO_COLUNM)-1), styleCellTitle)
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
	// excel.SetCellStyle(sheet, pos(1, 0), pos(1, len(*colonms)-1), styleCellHeader)
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
					excel.SetCellInt(sheet, pos(i+2, j), value.Int())
				default:
					excel.SetCellValue(sheet, pos(i+2, j), value)
				}
			}
			// if i%2 == 0 {
			// 	excel.SetCellStyle(sheet, pos(i+2, j), pos(i+2, j), styleCellNormal)
			// } else {
			// 	excel.SetCellStyle(sheet, pos(i+2, j), pos(i+2, j), styleCellGrey)
			// }
		}
	}
}
