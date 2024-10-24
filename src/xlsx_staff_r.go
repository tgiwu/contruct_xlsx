package main

import (
	"encoding/json"
	"fmt"
	"reflect"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx/v3"
)

const COL_STAFF_NAME = "人员姓名"
const COL_STAFF_SALARY = "薪资"
const COL_STAFF_QUIT_TIME = "离职时间"
const COL_STAFF_ACCOUNT_NAME = "代收人姓名"
const COL_STAFF_ACCOUNT = "收款账号"
const COL_STAFF_BACKUP = "备注"
const COL_STAFF_AREA = "区域"

const COL_SS_TYPE = "类型"
const COL_SS_SALARY_PER_DAY = "日薪"
const COL_SS_DESCRIPTION = "说明"

const FINISH_SIGNAL_STAFF = "staff finish!!"
const FINISH_SIGNAL_SALARY_STANDARDS = "salary standards finish!!"

type Staff struct {
	Name     string
	Salary   int
	Account  string
	ToName   string
	Area     string
	QuitTime string
	BackUp   BackUpStaff
}

type SalaryStandards struct {
	TempType     string //临勤类型
	SalaryPerDay int    //日薪
	Description  string //说明
}

type BackUpStaff struct {
	BackUpSal []BackUpStaffSalary `json:"salary"`
}

type BackUpStaffSalary struct {
	Month []int `json:"month"`
	Sal   int   `json:"sal"`
}

func readSalaryStandard(sheet *xlsx.Sheet, ssChan chan SalaryStandards, finishChan chan string) error {

	headerMap := make(map[int]string)

	for i := range sheet.MaxRow {
		row, err := sheet.Row(i)

		if err != nil {
			return err
		}

		var ss SalaryStandards

		err = visitRowSS(row, &headerMap, &ss)

		if err != nil {
			return err
		}

		if len(ss.TempType) != 0 {
			ssChan <- ss
		}
	}

	finishChan <- FINISH_SIGNAL_SALARY_STANDARDS

	return nil
}

func visitRowSS(row *xlsx.Row, headerMap *map[int]string, ss *SalaryStandards) error {
	isReadHeader := len(*headerMap) == 0

	for i := range row.Sheet.MaxCol {
		str, err := row.GetCell(i).FormattedValue()

		if err != nil {
			continue
		}

		if isReadHeader {
			switch str {
			case COL_SS_DESCRIPTION:
				(*headerMap)[i] = "Description"
			case COL_SS_TYPE:
				(*headerMap)[i] = "TempType"
			case COL_SS_SALARY_PER_DAY:
				(*headerMap)[i] = "SalaryPerDay"
			}
		} else {
			val, _ := strconv.Atoi(str)
			refType := reflect.TypeOf(*ss)
			if refType.Kind() != reflect.Struct {
				panic("not struct")
			}

			if fieldObj, ok := refType.FieldByName((*headerMap)[i]); ok {

				if fieldObj.Type.Kind() == reflect.Int {
					reflect.ValueOf(ss).Elem().FieldByName((*headerMap)[i]).SetInt(int64(val))

				}
				if fieldObj.Type.Kind() == reflect.String {
					reflect.ValueOf(ss).Elem().FieldByName((*headerMap)[i]).SetString(str)
				}
			}
		}
	}

	return nil
}

func readStaff(sheet *xlsx.Sheet, staffChan chan Staff, finishChan chan string) error {
	headerMap := make(map[int]string)

	for i := range sheet.MaxRow {
		row, err := sheet.Row(i)

		if err != nil {
			return err
		}

		var staff Staff

		visitRow(row, &headerMap, &staff)

		if len(staff.Name) == 0 {
			continue
		}

		//代发人员计入范崎路
		if staff.Area == "代发工资" {
			staff.Area = "范崎路"
		}

		staffChan <- staff
	}

	finishChan <- FINISH_SIGNAL_STAFF

	return nil
}

func readData(staffChan chan Staff, ssChan chan SalaryStandards, finishChan chan string) error {
	file, err := xlsx.OpenFile(mConf.StaffFilePath)
	if err != nil {
		fmt.Println(err)
		finishChan <- err.Error()
		return err
	}

	for _, sheet := range file.Sheets {

		switch sheet.Name {
		case mConf.StaffSheetName:
			err = readStaff(sheet, staffChan, finishChan)
		case mConf.SalaryStandardsSheetName:
			err = readSalaryStandard(sheet, ssChan, finishChan)
		}

		if err != nil {
			return err
		}
	}

	return nil
}

func visitRow(row *xlsx.Row, headerMap *map[int]string, staff *Staff) {
	isReadHeader := len(*headerMap) == 0

	for i := range row.Sheet.MaxCol {
		str, err := row.GetCell(i).FormattedValue()
		if err != nil {
			continue
		}
		if isReadHeader {

			switch str {
			case COL_STAFF_NAME:
				(*headerMap)[i] = "Name"
			case COL_STAFF_SALARY:
				(*headerMap)[i] = "Salary"
			case COL_STAFF_QUIT_TIME:
				(*headerMap)[i] = "QuitTime"
			case COL_STAFF_ACCOUNT_NAME:
				(*headerMap)[i] = "ToName"
			case COL_STAFF_ACCOUNT:
				(*headerMap)[i] = "Account"
			case COL_STAFF_BACKUP:
				(*headerMap)[i] = "BackUp"
			case COL_STAFF_AREA:
				(*headerMap)[i] = "Area"
			}
		} else {
			val, _ := strconv.Atoi(str)
			refType := reflect.TypeOf(*staff)
			if refType.Kind() != reflect.Struct {
				panic("not struct")
			}

			if fieldObj, ok := refType.FieldByName((*headerMap)[i]); ok {

				if (*headerMap)[i] == "BackUp" && strings.Index(str, "json:") == 0 {
					str = str[len("json:"):]
					var backupStaff BackUpStaff
					err := json.Unmarshal([]byte(str), &backupStaff)
					if err != nil {
						fmt.Printf("unmarshal backup %s failed\n %v", str, err)
					} else {
						staff.BackUp = backupStaff
					}
					continue
				}

				if fieldObj.Type.Kind() == reflect.Int {
					reflect.ValueOf(staff).Elem().FieldByName((*headerMap)[i]).SetInt(int64(val))

				}
				if fieldObj.Type.Kind() == reflect.String {
					reflect.ValueOf(staff).Elem().FieldByName((*headerMap)[i]).SetString(str)
				}
			}
		}
	}
}
