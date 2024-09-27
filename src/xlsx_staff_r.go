package main

import (
	"fmt"
	"reflect"
	"strconv"

	"github.com/tealeg/xlsx/v3"
)

const COL_STAFF_NAME = "人员姓名"
const COL_STAFF_SALARY = "薪资"
const COL_STAFF_QUIT_TIME = "离职时间"
const COL_STAFF_ACCOUNT_NAME = "代收人姓名"
const COL_STAFF_ACCOUNT = "收款账号"
const COL_STAFF_BACKUP = "备注"
const COL_STAFF_AREA = "区域"

const FINISH_SIGNAL_STAFF = "staff finish!!"

type Staff struct {
	Name     string
	Salary   int
	Account  string
	ToName   string
	Area     string
	QuitTime string
	BackUp   BackUpStaff
}

type BackUpStaff struct {
	BackUpSal []BackUpStaffSalary `json:"salary"`
}

type BackUpStaffSalary struct {
	Month []int `json:"month"`
	Sal   int   `json:"sal"`
}

func readFromXlsxStaff(staffChan chan Staff, finishChan chan string) error {
	file, err := xlsx.OpenFile(mConf.StaffFilePath)
	if err != nil {
		fmt.Println(err)
		finishChan <- err.Error()
		return err
	}

	for _, sheet := range file.Sheets {

		if sheet.Name != "员工详情（全部）" {
			continue
		}
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

			staffChan <- staff
		}
	}

	finishChan <- FINISH_SIGNAL_STAFF

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
