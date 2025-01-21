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

const COL_SP_TYPE = "类型"
const COL_SP_SALARY_PER_MONTH = "月薪"
const COL_SP_DESCRIPTION = "说明"

const FINISH_SIGNAL_STAFF = "staff finish!!"
const FINISH_SIGNAL_SALARY_STANDARDS_TEMP = "salary standards finish!!"
const FINISH_SIGNAL_SALARY_STANDARDS_POST = "salary standards post finish"

type Staff struct {
	Name     string      //姓名
	Salary   int         //工资
	Account  string      //收款账号
	ToName   string      //收款人姓名
	Area     string      //区域
	QuitTime string      //离职时间
	BackUp   BackUpStaff //备注
	Calc     Calc        //工资计算方式
	Sal      *Salary     //工资计算结果
	Att      *Attendance //考勤
}

type Calc func(staff *Staff, attendance *Attendance, salary *Salary) error



type BackUpStaff struct {
	BackUpSal []BackUpStaffSalary `json:"salary"`
}

type BackUpStaffSalary struct {
	Month []int `json:"month"`
	Sal   int   `json:"sal"`
}

//读取临勤工资标准
func readSalaryStandardTemp(sheet *xlsx.Sheet, ssChan chan SalaryStandardsTemp, finishChan chan string) error {

	headerMap := make(map[int]string)

	for i := range sheet.MaxRow {
		row, err := sheet.Row(i)

		if err != nil {
			return err
		}

		var ss SalaryStandardsTemp

		err = visitRowSS(row, &headerMap, &ss)

		if err != nil {
			return err
		}

		if len(ss.TempType) != 0 {
			ssChan <- ss
		}
	}

	finishChan <- FINISH_SIGNAL_SALARY_STANDARDS_TEMP

	return nil
}

//读取借调工资标准
func readSalaryStandardPost(sheet *xlsx.Sheet, spChan chan SalaryStandardsPost, finishChan chan string) error {

	headerMap := make(map[int]string)

	for i := range sheet.MaxRow {
		row, err := sheet.Row(i)

		if err != nil {
			return err
		}

		var sp SalaryStandardsPost

		err = visitRowSP(row, &headerMap, &sp)

		if err != nil {
			return err
		}

		if len(sp.PostType) != 0 {
			spChan <- sp
		}
	}

	finishChan <- FINISH_SIGNAL_SALARY_STANDARDS_POST

	return nil
}

//单行临勤工资标准
func visitRowSS(row *xlsx.Row, headerMap *map[int]string, ss *SalaryStandardsTemp) error {
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

//单行借调工资标准
func visitRowSP(row *xlsx.Row, headerMap *map[int]string, sp *SalaryStandardsPost) error {
	isReadHeader := len(*headerMap) == 0

	for i := range row.Sheet.MaxCol {
		str, err := row.GetCell(i).FormattedValue()

		if err != nil {
			continue
		}

		if isReadHeader {
			switch str {
			case COL_SP_DESCRIPTION:
				(*headerMap)[i] = "Description"
			case COL_SP_TYPE:
				(*headerMap)[i] = "PostType"
			case COL_SP_SALARY_PER_MONTH:
				(*headerMap)[i] = "SalaryPerMonth"
			}
		} else {
			val, _ := strconv.Atoi(str)
			refType := reflect.TypeOf(*sp)
			if refType.Kind() != reflect.Struct {
				panic("not struct")
			}

			if fieldObj, ok := refType.FieldByName((*headerMap)[i]); ok {

				if fieldObj.Type.Kind() == reflect.Int {
					reflect.ValueOf(sp).Elem().FieldByName((*headerMap)[i]).SetInt(int64(val))

				}
				if fieldObj.Type.Kind() == reflect.String {
					reflect.ValueOf(sp).Elem().FieldByName((*headerMap)[i]).SetString(str)
				}
			}
		}
	}

	return nil
}

//员工信息
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
		//填充计算方法
		switch staff.Area {
		case "范崎路":
			staff.Calc = CalcFQ
		//代发人员计入范崎路
		case "代发工资":
			staff.Calc = CalcFQ
			staff.Area = "范崎路"
		case "外派":
			staff.Calc = CalcWP
		default:
			staff.Calc = CalcCommon
		}

		staffChan <- staff
	}

	finishChan <- FINISH_SIGNAL_STAFF

	return nil
}
//读取员工、岗位信息
func readData(staffChan chan Staff, ssChan chan SalaryStandardsTemp, spChan chan SalaryStandardsPost, finishChan chan string) error {
	file, err := xlsx.OpenFile(mConf.StaffFilePath)
	if err != nil {
		fmt.Println(err)
		finishChan <- err.Error()
		return err
	}

	for _, sheet := range file.Sheets {

		switch sheet.Name {
		case mConf.SheetNameStaff:
			err = readStaff(sheet, staffChan, finishChan)
		case mConf.SheetNameSalaryStandardsTemp:
			err = readSalaryStandardTemp(sheet, ssChan, finishChan)
		case mConf.SheetNameSalaryStandardsPost:
			err = readSalaryStandardPost(sheet, spChan, finishChan)
		}

		if err != nil {
			return err
		}
	}

	return nil
}

//单行员工信息
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
