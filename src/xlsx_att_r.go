package main

import (
	"fmt"
	"reflect"
	"strconv"

	"github.com/tealeg/xlsx/v3"
)

type Attendance struct {
	Id               int    //序号
	Name             string //姓名
	Duty             int    //应出勤
	Actal            int    //实际出勤
	Temp_8           int    //8小时临勤
	Temp_12          int    //12小时临勤；外派人员法定节假日
	Temp_4           int    //4小时临勤；加班
	Temp_Guard       int    //外派值班
	Sickness         int    //病假
	Special          int    //特殊费用
	Deduction        int    //扣款，可能是社保或罚款
	Backup           string //备注
	Postion          string //区域
	TempTransfer     int    //借调天数
	TempTransferPost string //借调岗位简称
}

func readFormXlsxAttendance(path string, c chan Attendance, finishC chan string) {
	file, err := xlsx.OpenFile(path)
	if err != nil {
		return
	}

	fmt.Println("sheet len: ", len(file.Sheets))

	for _, sheet := range file.Sheets {
		var headerMap = map[int]string{}
		fmt.Printf("mix row %d, col %d\n", sheet.MaxRow, sheet.MaxCol)
		for index := range sheet.MaxRow {
			attendance := Attendance{Id: -1}
			row, err := sheet.Row(index)
			if err != nil {
				continue
			}

			err = visitorRow(row, &attendance, &headerMap)

			if err != nil {
				continue
			}

			attendance.Postion = sheet.Name
			if attendance.Id != -1 {
				// fmt.Printf("attendance: %+v\n", attendance)
				c <- attendance
			}
		}
	}

	finishC <- path
}

func visitorRow(row *xlsx.Row, attendance *Attendance, headerMap *map[int]string) error {
	isReadHeader := len(*headerMap) == 0

	for i := 0; i < row.Sheet.MaxCol; i++ {
		str, err := row.GetCell(i).FormattedValue()
		if err != nil {
			return err
		}

		if isReadHeader {
			switch str {
			case "序号":
				(*headerMap)[i] = "Id"
			case "姓名":
				(*headerMap)[i] = "Name"
			case "应出勤":
				(*headerMap)[i] = "Duty"
			case "实出勤":
				(*headerMap)[i] = "Actal"
			case "临勤（8）":
				(*headerMap)[i] = "Temp_8"
			case "临勤（12）":
				(*headerMap)[i] = "Temp_12"
			case "临勤（4）":
				(*headerMap)[i] = "Temp_4"
			case "值班":
				(*headerMap)[i] = "Temp_Guard"
			case "病假":
				(*headerMap)[i] = "Sickness"
			case "特殊费用":
				(*headerMap)[i] = "Special"
			case "借调天数":
				(*headerMap)[i] = "TempTransfer"
			case "借调岗位":
				(*headerMap)[i] = "TempTransferPost"
			case "备注":
				(*headerMap)[i] = "Backup"
			case "扣款":
				(*headerMap)[i] = "Deduction"
			}
		} else {
			val, _ := strconv.Atoi(str)
			refType := reflect.TypeOf(*attendance)
			if refType.Kind() != reflect.Struct {
				panic("not struct")
			}

			if fieldObj, ok := refType.FieldByName((*headerMap)[i]); ok {
				if fieldObj.Type.Kind() == reflect.Int {
					reflect.ValueOf(attendance).Elem().FieldByName((*headerMap)[i]).SetInt(int64(val))

				}
				if fieldObj.Type.Kind() == reflect.String {
					reflect.ValueOf(attendance).Elem().FieldByName((*headerMap)[i]).SetString(str)

				}
			}
		}
	}
	return nil
}

// func rowVisitor(r *xlsx.Row) (err error) {
// 	err = r.ForEachCell(cellVisitor)
// 	return
// }

// func cellVisitor(c *xlsx.Cell) error {
// 	value, err := c.FormattedValue()
// 	if err!= nil {
// 		fmt.Println(err.Error())
// 	} else {
// 		fmt.Println("cell value", value)
// 	}
// 	return err
// }
