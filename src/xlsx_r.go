package main

import (
	"fmt"

	"github.com/tealeg/xlsx/v3"
)

type Attendance struct {
	Id       int    `tag:"Id"`
	Name     string `tag:"Name"`
	Duty     int    `tag:"Duty"`
	Actal    int    `tag:"Actal"`
	Temp_8   int    `tag:"Temp_8"`
	Temp_12  int    `tag:"Temp_12"`
	Temp_4   int    `tag:"Temp_4"`
	Sickness int    `tag:"Sickness"`
}

func readFormXlsxAttendance(path string, name *string, attendances *[]Attendance) {
	file, err := xlsx.OpenFile(path)
	if err != nil {
		return
	}

	fmt.Println("sheet len: ", len(file.Sheets))

	for _, sheet := range file.Sheets {
		var headerMap = new(map[int]string)
		fmt.Printf("mix row %d, col %d\n", sheet.MaxRow, sheet.MaxCol)
		// sheet.ForEachRow(rowVisitor)
		for index := range sheet.MaxRow {
			var attendance *Attendance
			row, err := sheet.Row(index - 1)
			if err != nil {
				continue
			}
			visitorRow(row, attendance, headerMap)
		}
	}
}

func visitorRow(row *xlsx.Row, attendance *Attendance, headerMap *map[int]string) error {
	isReadHeader := len(*headerMap) == 0

	for i := 0; i < row.Sheet.MaxCol; i++ {
		str, err := row.GetCell(i).FormattedValue()
		if err != nil {
			continue
		}

		if isReadHeader {
			(*headerMap)[i] = str
		} else {
			if attendance == nil {

			}
		}
	}
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
