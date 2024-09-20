package main

import (
	"fmt"
)

func main() {

	readConfig()

	var filePaths *[]string = new([]string)
	err := listXlsxFile(filePaths)

	if err != nil {
		panic(err)
	}

	if len(*filePaths) == 0 {
		fmt.Println("Attendance folder is empty over")
	}

	attMap := make(map[string][]Attendance)
	for _, path := range *filePaths {
		var attendance []Attendance
		var name string
		readFormXlsxAttendance(path, &name, &attendance)
		if len(name) != 0 && len(attendance) != 0 {
			attMap[name] = attendance
		}
		fmt.Println("sheet name : ", name, " attendances ", attendance)
	}

	// wb := xlsx.NewFile()

	// sheet, err := wb.AddSheet("sheet_1")

	// if err != nil {
	// 	panic(err)
	// }

	// row := sheet.AddRow()
	// row.SetHeight(22.5)
	// cell := row.AddCell()
	// cell.Merge(11, 0)
	// cell.Value = "Merged"
	// cell.SetStyle(getStyle(STYLE_TYPE_TITLE))

	// wb.Save("..\\salaries\\merged_cells.xlsx")
}
