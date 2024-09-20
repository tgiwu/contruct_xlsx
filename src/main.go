package main

import (
	"fmt"
)

func main() {

	readConfig()

	var files *[]string = new([]string)
	err := listXlsxFile(files)

	if err != nil {
		panic(err)
	}

	if len(*files) == 0 {
		fmt.Println("Attendance folder is empty over")
	}

	for _, path := range *files {
		var attendance *[]Attendance
		var name string
		readFormXlsxAttendance(path, &name, attendance)

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
