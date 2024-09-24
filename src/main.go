package main

import (
	"fmt"
	"sync"
)

var attMap = make(map[string][]Attendance)
var staffMap = make(map[string][]Staff)
var wg sync.WaitGroup

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

	var attChan = make(chan Attendance)
	var staffChan = make(chan Staff)
	var finishChan = make(chan string)

	lockCount := len(*filePaths) //attendance
	lockCount += 1 //staff
	wg.Add(lockCount)     

	go handleChan(attChan, finishChan, staffChan, &wg, lockCount)

	for _, path := range *filePaths {
		go readFormXlsxAttendance(path, attChan, finishChan)
	}

	go readFromXlsxStaff(staffChan, finishChan)

	wg.Wait()

	fmt.Println("read all finish ~~~~~")

	// fmt.Printf(" %+v \n", attMap)
	// fmt.Println("----------------------------------")
	fmt.Printf(" %+v \n", staffMap)

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

func handleChan(attChan chan Attendance, finishChan chan string, staffChan chan Staff, wg *sync.WaitGroup, count int) {
	for {
		select {
		case att := <-attChan:
			attendances, found := attMap[att.Postion]
			if !found {
				attendances = make([]Attendance, 0)
			}

			attendances = append(attendances, att)
			attMap[att.Postion] = attendances
		case staff := <-staffChan:
			staffs, found := staffMap[staff.Local]
			if !found {
				staffs = make([]Staff, 0)
			}
			staffs = append(staffs, staff)
			staffMap[staff.Local] = staffs

		case signal := <-finishChan:

			wg.Done()
			fmt.Println("read finish ", signal)
			count--

			if count == 0 {
				return
			}
		}
	}
}
