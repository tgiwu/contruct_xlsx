package main

import (
	"fmt"
	"sync"

	"github.com/unidoc/unioffice/common/license"
)

var attMap = make(map[string]map[string]Attendance)
var staffMap = make(map[string]map[string]Staff)
var salaryMap = make(map[string]map[string]Salary)
var wg sync.WaitGroup

func main() {

	readConfig()
	err := license.SetMeteredKey(mConf.MeteredKey)
	if err != nil {
		panic(err)
	}

	var filePaths *[]string = new([]string)
	err = listXlsxFile(filePaths)

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
	lockCount += 1               //staff
	// lockCount += 1  
	fmt.Println("lock count ", lockCount)             //handle
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
	// fmt.Printf(" %+v \n", staffMap)

	err = buildSalaries(staffMap, attMap, &salaryMap)

	if err != nil {
		panic("build salary map failed " + err.Error())
	}

	// fmt.Printf("salary map %+v", salaryMap)

	constructXlsx(salaryMap)
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
				attendances = make(map[string]Attendance, 0)
			}

			attendances[att.Name] = att
			attMap[att.Postion] = attendances
		case staff := <-staffChan:
			staffs, found := staffMap[staff.Area]
			if !found {
				staffs = make(map[string]Staff, 0)
			}
			staffs[staff.Name] = staff
			staffMap[staff.Area] = staffs

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
