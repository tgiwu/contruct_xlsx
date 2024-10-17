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
	fmt.Println("lock count ", lockCount) //handle
	wg.Add(lockCount)

	go handleChan(attChan, finishChan, staffChan, &wg, lockCount)

	for _, path := range *filePaths {
		go readFormXlsxAttendance(path, attChan, finishChan)
	}

	go readFromXlsxStaff(staffChan, finishChan)

	wg.Wait()

	fmt.Println("read all finish ~~~~~")

	err = buildSalaries(staffMap, attMap, &salaryMap)

	if err != nil {
		panic("build salary map failed " + err.Error())
	}


	constructXlsx(salaryMap)

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
