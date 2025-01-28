package main

import (
	"fmt"
	"strings"
	"sync"

	"github.com/unidoc/unioffice/common/license"
)

// area to (name to attendance)
var attMap = make(map[string]map[string]Attendance)

// name to staff
var staffMap = make(map[string]Staff)

// area to (name to salary)
var salaryMap = make(map[string]map[string]Salary)

// area to (name to salary) with risk area staff
var salaryRiskMap = make(map[string]map[string]Salary)

// temp
var ssMap = make(map[string]int)

// post
var spMap = make(map[string]int)

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
	var ssChan = make(chan SalaryStandardsTemp)
	var spChan = make(chan SalaryStandardsPost)

	lockCount := len(*filePaths)          //attendance
	lockCount += 1                        //staff
	lockCount += 1                        //salaryStandardsTemp
	lockCount += 1                        //salaryStandardsPost
	fmt.Println("lock count ", lockCount) //handle
	wg.Add(lockCount)

	go handleChan(attChan, finishChan, staffChan, ssChan, spChan, &wg, lockCount)

	for _, path := range *filePaths {
		go readFormXlsxAttendance(path, attChan, finishChan)
	}

	go readData(staffChan, ssChan, spChan, finishChan)

	wg.Wait()

	fmt.Println("read all finish ~~~~~")

	err = buildSalaries(staffMap, attMap, &salaryMap, &salaryRiskMap)

	if err != nil {
		panic("build salary map failed " + err.Error())
	}
	constructSalaryXlsx(salaryMap, "")
	s := strings.Split(mConf.FileName, ".")
	constructSalaryXlsx(salaryRiskMap, fmt.Sprintf("%s_风险人员.%s", s[0], s[1]))

	transferInfos := new([]TransferInfo)
	buildTransferInfo(salaryRiskMap, staffMap, transferInfos)
	constructTransferInfoXlsx(transferInfos, "")
	
	fmt.Println("finish")
}

func handleChan(attChan chan Attendance, finishChan chan string, staffChan chan Staff, ssChan chan SalaryStandardsTemp, spChan chan SalaryStandardsPost, wg *sync.WaitGroup, count int) {
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
			staffMap[staff.Name] = staff
		case ss := <-ssChan:
			ssMap[ss.TempType] = ss.SalaryPerDay
		case sp := <-spChan:
			spMap[sp.PostType] = sp.SalaryPerMonth
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
