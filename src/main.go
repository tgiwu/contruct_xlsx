package main

import (
	"fmt"
	"strings"
	"sync"

	"github.com/unidoc/unioffice/common/license"

	log "github.com/sirupsen/logrus"
)

// area to (name to attendance)
var attMap = make(map[string][]Attendance)

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
		log.Infoln("Attendance folder is empty over")
	}

	var attChan = make(chan Attendance)
	var staffChan = make(chan Staff)
	var finishChan = make(chan string)
	var ssChan = make(chan SalaryStandardsTemp)
	var spChan = make(chan SalaryStandardsPost)

	lockCount := len(*filePaths) //attendance
	lockCount += 1               //staff
	lockCount += 1               //salaryStandardsTemp
	lockCount += 1               //salaryStandardsPost
	// fmt.Println("lock count ", lockCount) //handle

	var wg sync.WaitGroup

	wg.Add(lockCount)
	//start read handler
	go handleChan(attChan, finishChan, staffChan, ssChan, spChan, &wg, lockCount)

	for _, path := range *filePaths {
		go readFormXlsxAttendance(path, attChan, finishChan)
	}

	go readData(staffChan, ssChan, spChan, finishChan)

	wg.Wait()

	log.Infoln("--------------------read finish--------------------")

	err = buildSalaries(staffMap, attMap, &salaryMap, &salaryRiskMap)

	if err != nil {
		log.Panic(" build salary map failed " + err.Error())
	}

	//salary excel
	var xlsxWG sync.WaitGroup
	xlsxC := make(chan string)

	xlsxCount := 0
	xlsxCount += 1 //salsry separate by area
	xlsxCount += 1 //salary separate by risk
	xlsxCount += 1 //transfer info for no risk

	//start xlsx construction handler
	go handleXLSXSignal(xlsxC, &xlsxWG, xlsxCount)
	xlsxWG.Add(xlsxCount)

	//construct salary xlsx normal
	go constructSalaryXlsx(salaryMap, mConf.FileName, xlsxC, false)

	//construct salary xlsx separate by risk
	go func() {
		s := strings.Split(mConf.FileName, ".")
		constructSalaryXlsx(salaryRiskMap, fmt.Sprintf("%s_风险人员.%s", s[0], s[1]), xlsxC, true)
	}()

	//construct transfer information xlsx for no risk
	go func() {
		transferInfos := new([]TransferInfo)
		buildTransferInfo(salaryRiskMap, staffMap, transferInfos)
		constructTransferInfoXlsx(transferInfos, mConf.FileTransferName, xlsxC)
	}()
	xlsxWG.Wait()

	log.Infoln("all finish")
}

func handleChan(attChan chan Attendance, finishChan chan string, staffChan chan Staff, ssChan chan SalaryStandardsTemp, spChan chan SalaryStandardsPost, wg *sync.WaitGroup, count int) {
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
			staffMap[staff.Name] = staff
		case ss := <-ssChan:
			ssMap[ss.TempType] = ss.SalaryPerDay
		case sp := <-spChan:
			spMap[sp.PostType] = sp.SalaryPerMonth
		case signal := <-finishChan:

			wg.Done()
			log.Info("read finish ", signal)
			count--

			if count == 0 {
				return
			}
		}
	}
}

func handleXLSXSignal(c chan string, xlsxWg *sync.WaitGroup, count int) {

	for {
		s := <-c
		log.Infoln(s)
		xlsxWg.Done()
		count--
		if count == 0 {
			return
		}
	}
}
