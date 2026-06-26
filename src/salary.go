package main

import (
	"fmt"
	"math"
	"slices"
	"strconv"
	"strings"
	"sync"
	// log "github.com/sirupsen/logrus"
)

type Salary struct {
	Id             int    //序号
	StaffId        int    //staff Id for sort
	Name           string //姓名
	Should         int    //应出勤
	Actual         int    //实出勤
	Standard       int    //应发工资
	NetPay         int    //实发工资
	OvertimePay    int    //加班工资
	SpecialPay     int    //特殊费用
	Deduction      int    //扣款,社保扣款或罚款
	Account        int    //合计，仅用于总览页显示
	AccountFormula string //合计公式
	BackUp         string //备注
	Area           string //区域，用于分组
	SalaryTotal           //合计行
	//todo: 写入表之前不能确定具体位置，暂时只能绑定列名
	ErrorMap map[string]string //错误批注，如有错误单元格标红并添加批注;key:列名；value：错误描述
}

type SalaryTotal struct {
	TotalStandard string //应发合计（合计）Formula
	TotalNetPay   string //实发合计 Formula
	TotalAccount  string //共计 Formula
	TotalPerson   string //person sum Formula
	TotalP        int    //person sum
	TotalA        int    //account sum
}

type Overview struct {
	Area         string //区域
	NumOfStaff   int    //在岗人数
	AccountTotal int    //总计费用
	BackUp       string //特殊说明
}

// 临勤工资标准
type SalaryStandardsTemp struct {
	TempType     string //临勤类型
	SalaryPerDay int    //日薪
	Description  string //说明
}

// 借调工资标准
type SalaryStandardsPost struct {
	PostType       string //岗位类型
	SalaryPerMonth int    //月薪
	Description    string //描述
}

type OverviewItems struct {
	lock        sync.Mutex
	overviewArr []Overview
}

// 根据字数定制备注列宽度
var maxLenForBackupMap map[string]int

// 工资计算错误，定义负值，便于在excel中标记
const ERROR_SALARY = -99999

type SalaryBuildError struct {
	msg string
}

func (sbe SalaryBuildError) Error() string {
	return sbe.msg
}

func buildSalaries2(staffs map[string]Staff, attendances map[string][]Attendance,
	salaryMap *map[string][]Salary, salaryRiskMap *map[string][]Salary) error {

	keys := make([]string, 0, len(attendances))

	//risk table overview items,first risk, second no risk
	overviewAreaRows := make([]Salary, 0)
	//cache salary account by area
	areaToSumRowMap := make(map[string]Salary, len(attendances))

	//table area construct sum Formula
	// indexStandard, indexNetPay, indexAccount := 0, 0, 0
	indexStandard := slices.Index(mConf.Headers, "应发工资")
	indexNetPay := slices.Index(mConf.Headers, "实发工资")
	indexAccount := slices.Index(mConf.Headers, "合计")
	// for i, s := range mConf.Headers {
	// 	switch s {
	// 	case "应发工资":
	// 		indexStandard = i
	// 	case "实发工资":
	// 		indexNetPay = i
	// 	case "合计":
	// 		indexAccount = i
	// 	}
	// }

	for k := range attendances {
		keys = append(keys, k)
	}

	if len(mConf.AreaSortArray) > 0 {
		areaIndexMap := make(map[string]int, len(mConf.AreaSortArray))
		for i, area := range mConf.AreaSortArray {
			areaIndexMap[area] = i
		}

		slices.SortFunc(keys, func(x string, y string) int {
			xi, yi := -1, -1

			if i, found := areaIndexMap[x]; found {
				xi = i
			}

			if i, found := areaIndexMap[y]; found {
				yi = i
			}
			return xi - yi
		})
	}
	for _, key := range keys {

		for _, attendance := range attendances[key] {

			if len(attendance.Name) == 0 {
				continue
			}
			if len(mConf.Ignore) != 0 {
				for _, ignore := range mConf.Ignore {
					if ignore == attendance.Name {
						//ignore
						continue
					}
				}
			}
			staff, found := staffs[attendance.Name]

			if !found {
				return SalaryBuildError{msg: fmt.Sprintf("Can not find staff named %s in staffs!!", attendance.Name)}
			}

			salary := new(Salary{StaffId: staff.RowId})
			err := staff.Calc(&staff, &attendance, salary)

			//do not differential risk
			if err != nil {
				fmt.Println("build Salary item FAILED ", err.Error())
			} else {
				items, found := (*salaryMap)[attendance.Postion]
				if !found {
					items = make([]Salary, 0)
				}

				items = append(items, *salary)

				(*salaryMap)[attendance.Postion] = items

				var sumRow Salary
				if item, found := areaToSumRowMap[attendance.Postion]; found {
					sumRow = item
				} else {
					sumRow = Salary{
						Area:    attendance.Postion,
						Name:    "合计",
						StaffId: 999,
					}
				}

				sumRow.TotalA += salary.Account
				sumRow.TotalP += 1

				areaToSumRowMap[attendance.Postion] = sumRow
				//get backup max length
				if length, found := maxLenForBackupMap[staff.Area]; found {
					if len(salary.BackUp) > length {
						maxLenForBackupMap[staff.Area] = len(salary.BackUp)
					}
				} else {
					maxLenForBackupMap[staff.Area] = len(salary.BackUp)
				}
			}

			salaryCopy := new(Salary)
			DeepCopy(salary, salaryCopy)
			//differential risk
			if strings.HasPrefix(staff.Account, "00000") ||
				(!staff.RiskIgnore && ((staff.Age < 60 && staff.Sex == 1) || (staff.Age < 50 && staff.Sex == 0))) {
				items, found := (*salaryRiskMap)[SHEET_NAME_RISK]
				if !found {
					items = make([]Salary, 0)
				}
				salaryCopy.Id = staff.RowId

				items = append(items, *salaryCopy)

				(*salaryRiskMap)[SHEET_NAME_RISK] = items
				//get backup max length
				if length, found := maxLenForBackupMap[SHEET_NAME_RISK]; found {
					if len(salary.BackUp) > length {
						maxLenForBackupMap[SHEET_NAME_RISK] = len(salary.BackUp)
					}
				} else {
					maxLenForBackupMap[SHEET_NAME_RISK] = len(salary.BackUp)
				}
			} else {
				items, found := (*salaryRiskMap)[SHEET_NAME_NO_RISK]
				if !found {
					items = make([]Salary, 0)
				}
				salaryCopy.Id = staff.RowId

				items = append(items, *salaryCopy)

				(*salaryRiskMap)[SHEET_NAME_NO_RISK] = items
				//get backup max length
				if length, found := maxLenForBackupMap[SHEET_NAME_NO_RISK]; found {
					if len(salary.BackUp) > length {
						maxLenForBackupMap[SHEET_NAME_NO_RISK] = len(salary.BackUp)
					}
				} else {
					maxLenForBackupMap[SHEET_NAME_NO_RISK] = len(salary.BackUp)
				}
			}
		}

	}

	for key := range *salaryMap {
		list := (*salaryMap)[key]
		slices.SortFunc(list, func(s1, s2 Salary) int {
			return s1.StaffId - s2.StaffId
		})
		recalcAccountRowFormula(&list, 2)
		sumRow := areaToSumRowMap[key]
		sumRow.TotalStandard = fmt.Sprintf("=SUM(%s:%s)", pos(2, indexStandard), pos(len(list)+1, indexStandard))
		sumRow.TotalNetPay = fmt.Sprintf("=SUM(%s:%s)", pos(2, indexNetPay), pos(len(list)+1, indexNetPay))
		sumRow.TotalAccount = fmt.Sprintf("=SUM(%s:%s)", pos(2, indexAccount), pos(len(list)+1, indexAccount))

		list = append(list, sumRow)

		(*salaryMap)[key] = list

	}

	for key := range *salaryRiskMap {
		list := (*salaryRiskMap)[key]
		slices.SortFunc(list, func(s1, s2 Salary) int {
			return s1.StaffId - s2.StaffId
		})

		recalcAccountRowFormula(&list, 2)
		riskSumRow := Salary{StaffId: 999,
			Name: "合计",
		}

		riskSumRow.TotalStandard = fmt.Sprintf("=SUM(%s:%s)", pos(2, indexStandard), pos(len(list)+1, indexStandard))
		riskSumRow.TotalNetPay = fmt.Sprintf("=SUM(%s:%s)", pos(2, indexNetPay), pos(len(list)+1, indexNetPay))
		riskSumRow.TotalAccount = fmt.Sprintf("=SUM(%s:%s)", pos(2, indexAccount), pos(len(list)+1, indexAccount))
		riskSumRow.TotalP = len(list)
		list = append((*salaryRiskMap)[key], riskSumRow)

		(*salaryRiskMap)[key] = list
	}

	for _, key := range keys {
		row := areaToSumRowMap[key]
		row.StaffId = -1
		overviewAreaRows = append(overviewAreaRows, row)
	}

	overviewSumRow := Salary{
		Name:    "合计",
		Area:    "合计",
		StaffId: 999,
	}
	// person sum Formula
	// totalAccountIndex, totalPersonIndex := 0, 0
	totalAccountIndex := slices.Index(mConf.OverviewHeader, "总计费用")
	totalPersonIndex := slices.Index(mConf.OverviewHeader, "发放人数")
	// for i, s := range mConf.OverviewHeader {
	// 	switch s {
	// 	case "总计费用":
	// 		totalAccountIndex = i
	// 	case "发放人数":
	// 		totalPersonIndex = i
	// 	}
	// }

	overviewSumRow.TotalAccount = fmt.Sprintf("=SUM(%s:%s)", pos(2, totalAccountIndex), pos(len(overviewAreaRows)+1, totalAccountIndex))
	overviewSumRow.TotalPerson = fmt.Sprintf("=SUM(%s:%s)", pos(2, totalPersonIndex), pos(len(overviewAreaRows)+1, totalPersonIndex))
	overviewAreaRows = append(overviewAreaRows, overviewSumRow)

	// add account row to overview
	(*salaryRiskMap)["总览"] = overviewAreaRows

	return nil
}

// after sort recalc account formula
func recalcAccountRowFormula(items *[]Salary, start int) {

	sumStart := slices.Index(mConf.Headers, "实发工资")
	sumEnd := slices.Index(mConf.Headers, "特殊费用")
	deduction := slices.Index(mConf.Headers, "扣款")

	// sumStart, sumEnd, deduction := 0, 0, 0
	// for i, s := range mConf.Headers {
	// 	switch s {
	// 	case "实发工资":
	// 		sumStart = i
	// 	case "特殊费用":
	// 		sumEnd = i
	// 	case "扣款":
	// 		deduction = i
	// 	}
	// }

	for i := range *items {
		(*items)[i].AccountFormula = fmt.Sprintf("=SUM(%s:%s) - %s", pos(i+start, sumStart),
			pos(i+start, sumEnd), pos(i+start, deduction))
	}
}

func calcBefore(staff *Staff, attendance *Attendance, salary *Salary) error {
	if len(staff.Name) == 0 || len(attendance.Name) == 0 {
		return SalaryBuildError{
			msg: fmt.Sprintf("staff name is empty or attendance name is empty %s, %s",
				staff.Name,
				attendance.Name),
		}
	}

	salary.ErrorMap = make(map[string]string)

	if maxLenForBackupMap == nil {
		maxLenForBackupMap = make(map[string]int)
	}

	salary.Id = attendance.Id
	salary.Name = staff.Name
	salary.Should = attendance.Duty
	salary.Actual = attendance.Actal
	if len(staff.BackUp.BackUpSal) != 0 {
		// 切面中文字符在utf-8下占3字节
		salMonthStr := attendance.Backup[(strings.Index(attendance.Backup, "发") + 3):strings.Index(attendance.Backup, "月")]
		salMonth, _ := strconv.Atoi(salMonthStr)
	salRuleLabel:
		for _, salRule := range staff.BackUp.BackUpSal {
			for _, month := range salRule.Month {
				if month == salMonth {
					staff.Salary = salRule.Sal
					break salRuleLabel
				}
			}
		}
	}
	salary.Standard = staff.Salary
	salary.Deduction = attendance.Deduction
	return nil
}

func calcAfter(staff *Staff, attendance *Attendance, salary *Salary) error {

	//处理借调
	if attendance.TempTransfer != 0 || len(attendance.TempTransferPost) != 0 {
		v, found := spMap[attendance.TempTransferPost]
		if found {

			if attendance.TempTransferPost == "PD100" {
				salary.SpecialPay += attendance.TempTransfer * 100
			} else {
				salary.SpecialPay += v / attendance.Duty * attendance.TempTransfer
			}
		} else {

			salary.ErrorMap["特殊费用"] += fmt.Sprintf("未找到借调岗位 %s 对应的新进标准；", attendance.TempTransferPost)
			salary.SpecialPay = -999999 //我找到岗位
		}
	}
	//病假
	if attendance.Sickness != 0 {
		salary.SpecialPay += int(float64(staff.Salary/attendance.Duty) * 0.8 * float64(attendance.Sickness))
	}
	//特殊费用
	if attendance.Special != 0 {
		salary.SpecialPay += attendance.Special
	}

	//已替换为公式
	salary.Account = salary.NetPay + salary.OvertimePay + salary.SpecialPay - salary.Deduction

	sumStart := slices.Index(mConf.Headers, "实发工资")
	sumEnd := slices.Index(mConf.Headers, "特殊费用")
	deduction := slices.Index(mConf.Headers, "扣款")
	// sumStart, sumEnd, deduction := 0, 0, 0
	// for i, s := range mConf.Headers {
	// 	switch s {
	// 	case "实发工资":
	// 		sumStart = i
	// 	case "特殊费用":
	// 		sumEnd = i
	// 	case "扣款":
	// 		deduction = i
	// 	}
	// }

	salary.AccountFormula = fmt.Sprintf("=SUM(%s:%s) - %s", pos(salary.Id+1, sumStart), pos(salary.Id+1, sumEnd), pos(salary.Id+1, deduction))
	salary.Area = staff.Area
	salary.BackUp = attendance.Backup
	if strings.HasPrefix(staff.Account, "00000") {
		if len(salary.BackUp) != 0 {
			salary.BackUp += ";"
		}
		salary.BackUp += "未提供工资卡"
	}

	staff.Sal = salary
	staff.Att = attendance
	return nil
}

// 范崎路
func CalcFQ(staff *Staff, attendance *Attendance, salary *Salary) error {
	err := calcBefore(staff, attendance, salary)

	if err != nil {
		return err
	}

	switch {
	//入职月工作天数不满月，每天100
	case strings.Contains(attendance.Backup, "入职") && attendance.Actal < attendance.Duty:
		//试行逻辑，领导说不清楚怎么发--！，月薪超过3000按日平均工资计算首月工资；未超过3000的按日薪100计算
		if staff.Salary > 3000 {
			salary.NetPay = staff.Salary / attendance.Duty * attendance.Actal
		} else {
			salary.NetPay = 100 * attendance.Actal
		}
		//超出应出勤的天数每天100
	// case attendance.Duty == attendance.Actal:
	// 	salary.NetPay = staff.Salary + 200
	// salary.OvertimePay += 100 * (attendance.Actal - attendance.Duty)
	case attendance.Actal >= attendance.Duty:
		salary.NetPay = staff.Salary + 100*(attendance.Actal-attendance.Duty)
		//工作天数少于15天时按日工资结算
	case attendance.Actal < 15:
		salary.NetPay = staff.Salary / attendance.Duty * attendance.Actal
	default:
		salary.NetPay = staff.Salary / attendance.Duty * attendance.Actal

	}
	//范崎路加班每天100
	salary.SpecialPay += int(attendance.Temp_4 * float64(ssMap["Temp_Guard_Cleaner"]))

	err = calcAfter(staff, attendance, salary)
	//modify should and standard
	//hard code very ugly
	if salary.Area == "范崎路" && salary.Name != "张阳" {
		salary.Should = LastDayInMonth(mConf.Year, mConf.Month)
		salary.Standard = staff.Salary + 200
	}
	if err != nil {
		return err
	}

	return nil
}

// 外派
func CalcWP(staff *Staff, attendance *Attendance, salary *Salary) error {

	err := calcBefore(staff, attendance, salary)

	if err != nil {
		return err
	}
	//实际出勤天数包含法定节假日工作天数
	//若法定节假日天数多余实际出勤天数，考勤错误
	if salary.Actual-attendance.Temp_12 < 0 {
		salary.ErrorMap["实际出勤"] += fmt.Sprintf("法定节假日数多余实际出勤天数 实际出勤 %d，法定节假日 %d", attendance.Actal, attendance.Temp_12)
		salary.Actual = ERROR_SALARY
	} else {
		switch {
		case attendance.Duty < attendance.Actal:
			salary.ErrorMap["实发工资"] += fmt.Sprintf("实际出勤天数大于应出勤天数，但没有找到只算方法 应出勤 %d， 实际出勤 %d；", attendance.Duty, attendance.Actal)
			salary.NetPay = ERROR_SALARY //出勤天数大于应出勤天数，需确认计算方式
		case attendance.Duty == attendance.Actal:
			salary.NetPay = staff.Salary
		default:
			salary.NetPay = int(math.Round(float64(staff.Salary) / float64(attendance.Duty) * float64(attendance.Actal)))
		}

		//法定节假日三倍工资,每天基本工资80，3倍240
		salary.OvertimePay += int(attendance.Temp_12 * ssMap["Temp_Guard_Holiday"])
		//值班每天 60
		salary.OvertimePay += int(attendance.Temp_4 * float64(ssMap["Temp_Guard"]))
	}

	err = calcAfter(staff, attendance, salary)

	return err

}

// 梦想山
func CalcDM(staff *Staff, attendance *Attendance, salary *Salary) error {
	err := calcBefore(staff, attendance, salary)
	if err != nil {
		return err
	}

	diff := attendance.Duty - attendance.Actal

	switch {
	//正常工作，有请假
	case diff > 0:
		salary.NetPay = int(math.Round(float64(staff.Salary) / float64(attendance.Duty) * float64(attendance.Actal)))
	//实际工作时间多余应到岗天数，按当月天数折算日工资，算作加班
	case diff < 0:
		salary.NetPay = staff.Salary
		salary.OvertimePay += -(int(math.Round(float64(staff.Salary) / float64(attendance.Duty) * float64(diff))))
	//全勤
	case diff == 0:
		salary.NetPay = staff.Salary
	}

	err = calcAfter(staff, attendance, salary)
	return err
}

// 通用
func CalcCommon(staff *Staff, attendance *Attendance, salary *Salary) error {

	err := calcBefore(staff, attendance, salary)

	if err != nil {
		return err
	}

	if attendance.Duty < attendance.Actal {
		salary.ErrorMap["实发工资"] += fmt.Sprintf("实际出勤天数大于应出勤天数，但没有找到只算方法 应出勤 %d， 实际出勤 %d；", attendance.Duty, attendance.Actal)
		salary.NetPay = -999999 //出勤天数大于应出勤天数，需确认计算方式
	} else if attendance.Duty == attendance.Actal {
		salary.NetPay = staff.Salary
	} else {
		salary.NetPay = int(math.Round(float64(staff.Salary) / float64(attendance.Duty) * float64(attendance.Actal)))
	}
	salary.SpecialPay += attendance.Temp_12 * ssMap["Temp_12"]
	salary.SpecialPay += int(math.Round(attendance.Temp_4 * float64(ssMap["Temp_4"])))
	salary.SpecialPay += attendance.Temp_8 * ssMap["Temp_8"]

	err = calcAfter(staff, attendance, salary)
	if err != nil {
		return err
	}

	return nil
}
