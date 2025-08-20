package main

import (
	"fmt"
	"math"
	"slices"
	"strconv"
	"strings"
	"sync"
)

type Salary struct {
	Id             int    //序号
	Name           string //姓名
	Should         int    //应出勤
	Actual         int    //实出勤
	Standard       int    //应发工资
	NetPay         int    //实发工资
	OvertimePay    int    //加班工资
	SpecialPay     int    //特殊费用
	Deduction      int    //扣款,社保扣款或罚款
	Account        int    //合计（不再用于展示）
	AccountFormula string //合计公式
	BackUp         string //备注
	Area           string //区域，用于分组
	SalaryTotal           //合计行
	//todo: 写入表之前不能确定具体位置，暂时只能绑定列名
	ErrorMap map[string]string //错误批注，如有错误单元格标红并添加批注;key:列名；value：错误描述
}

type SalaryTotal struct {
	totalStandard string //应发合计（合计）
	totalNetPay   string //实发合计
	totalAccount  string //共计
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

var salaryNextIdMap = make(map[string]int, 0)

// 工资计算错误，定义负值，便于在excel中标记
const ERROR_SALARY = -99999

type SalaryBuildError struct {
	msg string
}

func (sbe SalaryBuildError) Error() string {
	return sbe.msg
}

func (obj *OverviewItems) addItems(item Overview) {
	for {
		if obj.lock.TryLock() {
			obj.overviewArr = append(obj.overviewArr, item)
			obj.lock.Unlock()
			break
		}
	}
}

func (obj *OverviewItems) items() []Overview {
	return obj.overviewArr
}

func buildSalaries(staffs map[string]Staff, attendances map[string][]Attendance,
	salaryMap *map[string]map[string]Salary, salaryRiskMap *map[string]map[string]Salary) error {

	keys := make([]string, 0, len(attendances))
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
				fmt.Printf("staff %+v", staffMap)
				return SalaryBuildError{msg: fmt.Sprintf("Can not find staff named %s in staffs!!", attendance.Name)}
			}

			salary := new(Salary)
			err := staff.Calc(&staff, &attendance, salary)

			//do not differential risk
			if err != nil {
				fmt.Println("build Salary item FAILED ", err.Error())
			} else {
				items, found := (*salaryMap)[attendance.Postion]
				if !found {
					items = make(map[string]Salary, 0)
				}

				items[attendance.Name] = *salary

				(*salaryMap)[attendance.Postion] = items
			}

			sumStart, sumEnd, deduction := 0, 0, 0
			for i, s := range mConf.Headers {
				switch s {
				case "实发工资":
					sumStart = i
				case "特殊费用":
					sumEnd = i
				case "扣款":
					deduction = i
				}
			}

			salaryCopy := new(Salary)
			DeepCopy(salary, salaryCopy)
			//differential risk
			if !staff.RiskIgnore && ((staff.Age < 60 && staff.Sex == 1) || (staff.Age < 50 && staff.Sex == 0)) {
				items, found := (*salaryRiskMap)[SHEET_NAME_RISK]
				if !found {
					items = make(map[string]Salary, 0)
					salaryNextIdMap[SHEET_NAME_RISK] = 1
				}
				salaryCopy.Id = salaryNextIdMap[SHEET_NAME_RISK]
				//recalc account formula
				salaryCopy.AccountFormula = fmt.Sprintf("=SUM(%s:%s) - %s", pos(salaryCopy.Id+1, sumStart),
					pos(salaryCopy.Id+1, sumEnd), pos(salaryCopy.Id+1, deduction))
				items[attendance.Name] = *salaryCopy

				(*salaryRiskMap)[SHEET_NAME_RISK] = items
				salaryNextIdMap[SHEET_NAME_RISK]++
			} else {
				items, found := (*salaryRiskMap)[SHEET_NAME_NO_RISK]
				if !found {
					items = make(map[string]Salary, 0)
					salaryNextIdMap[SHEET_NAME_NO_RISK] = 1
				}
				salaryCopy.Id = salaryNextIdMap[SHEET_NAME_NO_RISK]
				//recalc account formula
				salaryCopy.AccountFormula = fmt.Sprintf("=SUM(%s:%s) - %s", pos(salaryCopy.Id+1, sumStart),
					pos(salaryCopy.Id+1, sumEnd), pos(salaryCopy.Id+1, deduction))

				items[attendance.Name] = *salaryCopy

				(*salaryRiskMap)[SHEET_NAME_NO_RISK] = items
				salaryNextIdMap[SHEET_NAME_NO_RISK]++
			}
		}
	}

	return nil
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

	sumStart, sumEnd, deduction := 0, 0, 0
	for i, s := range mConf.Headers {
		switch s {
		case "实发工资":
			sumStart = i
		case "特殊费用":
			sumEnd = i
		case "扣款":
			deduction = i
		}
	}

	salary.AccountFormula = fmt.Sprintf("=SUM(%s:%s) - %s", pos(salary.Id+1, sumStart), pos(salary.Id+1, sumEnd), pos(salary.Id+1, deduction))
	salary.Area = staff.Area
	salary.BackUp = attendance.Backup

	if length, found := maxLenForBackupMap[staff.Area]; found {
		if len(salary.BackUp) > length {
			maxLenForBackupMap[staff.Area] = len(salary.BackUp)
		}
	} else {
		maxLenForBackupMap[staff.Area] = len(salary.BackUp)
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
	//入职月工作天数不满月，每天100
	if strings.Contains(attendance.Backup, "入职") && attendance.Actal < attendance.Duty {
		//试行逻辑，领导说不清楚怎么发--！，月薪超过3000按日平均工资计算首月工资；未超过3000的按日薪100计算
		if staff.Salary > 3000 {
			salary.NetPay = staff.Salary / attendance.Duty * attendance.Actal
		} else {
			salary.NetPay = 100 * attendance.Actal
		}
		//超出应出勤的天数每天100
	} else if attendance.Duty <= attendance.Actal {
		salary.NetPay = staff.Salary
		salary.OvertimePay += 100 * (attendance.Actal - attendance.Duty)
	} else {
		//请假按1天100算
		salary.NetPay = staff.Salary - (attendance.Duty-attendance.Actal)*100
	}
	//范崎路加班每天100
	salary.SpecialPay += int(attendance.Temp_4 * float64(ssMap["Temp_Guard_Cleaner"]))

	err = calcAfter(staff, attendance, salary)
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

	if err != nil {
		return err
	}

	return nil
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
