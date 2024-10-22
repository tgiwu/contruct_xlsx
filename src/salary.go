package main

import (
	"fmt"
	"strconv"
	"strings"
)

type Salary struct {
	Id          int    //序号
	Name        string //姓名
	Should      int    //应出勤
	Actual      int    //实出勤
	Standard    int    //应发工资
	NetPay      int    //实发工资
	OvertimePay int    //加班工资
	SpecialPay  int    //特殊费用
	Deduction   int    //扣款,社保扣款或罚款
	Account     int    //合计
	BackUp      string //备注
	Postion     string //区域，用于分组
}

type Overview struct {
	Area         string //区域
	NumOfStaff   int    //在岗人数
	AccountTotal int    //总计费用
	BackUp       string //特殊说明
}

// 根据字数定制备注列宽度
var maxLenForBackupMap map[string]int

type SalaryBuildError struct {
	msg string
}

func (sbe SalaryBuildError) Error() string {
	return sbe.msg
}

func buildSalaryItem(staff Staff, attendance Attendance, salary *Salary) error {
	if len(staff.Name) == 0 || len(attendance.Name) == 0 {
		return SalaryBuildError{
			msg: fmt.Sprintf("staff name is empty or attendance name is empty %s, %s",
				staff.Name,
				attendance.Name),
		}
	}

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
	if attendance.Duty <= attendance.Actal {
		salary.NetPay = staff.Salary
		salary.OvertimePay = 100 * (attendance.Actal - attendance.Duty)
	} else {
		salary.NetPay = staff.Salary / attendance.Duty * attendance.Actal
	}

	if attendance.Temp_12 != 0 || attendance.Temp_8 != 0 || attendance.Temp_4 != 0 || attendance.Temp_Guard != 0 {
		salary.SpecialPay += attendance.Temp_12 * ssMap["Temp_12"]
		salary.SpecialPay += attendance.Temp_4 * ssMap["Temp_4"]
		salary.SpecialPay += attendance.Temp_8 * ssMap["Temp_8"]
		salary.SpecialPay += attendance.Temp_Guard * ssMap["Temp_Guard"]
	}

	if attendance.Sickness != 0 {
		salary.SpecialPay += int(float64(staff.Salary/attendance.Duty) * 0.8 * float64(attendance.Sickness))
	}

	if attendance.Special != 0 {
		salary.SpecialPay += attendance.Special
	}

	salary.Account = salary.NetPay + salary.OvertimePay + salary.SpecialPay - salary.Deduction
	salary.Postion = staff.Area
	salary.BackUp = attendance.Backup

	if length, found := maxLenForBackupMap[staff.Area]; found {
		if len(salary.BackUp) > length {
			maxLenForBackupMap[staff.Area] = len(salary.BackUp)
		}
	} else {
		maxLenForBackupMap[staff.Area] = len(salary.BackUp)
	}
	return nil
}

func buildSalaries(staffs map[string]map[string]Staff, attendances map[string]map[string]Attendance,
	salaryMap *map[string]map[string]Salary) error {

	for area, atts := range attendances {
		staffMap, found := staffs[area]
		if !found {
			return SalaryBuildError{msg: fmt.Sprintf("Can not find area %s in staffs!!", area)}
		}

	label:
		for name, attendance := range atts {

			if len(mConf.Ignore) != 0 {
				for _, ignore := range mConf.Ignore {
					if ignore == name {
						//ignore
						continue label
					}
				}
			}

			staff, found := staffMap[name]
			if !found {
				fmt.Printf("staff %+v", staffMap)
				return SalaryBuildError{msg: fmt.Sprintf("Can not find staff named %s in staffs!!", name)}
			}

			salary := new(Salary)
			err := buildSalaryItem(staff, attendance, salary)

			if err != nil {
				fmt.Println("build Salary item FAILED ", err.Error())
			} else {
				items, found := (*salaryMap)[attendance.Postion]
				if !found {
					items = make(map[string]Salary, 0)
				}

				items[name] = *salary

				(*salaryMap)[attendance.Postion] = items
			}
		}
	}

	return nil
}
