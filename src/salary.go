package main

import "fmt"

type Salary struct {
	Id             int    //序号
	Name           string //姓名
	Should         int    //应出勤
	Actual         int    //实出勤
	Standard       int    //应发工资
	NetPay         int    //实发工资
	OvertimePay    int    //加班工资
	PerformancePay int    //绩效工资 由于模板中有此项，暂时保留，值为0
	SpecialPay     int    //特殊费用
	Deduction      int    //扣款 由于模板中有此项，暂时保留，值为0
	Account        int    //合计
	BackUp         string //备注
	Postion        string //区域，用于分组
}

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

	salary.Id = attendance.Id
	salary.Name = staff.Name
	salary.Should = attendance.Duty
	salary.Actual = attendance.Actal
	salary.Standard = staff.Salary
	if attendance.Duty <= attendance.Actal {
		salary.NetPay = staff.Salary
	} else {
		salary.NetPay = staff.Salary / attendance.Duty * attendance.Actal
		salary.OvertimePay = 100 * (attendance.Actal - attendance.Duty)
	}

	if attendance.Temp_12 != 0 || attendance.Temp_8 != 0 || attendance.Temp_4 != 0 {
		salary.SpecialPay = attendance.Temp_12*180 + attendance.Temp_4*80 + attendance.Temp_8*150
	}

	if attendance.Sickness != 0 {
		salary.SpecialPay += int(float64(staff.Salary/attendance.Duty) * 0.8 * float64(attendance.Sickness))
	}

	if attendance.Special != 0 {
		salary.SpecialPay += attendance.Special
	}

	salary.Account = salary.NetPay + salary.OvertimePay + salary.PerformancePay + salary.SpecialPay - salary.Deduction
	salary.Postion = staff.Area
	return nil
}

func buildSalaries(staffs map[string]map[string]Staff, attendances map[string]map[string]Attendance, 
	salaryMap *map[string]map[string]Salary) error {
	
	for area, atts := range attendances {
		staffMap, found := staffs[area]
		if !found {
			return SalaryBuildError{msg: fmt.Sprintf("Can not find area %s in staffs!!", area)}
		}

		for name, attendance := range atts {

			if len(mConf.Ignore) != 0 {
				for _, str := range mConf.Ignore {
					if str == name {
						//ignore
						continue
					}
				}
			}

			staff, found := staffMap[name]
			if !found {
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
