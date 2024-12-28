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
	SalaryTotal        //合计行
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

// 根据字数定制备注列宽度
var maxLenForBackupMap map[string]int

type SalaryBuildError struct {
	msg string
}

func (sbe SalaryBuildError) Error() string {
	return sbe.msg
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
			err := staff.Calc(&staff, &attendance, salary)

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

	if attendance.TempTransfer != 0 || len(attendance.TempTransferPost) != 0 {
		v, found := spMap[attendance.TempTransferPost]
		if found {

			if attendance.TempTransferPost == "PD100" {
				salary.SpecialPay += attendance.TempTransfer * 100
			} else {
				salary.SpecialPay += v / attendance.Duty * attendance.TempTransfer
				// fmt.Printf("%s temp transfer post is %s; during %d;transfer salary is %d\n", attendance.Name, attendance.TempTransferPost, attendance.TempTransfer, v/attendance.Duty*attendance.TempTransfer)
			}
		} else {

			salary.ErrorMap["特殊费用"] += fmt.Sprintf("未找到借调岗位 %s 对应的新进标准；", attendance.TempTransferPost)
			salary.SpecialPay = -999999 //我找到岗位
		}
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

	if attendance.Duty <= attendance.Actal {
		salary.NetPay = staff.Salary
		salary.OvertimePay += 100 * (attendance.Actal - attendance.Duty)
	} else {
		//请假按1天100算
		salary.NetPay = staff.Salary - (attendance.Duty-attendance.Actal)*100
	}
	//范崎路加班每天100
	salary.SpecialPay += attendance.Temp_4 * ssMap["Temp_Guard_Cleaner"]

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
		salary.Actual = -999999
	} else {
		if attendance.Duty < attendance.Actal {
			salary.ErrorMap["实发工资"] += fmt.Sprintf("实际出勤天数大于应出勤天数，但没有找到只算方法 应出勤 %d， 实际出勤 %d；", attendance.Duty, attendance.Actal)
			salary.NetPay = -999999 //出勤天数大于应出勤天数，需确认计算方式
		} else {
			salary.NetPay = staff.Salary / attendance.Duty * attendance.Actal
		}
		//法定节假日三倍工资,三倍以北京市最低工资2420计算，每月平均工作天数21.75天
		salary.OvertimePay += int(float64(attendance.Temp_12) * (float64(2420) / float64(21.75)) * 3)
		//值班每天 60
		salary.OvertimePay += attendance.Temp_4 * ssMap["Temp_Guard"]
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
		salary.NetPay = staff.Salary / attendance.Duty * attendance.Actal
	}
	salary.SpecialPay += attendance.Temp_12 * ssMap["Temp_12"]
	salary.SpecialPay += attendance.Temp_4 * ssMap["Temp_4"]
	salary.SpecialPay += attendance.Temp_8 * ssMap["Temp_8"]

	err = calcAfter(staff, attendance, salary)
	if err != nil {
		return err
	}

	return nil
}
