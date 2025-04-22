package main

const SALARY_SHEET_NAME_OVERVIEW = "总览"
const SALARY_OVERVIEW_COLUMN_INDEX = "序号"
const SALARY_OVERVIEW_COLUMN_AREA = "区域"
const SALARY_OVERVIEW_COLUMN_NUMBER = "发放人数"
const SALARY_OVERVIEW_COLUMN_SALARY = "总计费用"
const SALARY_OVERVIEW_COLUMN_BACKUP = "备注"

var TRANSFER_INFO_COLUNM = []string{"付款账号名称/卡名称", "付款账号/卡号", "收款账号名称", "收款账号", "金额", "汇款用途"}

var TRANSFER_INFO_COLUNM_TAG = []string{"Corporation", "CorpAccount", "StaffName", "StaffAccount", "Salary", "Purpose"}

type MyError struct {
	msg string
}

func (e MyError) Error() string {
	return e.msg
}

