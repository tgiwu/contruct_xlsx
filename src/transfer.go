package main

import "slices"

type TransferInfo struct {
	Index        int
	Corporation  string
	CorpAccount  string
	StaffName    string
	StaffAccount string
	Salary       int
	Purpose      string
}

func buildTransferInfo(salaryMap map[string]map[string]Salary, staffs map[string]Staff, transferInfos *[]TransferInfo) error {
	index := 0

	keyList := make([]string, 0)
	for key := range salaryMap {
		if key != SHEET_NAME_RISK {
			keyList = append(keyList, key)
		}
	}

	slices.Sort(keyList)
	for _, key := range keyList {
		salaries := salaryMap[key]

		if len(salaries) == 0 {
			continue
		}

		for _, salary := range salaries {
			transInfo := TransferInfo{
				Index:        index + salary.Id,
				StaffAccount: staffs[salary.Name].Account,
				StaffName:    salary.Name,
				Salary:       salary.Account,
				Corporation:  mConf.CorporationName,
				CorpAccount:  mConf.CorporationAccount,
				Purpose:      mConf.SalaryPurpose,
			}
			*transferInfos = append(*transferInfos, transInfo)
		}
		index += len(salaries)

	}

	slices.SortFunc(*transferInfos, func(t1, t2 TransferInfo) int {
		return t1.Index - t2.Index
	})
	return nil
}
