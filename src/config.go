package main

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"runtime"

	"github.com/spf13/viper"
)

const CONFIG_PATH = "config/"

type config struct {
	AttendanceFolder             string   `mapstructure:"attendance_folder"`
	StaffFilePath                string   `mapstructure:"staff_file_path"`
	Ignore                       []string `mapstructure:"ignore"`
	OutputPath                   string   `mapstructure:"output_path"`
	FileName                     string   `mapstructure:"file_name"`
	FileTransferName             string   `mapstructure:"file_transfer_name"`
	Month                        int      `mapstructure:"month"`
	Year                         int      `mapstructure:"year"`
	Headers                      []string `mapstructure:"headers"`
	HeadersRisk                  []string `mapstructure:"headers_risk"`
	HeadersRiskMap               map[string]string
	HeadersMap                   map[string]string
	MeteredKey                   string   `mapstructure:"metered_key"`
	CorporationName              string   `mapstructure:"corporation_name"`
	CorporationAccount           string   `mapstructure:"corporation_account"`
	SalaryPurpose                string   `mapstructure:"salary_purpose"`
	OverviewHeader               []string `mapstructure:"overview_header"`
	OverviewHeaderMap            map[string]string
	SheetNameStaff               string   `mapstructure:"staff_sheet_name"`
	SheetNameSalaryStandardsTemp string   `mapstructure:"salary_standards_temp_sheet_name"`
	SheetNameSalaryStandardsPost string   `mapstructure:"salary_standards_post_sheet_name"`
	ConstructTransferFile        bool     `mapstructure:"construct_transfer_file"`
	AreaSortArray                []string `mapstructure:"area_sort_array"`
}

var mConf config

func readConfig() {

	vip := viper.New()
	vip.AddConfigPath(CONFIG_PATH)
	vip.SetConfigName("config_common.yaml")

	vip.SetConfigType("yaml")

	if err := vip.ReadInConfig(); err != nil {
		if _, ok := err.(viper.ConfigFileNotFoundError); ok {
			fmt.Println("not find " + err.Error())
		} else {
			fmt.Printf("read config file err, %v \n", err)
		}
	}

	configName := "config"
	switch runtime.GOOS {
	case "windows":
		configName += "_win"
	case "linux":
		configName += "_lin"
	case "darwin":
		configName += "_lin"
	default:
		fmt.Println("unsupport os ", runtime.GOOS)
	}

	bs, err := os.ReadFile(filepath.Join(CONFIG_PATH, configName+".yaml"))

	vip.MergeConfig(bytes.NewReader(bs))

	if err != nil {
		panic(err)
	}

	err = vip.Unmarshal(&mConf)
	if err != nil {
		panic(err)
	}

	if len(mConf.Headers) != 0 {
		headersMap := new(map[string]string)
		analysisHeader(mConf.Headers, headersMap)
		mConf.HeadersMap = *headersMap
	}

	if len(mConf.OverviewHeader) != 0 {

		overviewHeaderMap := new(map[string]string)
		analysisHeader(mConf.OverviewHeader, overviewHeaderMap)
		mConf.OverviewHeaderMap = *overviewHeaderMap

	}

	if len(mConf.HeadersRisk) != 0 {
		headerRiskMap := new(map[string]string)
		analysisHeader(mConf.HeadersRisk, headerRiskMap)
		mConf.HeadersRiskMap = *headerRiskMap
	}

	fmt.Printf("config %+v \n", mConf)
}

func analysisHeader(list []string, resultMap *map[string]string) error {

	if (*resultMap) == nil {
		*resultMap = make(map[string]string, len(list))
	}

	if len(list) != 0 {

		for _, header := range list {
			switch header {
			case "序号":
				(*resultMap)[header] = "Id"
			case "姓名":
				(*resultMap)[header] = "Name"
			case "应出勤":
				(*resultMap)[header] = "Should"
			case "实际出勤":
				(*resultMap)[header] = "Actual"
			case "应发工资":
				(*resultMap)[header] = "Standard"
			case "实发工资":
				(*resultMap)[header] = "NetPay"
			case "加班工资":
				(*resultMap)[header] = "OvertimePay"
			case "特殊费用":
				(*resultMap)[header] = "SpecialPay"
			case "扣款":
				(*resultMap)[header] = "Deduction"
			case "合计":
				(*resultMap)[header] = "AccountFormula"
			case "备注":
				(*resultMap)[header] = "BackUp"
			case "区域":
				(*resultMap)[header] = "Area"
			case "发放人数":
				(*resultMap)[header] = "NumOfStaff"
			case "总计费用":
				(*resultMap)[header] = "AccountTotal"
			default:
				return MyError{fmt.Sprintf("UNKNOWN HEADER named %s \n", header)}

			}
		}
	}
	return nil
}
