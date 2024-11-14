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
	Month                        int      `mapstructure:"month"`
	Year                         int      `mapstructure:"year"`
	Headers                      []string `mapstructure:"headers"`
	HeadersMap                   map[string]string
	MeteredKey                   string   `mapstructure:"metered_key"`
	CorporationName              string   `mapstructure:"corporation_name"`
	CorporationAccount           string   `mapstructure:"corporation_account"`
	OverviewHeader               []string `mapstructure:"overview_header"`
	OverviewHeaderMap            map[string]string
	SheetNameStaff               string `mapstructure:"staff_sheet_name"`
	SheetNameSalaryStandardsTemp string `mapstructure:"salary_standards_temp_sheet_name"`
	SheetNameSalaryStandardsPost string `mapstructure:"salary_standards_post_sheet_name"`
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
		headersMap := make(map[string]string, len(mConf.Headers))

		//Id             int    //序号
		// Name           string //姓名
		// Should         int    //应出勤
		// Actual         int    //实出勤
		// Standard       int    //应发工资
		// NetPay         int    //实发工资
		// OvertimePay    int    //加班工资
		// PerformancePay int    //绩效工资 由于模板中有此项，暂时保留，值为0
		// SpecialPay     int    //特殊费用
		// Deduction      int    //扣款 由于模板中有此项，暂时保留，值为0
		// Account        int    //合计
		// BackUp         string //备注
		// Postion        string //区域，用于分组

		for _, header := range mConf.Headers {
			switch header {
			case "序号":
				headersMap[header] = "Id"
			case "姓名":
				headersMap[header] = "Name"
			case "应出勤":
				headersMap[header] = "Should"
			case "实际出勤":
				headersMap[header] = "Actual"
			case "应发工资":
				headersMap[header] = "Standard"
			case "实发工资":
				headersMap[header] = "NetPay"
			case "加班工资":
				headersMap[header] = "OvertimePay"
			case "特殊费用":
				headersMap[header] = "SpecialPay"
			case "扣款":
				headersMap[header] = "Deduction"
			case "合计":
				headersMap[header] = "Account"
			case "备注":
				headersMap[header] = "BackUp"
			default:
				fmt.Printf("UNKNOWN HEADER named %s \n", header)
			}
		}

		mConf.HeadersMap = headersMap
	}

	if len(mConf.OverviewHeader) != 0 {
		overviewHeaderMap := make(map[string]string, len(mConf.OverviewHeader))
		for _, header := range mConf.OverviewHeader {
			//序号由行号决定
			switch header {
			case "区域":
				overviewHeaderMap[header] = "Area"
			case "发放人数":
				overviewHeaderMap[header] = "NumOfStaff"
			case "总计费用":
				overviewHeaderMap[header] = "AccountTotal"
			case "备注":
				overviewHeaderMap[header] = "BackUp"
			default:
				//ignore
			}
		}
		mConf.OverviewHeaderMap = overviewHeaderMap
	}

	fmt.Printf("config %+v \n", mConf)
}
