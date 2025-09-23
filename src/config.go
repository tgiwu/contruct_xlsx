package main

import (
	"bytes"
	"fmt"
	log "github.com/sirupsen/logrus"
	"os"
	"path/filepath"
	"runtime"

	"github.com/spf13/viper"
)

const CONFIG_PATH = "config/"
const CONFIG_COMMON_PATH = "C:/Users/Lenovo/"

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

	initLog()

	vip := viper.New()
	vip.AddConfigPath(CONFIG_COMMON_PATH)
	vip.SetConfigName("config_common_salary.yaml")

	vip.SetConfigType("yaml")

	if err := vip.ReadInConfig(); err != nil {
		if _, ok := err.(viper.ConfigFileNotFoundError); ok {
			log.Infoln("not find " + err.Error())
		} else {
			log.Infof("read config file err, %v \n", err)
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
		log.Infoln("unsupport os ", runtime.GOOS)
	}

	bs, err := os.ReadFile(filepath.Join(CONFIG_PATH, configName+".yaml"))

	vip.MergeConfig(bytes.NewReader(bs))

	if err != nil {
		panic(err)
	}

	err = vip.Unmarshal(&mConf)
	if err != nil {
		log.Panic(err)
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

	log.Infof("config %+v \n", mConf)
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

func initLog() error {
	log.SetOutput(os.Stdout)
	log.SetLevel(log.DebugLevel)
	return  nil
}
