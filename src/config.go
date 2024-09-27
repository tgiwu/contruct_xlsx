package main

import (
	"fmt"
	"runtime"

	"github.com/spf13/viper"
)

const CONFIG_PATH = "config/"

type config struct {
	AttendanceFolder string   `mapstructure:"attendance_folder"`
	StaffFilePath    string   `mapstructure:"staff_file_path"`
	Ignore           []string `mapstructure:"ignore"`
	OutputPath       string   `mapstructure:"output_path"`
	FileName string `mapstructure:"file_name"`
}

type envWin struct {
	Conf config `mapstructure:"config"`
}

type envLin struct {
	Conf config `mapstructure:"config"`
}

type allConf struct {
	EnvWin envWin `mapstructure:"env_win"`
	EnvLin envLin `mapstructure:"env_lin"`
}

var mConf config
var mAllConf allConf

func readConfig() {

	vip := viper.New()
	vip.AddConfigPath(CONFIG_PATH)
	vip.SetConfigName("config.yaml")
	vip.SetConfigType("yaml")

	if err := vip.ReadInConfig(); err != nil {
		if _, ok := err.(viper.ConfigFileNotFoundError); ok {
			fmt.Println("not find " + err.Error())
		} else {
			fmt.Printf("read config file err, %v \n", err)
		}
	}

	err := vip.Unmarshal(&mAllConf)

	switch runtime.GOOS {
	case "windows":
		mConf = mAllConf.EnvWin.Conf
	case "linux":
		mConf = mAllConf.EnvLin.Conf
	case "darwin":
		mConf = mAllConf.EnvLin.Conf
	default:
		fmt.Println("unsupport os ", runtime.GOOS)
	}

	if err != nil {
		panic(err)
	}

	fmt.Printf("config %+v \n", mAllConf)

	fmt.Printf("current config %+v \n", mConf)
}
