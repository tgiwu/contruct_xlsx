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
	FileName         string   `mapstructure:"file_name"`
	Month            int      `mapstructure:"month"`
	Year             int      `mapstructure:"year"`
	Headers          []string `mapstructure:"headers"`
}

var mConf config

func readConfig() {

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

	vip := viper.New()
	vip.AddConfigPath(CONFIG_PATH)
	vip.SetConfigName(configName + ".yaml")
	vip.SetConfigType("yaml")

	if err := vip.ReadInConfig(); err != nil {
		if _, ok := err.(viper.ConfigFileNotFoundError); ok {
			fmt.Println("not find " + err.Error())
		} else {
			fmt.Printf("read config file err, %v \n", err)
		}
	}

	err := vip.Unmarshal(&mConf)

	if err != nil {
		panic(err)
	}

	fmt.Printf("config %+v \n", mConf)

	fmt.Printf("current config %+v \n", mConf)
}
