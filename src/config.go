package main

import (
	"fmt"

	"github.com/spf13/viper"
)

const CONFIG_PATH = "config\\"

type config struct {
	AttendanceFolder string `mapstructure:"attendance_folder"`
}

var mConf config

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

	err := vip.Unmarshal(&mConf)

	if err != nil {
		panic(err)
	}

	// path, err := os.Getwd()

	// if err != nil {
	// 	panic(err)
	// }
	// path = strings.Replace(path, "\\src","", -1)
	// mConf.AttendanceFolder = fmt.Sprint(path, "\\", mConf.AttendanceFolder)

	fmt.Printf("config %+v", mConf)
}
