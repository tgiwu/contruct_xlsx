package main

import (
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"strings"
	log "github.com/sirupsen/logrus"
)

func listXlsxFile(files *[]string) (err error) {

	err = filepath.Walk(mConf.AttendanceFolder,
		func(path string, info os.FileInfo, err error) error {

			if !info.IsDir() && strings.HasSuffix(path, ".xlsx") && !regexp.MustCompile(`^[$#%~]`).MatchString(info.Name()) {
				*files = append(*files, path)
			}

			return err
		})

	return
}

func delFileIfExist(path string, name string) {
	if len(path) == 0 || len(name) == 0 {
		return
	}

	file := filepath.Join(path, name)
	_, err := os.Stat(file)

	if err == nil {
		err = os.Remove(file)
		if err != nil {
			log.Infof("del %s failed: %s\n", file, err.Error())
		}
	}

	if os.IsNotExist(err) {
		log.Infof("file %s not exist \n", file)
	}

	_, err = os.Stat(path)
	if err != nil {
		if os.IsNotExist(err) {
			log.Infof("create out dir %s \n", path)
			os.MkdirAll(path, os.ModePerm)
		}
	}
}

//  func transferInterfaceToString(structField reflect.StructField, data interface{}) string {

// 	var (
// 		float64Ptr *float64
// 		float32Ptr *float32
// 		intPtr     *int
// 		int8Ptr    *int8
// 		int16Ptr   *int16
// 		int32Ptr   *int32
// 		int64Ptr   *int64
// 		stringPtr  *string
// 	)

// 	switch structField.Type {
// 	case reflect.TypeOf(float32Ptr):
// 		d := data.(*float32)
// 		if d == nil {
// 			return ""
// 		} else {
// 			return strconv.FormatFloat(float64(*d), 'f', 6, 32)
// 		}
// 	case reflect.TypeOf(float64Ptr):
// 		d := data.(*float64)
// 		if d == nil {
// 			return ""
// 		} else {
// 			return strconv.FormatFloat(*d, 'f', 6, 64)
// 		}
// 	case reflect.TypeOf(float32(1)):
// 		return strconv.FormatFloat(float64(data.(float32)), 'f', 6, 32)
// 	case reflect.TypeOf(float64(1)):
// 		return strconv.FormatFloat(data.(float64), 'f', 6, 64)
// 	case reflect.TypeOf(intPtr):
// 		d := data.(*int)
// 		if d == nil {
// 			return ""
// 		} else {
// 			return strconv.Itoa(*d)
// 		}
// 	case reflect.TypeOf(int8Ptr):
// 		d := data.(*int)
// 		if d == nil {
// 			return ""
// 		} else {
// 			return strconv.Itoa(*d)
// 		}
// 	case reflect.TypeOf(int16Ptr):
// 		d := data.(*int)
// 		if d == nil {
// 			return ""
// 		} else {
// 			return strconv.Itoa(*d)
// 		}
// 	case reflect.TypeOf(int32Ptr):
// 		d := data.(*int)
// 		if d == nil {
// 			return ""
// 		} else {
// 			return strconv.Itoa(*d)
// 		}
// 	case reflect.TypeOf(int64Ptr):
// 		d := data.(*int)
// 		if d == nil {
// 			return ""
// 		} else {
// 			return strconv.Itoa(*d)
// 		}
// 	case reflect.TypeOf(1):
// 		return strconv.Itoa(data.(int))
// 	case reflect.TypeOf(int8(1)):
// 		return strconv.Itoa(int(data.(int8)))
// 	case reflect.TypeOf(int16(1)):
// 		return strconv.Itoa(int(data.(int16)))
// 	case reflect.TypeOf(int32(1)):
// 		return strconv.Itoa(int(data.(int32)))
// 	case reflect.TypeOf(int64(1)):
// 		return strconv.Itoa(int(data.(int64)))
// 	case reflect.TypeOf(stringPtr):
// 		d := data.(*string)
// 		if d == nil {
// 			return ""
// 		} else {
// 			return *d
// 		}
// 	case reflect.TypeOf(""):
// 		return data.(string)
// 	default:
// 		return ""
// 	}
// }

func sortSalaryById(salaries map[string]Salary) []Salary {
	if len(salaries) == 0 {
		return make([]Salary, 0)
	}

	salariesList := make([]Salary, len(salaries))

	for _, salary := range salaries {
		if salary.Id > -1 {
			salariesList[salary.Id-1] = salary
		}
	}

	return salariesList
}

//计算表格位置，坐标由0开始
func pos(row int, col int) string {
	first := int('A')

	if row < 0 {
		return string(byte(first + col))
	}

	if col < 0 {
		return fmt.Sprint(row + 1)
	}

	return fmt.Sprintf("%s%d", string(byte(first+col)), row+1)
}

func DeepCopy(src, dst interface{}) {
	srcVal := reflect.ValueOf(src).Elem()
	dstVal := reflect.ValueOf(dst).Elem()

	for i := 0; i < srcVal.NumField(); i++ {
		fieldVal := srcVal.Field(i)
		if fieldVal.Kind() == reflect.Ptr {
			if fieldVal.IsNil() {
				continue
			}

			newPtr := reflect.New(fieldVal.Type().Elem()).Elem()
			newPtr.Set(fieldVal.Elem())
			dstVal.Field(i).Set(newPtr.Addr())
		} else {
			dstVal.Field(i).Set(fieldVal)
		}
	}
}