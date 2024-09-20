package main

import (
	"os"
	"path/filepath"
	"strings"
)

func listXlsxFile(files *[]string) (err error) {

	err = filepath.Walk(mConf.AttendanceFolder, 
		func(path string, info os.FileInfo, err error) error {

			if !info.IsDir() && strings.HasSuffix(path, ".xlsx") {
				*files = append(*files, path)
			}

			return err
	})

	return 
}