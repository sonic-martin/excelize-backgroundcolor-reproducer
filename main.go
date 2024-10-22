package main

import (
	"bytes"
	_ "embed"
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	worksheet = "Table1"
	strCell   = "A2"
	dateCell  = "B2"
)

//go:embed template.xlsx
var _excelTemplate []byte

func main() {
	template := bytes.NewReader(_excelTemplate)
	excel, errOpenReader := excelize.OpenReader(template)
	if errOpenReader != nil {
		fmt.Println(errOpenReader)
	}
	defer func() {
		if err := excel.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	strValue := "Fubar"
	if err := excel.SetCellValue(worksheet, strCell, strValue); err != nil {
		fmt.Println(err)
	}

	dateValue := time.Now()
	if err := excel.SetCellValue(worksheet, dateCell, dateValue); err != nil {
		fmt.Println(err)
	}

	if err := excel.SaveAs("doc.xlsx"); err != nil {
		fmt.Println(err)
	}
}
