package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
)

func main() {
	excelFileName := "attendanceRegister.xlsx"
	file := xlsx.NewFile()
	sheet, err := file.AddSheet("Sheet1")


	if err != nil {
		fmt.Printf(err.Error())
	}

	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "I am a cell"

	err = file.Save(excelFileName)
	if err != nil {
		fmt.Println(err.Error())
	}
}
