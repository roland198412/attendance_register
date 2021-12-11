package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"strconv"
	"strings"
	"time"
)

func main() {
	arguments := os.Args[1:]

	usage := fmt.Sprintf("./prog fromMonthInNumber toMonthInNumber year | E.g ./prog 1 12 2013")

	if len(arguments) < 3 {
		fmt.Println(usage)
		os.Exit(-1)
	}

	monthfrom, _ := strconv.Atoi(arguments[0])
	monthto, _ := strconv.Atoi(arguments[1])
	year, _ := strconv.Atoi(arguments[2])

	if monthfrom < 1 || monthfrom > 12 || monthfrom > monthto {
		fmt.Println(usage)
		os.Exit(-1)
	} else if year < 1900 || year > 2050 {
		fmt.Println(usage)
		os.Exit(-1)
	}

	saveExcel(monthfrom, monthto, year)
}

func headingBoldStyle(align, bottomBorderStyle string) *xlsx.Style {
	headingBoldStyle := xlsx.NewStyle()
	headingBoldStyleFont := xlsx.NewFont(12, "Arial")
	headingBoldStyleFont.Bold = true
	headingBoldStyle.Font = *headingBoldStyleFont
	headingBoldStyle.ApplyFont = true
	headingBoldStyle.Alignment.Horizontal = align
	headingBoldStyle.ApplyAlignment = true
	headingBoldStyle.Border = *xlsx.NewBorder("thin", "thin", "thin", bottomBorderStyle)
	headingBoldStyle.ApplyBorder = true
	return headingBoldStyle
}

func generalRowStyle() *xlsx.Style {
	centerAlign := xlsx.NewStyle()
	centerAlign.Alignment.Horizontal = "center"
	centerAlign.Border = *xlsx.NewBorder("thin", "thin", "thin", "thin")
	centerAlign.ApplyBorder = true
	centerAlign.ApplyAlignment = true
	return centerAlign
}

func fillBlank() *xlsx.Style {
	sfill := xlsx.NewStyle()
	sfill.Border = *xlsx.NewBorder("thin", "thin", "thin", "thin")
	sfill.ApplyBorder = true
	sfill.Fill = fill()
	sfill.ApplyFill = true
	return sfill
}

func fill() xlsx.Fill {
	fill := xlsx.NewFill("solid", "00C0C0C0", "FF000000")
	return *fill
}

func saveExcel(from, to, year int) {
	excelFileName := fmt.Sprintf("Attendance Register %d.xlsx", year)
	file := xlsx.NewFile()

	for i := from; i <= to; i++ {
		// Parsing date
		start, err := time.Parse("2006-1-2", fmt.Sprintf("%s-%s-%s", strconv.Itoa(year), strconv.Itoa(i), "01"))
		if err != nil {
			fmt.Println(err)
			os.Exit(-1)
		}

		monthYear := fmt.Sprintf("%s %s", start.Month(), strconv.Itoa(start.Year()))

		sheet, err := file.AddSheet(monthYear)
		if err != nil {
			fmt.Printf(err.Error())
		}

		/**
		Building headings
		*/

		row := sheet.AddRow()

		cell := row.AddCell()
		cell.SetStyle(headingBoldStyle("left", "double"))

		cell.Value = "MONTH: " + strings.ToUpper(monthYear)
		cell.String()
		cell.HMerge = 1
		cell = row.AddCell()

		cell = row.AddCell()
		cell.SetStyle(headingBoldStyle("center", "double"))
		cell.Value = "SIGNATURE"

		cell = row.AddCell()
		cell.SetStyle(headingBoldStyle("center", "double"))
		cell.Value = "SIGN-IN"

		cell = row.AddCell()
		cell.SetStyle(headingBoldStyle("center", "double"))
		cell.Value = "SIGN-OUT"

		/**
		second tier headings
		*/
		row = sheet.AddRow()
		cell = row.AddCell()
		cell.Value = "DATE"
		cell.SetStyle(headingBoldStyle("center", "double"))

		cell = row.AddCell()
		cell.Value = "DAY"
		cell.SetStyle(headingBoldStyle("center", "double"))

		cell = row.AddCell()
		cell.SetStyle(fillBlank())
		cell = row.AddCell()
		cell.SetStyle(fillBlank())
		cell = row.AddCell()
		cell.SetStyle(fillBlank())

		for d := start; d.Month() == start.Month(); d = d.AddDate(0, 0, 1) {
			if d.Weekday().String() == "Monday" && d.Day() != 1 {
				row = sheet.AddRow()
				cell = row.AddCell()
				cell.HMerge = 4
				cell.SetStyle(fillBlank())
			}

			row = sheet.AddRow()
			cell = row.AddCell()
			cell.SetStyle(generalRowStyle())
			cell.Value = strconv.Itoa(d.Day())
			cell = row.AddCell()
			cell.SetStyle(generalRowStyle())
			cell.Value = d.Weekday().String()
			cell = row.AddCell()
			cell.SetStyle(generalRowStyle())
			cell = row.AddCell()
			cell.SetStyle(generalRowStyle())
			cell = row.AddCell()
			cell.SetStyle(generalRowStyle())

			//	fmt.Println(d.Month(), d.Weekday(), d.Day())
		}

		// Setting the column widths
		col1 := sheet.Cols[0]
		col1.Width = 6.45

		col2 := sheet.Cols[1]
		col2.Width = 15.4

		col3 := sheet.Cols[2]
		col3.Width = 13.5

		col4 := sheet.Cols[3]
		col4.Width = 13.5

		col5 := sheet.Cols[4]
		col5.Width = 13.5
	}

	err := file.Save(excelFileName)
	if err != nil {
		fmt.Println(err.Error())
	}
}
