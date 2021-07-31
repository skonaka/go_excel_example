package main

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func main() {
	const newfile = "Book1.xlsx"

	// Create a new Excel File
	f := excelize.NewFile()
	if err := f.SaveAs(newfile); err != nil {
		fmt.Println(err)
	}

	// Writing numbers
	f, err := excelize.OpenFile(newfile)
	if err != nil {
		fmt.Println(err)
		return
	}

	f.SetCellValue("Sheet1", "A1", "Name")
	f.SetCellValue("Sheet1", "B1", "Number")
	f.SetCellValue("Sheet1", "A2", "aaa")
	f.SetCellValue("Sheet1", "A3", "bbb")
	f.SetCellValue("Sheet1", "A4", "ccc")
	f.SetCellValue("Sheet1", "B2", 100)
	f.SetCellValue("Sheet1", "B3", 200)
	f.SetCellValue("Sheet1", "B4", 300)

	if err := f.SaveAs(newfile); err != nil {
		fmt.Println(err)
	}

	// Read numbers
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	for i, row := range rows {
		for j, colCell := range row {
			if i == 0 || j == 0 {
				continue
			}
			f, _ := strconv.ParseFloat(colCell, 64)
			fmt.Println(colCell, f*1.08)
		}
	}
}
