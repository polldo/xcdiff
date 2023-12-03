package main

import (
	"flag"
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
)

func main() {
	f1 := flag.String("f1", "", "First input file.")
	f2 := flag.String("f2", "", "Second input file.")
	out := flag.String("o", "o.xlsx", "Output file path.")
	flag.Parse()

	if err := run(*f1, *f2, *out); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}

func run(f1name string, f2name string, out string) error {
	f1, err := excelize.OpenFile(f1name)
	if err != nil {
		return err
	}

	f2, err := excelize.OpenFile(f2name)
	if err != nil {
		return err
	}

	sheet := f1.GetSheetName(0)

	rows1, err := f1.GetRows(sheet)
	if err != nil {
		return err
	}

	rows2, err := f2.GetRows(sheet)
	if err != nil {
		return err
	}

	fOut := excelize.NewFile()
	fOut.NewSheet(sheet)

	maxRows := len(rows1)
	if len(rows2) > maxRows {
		maxRows = len(rows2)
	}

	for i := 0; i < maxRows; i++ {
		maxCols := 0
		var row1, row2 []string
		if i < len(rows1) {
			row1 = rows1[i]
			if len(row1) > maxCols {
				maxCols = len(row1)
			}
		}
		if i < len(rows2) {
			row2 = rows2[i]
			if len(row2) > maxCols {
				maxCols = len(row2)
			}
		}

		for j := 0; j < maxCols; j++ {
			var val1, val2 string
			if j < len(row1) {
				val1 = row1[j]
			}
			if j < len(row2) {
				val2 = row2[j]
			}

			var cellValue string
			if val1 == val2 {
				cellValue = val1
			} else {
				cellValue = val1 + " => " + val2
			}

			cellName, _ := excelize.CoordinatesToCellName(j+1, i+1)
			fOut.SetCellValue(sheet, cellName, cellValue)
		}
	}

	if err := fOut.SaveAs(out); err != nil {
		return err
	}
	return nil
}
