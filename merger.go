package main

import (
	"fmt"
	"log"
	"strconv"

	"github.com/araddon/dateparse"
	docopt "github.com/docopt/docopt-go"
	xlsx "github.com/tealeg/xlsx"
)

func main() {
	usage := `Merger.

    Usage:
      merger <infile> <outfile> <date> <column_from> <column_to> <sheet>...
      merger -h | --help
      merger --version

    Options:
      -h --help     Show this screen.
      --version     Show version.
      --drifting    Drifting mine.`

	arguments, _ := docopt.ParseDoc(usage)
	fmt.Printf("%s\n", arguments)

	reqDate, err := dateparse.ParseLocal(arguments["<date>"].(string))
	if err != nil {
		log.Fatal(err)
	}

	columnFrom, err := strconv.Atoi(arguments["<column_from>"].(string))
	if err != nil {
		log.Fatal(err)
	}

	columnTo, err := strconv.Atoi(arguments["<column_to>"].(string))
	if err != nil {
		log.Fatal(err)
	}

	inFile, err := xlsx.OpenFile(arguments["<infile>"].(string))
	if err != nil {
		log.Fatal(err)
	}

	sheets := arguments["<sheet>"].([]string)
	for _, wantSheet := range sheets {
		ok := false
		for _, infileSheet := range inFile.Sheets {
			if wantSheet == infileSheet.Name {
				ok = true
			}
		}

		if !ok {
			log.Fatal(fmt.Errorf("No sheet %s", wantSheet))
		}
	}

	outFile := xlsx.NewFile()
	outSheet, err := outFile.AddSheet("Sheet1")

	if err != nil {
		log.Fatal(err)
	}

	for _, sheetName := range sheets {
		sheet := inFile.Sheet[sheetName]
		for _, inRow := range sheet.Rows {
			inCell := inRow.Cells[0]

			if inCell.Type() == xlsx.CellTypeNumeric {
				v, _ := inCell.Float()
				t := xlsx.TimeFromExcelTime(v, false)
				// fmt.Printf("%s\n", t)

				if t.Year() == reqDate.Year() && t.Month() == reqDate.Month() && t.Day() == reqDate.Day() {
					var sheetCell *xlsx.Cell
					outRow := outSheet.AddRow()

					sheetCell = outRow.AddCell()
					sheetCell.SetValue(sheetName)

					for i := columnFrom; i <= columnTo; i++ {
						sheetCell = outRow.AddCell()
						sheetCell.SetValue(inRow.Cells[i].Value)
					}
				}
			}
		}
	}

	err = outFile.Save(arguments["<outfile>"].(string))
	if err != nil {
		fmt.Printf(err.Error())
	}
}
