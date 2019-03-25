package main

import (
	"fmt"
	"log"
	"strconv"

	"github.com/araddon/dateparse"
	docopt "github.com/docopt/docopt-go"
	xlsx "github.com/tealeg/xlsx"
	funk "github.com/thoas/go-funk"
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

	haveSheets := funk.Map(inFile.Sheets, func(sheet *xlsx.Sheet) string {
		return sheet.Name
	}).([]string)

	log.Print("Infile sheets: ", funk.Map(haveSheets, func(name string) string {
		return fmt.Sprintf("'%s'", name)
	}).([]string))

	wantSheets := arguments["<sheet>"].([]string)
	for _, wantSheet := range wantSheets {
		ok := false
		for _, infileSheet := range haveSheets {
			if wantSheet == infileSheet {
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

	for _, sheetName := range wantSheets {
		log.Print("Processing sheet: ", sheetName)
		sheet := inFile.Sheet[sheetName]
		matchedRows := 0
		for i, inRow := range sheet.Rows {
			if len(inRow.Cells) == 0 {
				log.Print("Empty row: #", i)
				continue
			}

			inCell := inRow.Cells[0]

			if inCell.Type() == xlsx.CellTypeNumeric {
				v, _ := inCell.Float()
				t := xlsx.TimeFromExcelTime(v, false)

				if t.Year() == reqDate.Year() && t.Month() == reqDate.Month() && t.Day() == reqDate.Day() {
					matchedRows++
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
		log.Print("Matched rows: ", matchedRows)
	}

	err = outFile.Save(arguments["<outfile>"].(string))
	if err != nil {
		fmt.Printf(err.Error())
	}
}
