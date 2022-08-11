package main

import (
	"bufio"
	"flag"
	"fmt"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

//./excel -type xml -convert from -fromname ./Book12.xlsx -toname anroid_string.xml
//./excel -type xml -convert to -fromname ./com.lenovo.lsf.strings.xml -toname Book12.xlsx
func main() {
	outtype := flag.String("type", "xml", "output type,xml default,xml,ios")
	convert := flag.String("convert", "from", "convert file,from(resource xlsx) or to(resource file)")
	fromname := flag.String("fromname", "", "from file name")
	toname := flag.String("toname", "", "to file name")
	sheetname := flag.String(("sheetname"), "Sheet1", "sheet name")

	flag.Parse()

	if len(*fromname) == 0 || len(*toname) == 0 {
		return
	}

	if *convert == "to" {
		convertToXlsx(outtype, *fromname, *toname, *sheetname)
	} else {
		convertFromXlsx(outtype, *fromname, *toname, *sheetname)
	}
}

func convertFromXlsx(outtype *string, fromname string, toname string, sheetname string) {
	f, err := excelize.OpenFile(fromname)
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	outfile, err := os.Create(toname)

	if err != nil {
		fmt.Println(err)
		return
	}

	defer outfile.Close()

	// Get all the rows in the Sheet1.
	rows, err := f.GetRows(sheetname)
	if err != nil {
		fmt.Println(err)
		return
	}

	for _, row := range rows {
		if len(row) < 2 {
			continue
		}
		if *outtype == "xml" {
			line := "    <string name=\"" + strings.TrimSpace(row[0]) + "\">" + strings.TrimSpace(row[1]) + "</string>\n"
			outfile.Write([]byte(line))
		} else if *outtype == "ios" {
			line := "\"" + strings.TrimSpace(row[0]) + "\"=\"" + strings.TrimSpace(row[1]) + "\";\n"
			outfile.Write([]byte(line))
		}

	}
}

func convertToXlsx(outtype *string, fromname string, toname string, sheetname string) {
	f := excelize.NewFile()
	// Create a new sheet.
	index := f.NewSheet(sheetname)
	file, err := os.Open(fromname)
	if err != nil {
		return
	}

	defer file.Close()

	scanner := bufio.NewScanner(file)
	i := 0
	if *outtype == "xml" {
		for scanner.Scan() {
			line := scanner.Text()
			line = strings.TrimSpace(line)
			if strings.HasPrefix(line, "<!--") {
				continue
			}

			line = strings.Replace(line, "<string name=\"", "", -1)
			line = strings.Replace(line, "</string>", "", -1)
			datas := strings.Split(line, "\">")
			if len(datas) != 2 {
				continue
			}
			// Set value of a cell.
			i++
			f.SetCellValue(sheetname, "A"+fmt.Sprintf("%d", i), datas[0])
			f.SetCellValue(sheetname, "B"+fmt.Sprintf("%d", i), datas[1])
		}

	} else if *outtype == "ios" {
		for scanner.Scan() {
			line := scanner.Text()
			line = strings.TrimSpace(line)
			if strings.HasPrefix(line, "//") {
				continue
			}

			datas := strings.Split(line, "\"=\"")
			if len(datas) != 2 {
				continue
			}
			datas[0] = strings.Replace(datas[0], "\"", "", -1)
			datas[1] = strings.Replace(datas[1], "\"", "", -1)
			datas[1] = strings.Replace(datas[1], ";", "", -1)
			// Set value of a cell.
			i++
			f.SetCellValue(sheetname, "A"+fmt.Sprintf("%d", i), datas[0])
			f.SetCellValue(sheetname, "B"+fmt.Sprintf("%d", i), datas[1])
		}
	}

	// Set active sheet of the workbook.
	f.SetActiveSheet(index)
	// Save spreadsheet by the given path.
	if err := f.SaveAs(toname); err != nil {
		fmt.Println(err)
	}
}
