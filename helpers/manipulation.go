// Package helpers stores functions to open and choose content from the spreadsheet and more
package helpers

import (
	"fmt"

	"github.com/tealeg/xlsx/v3"
)

var rbData []string

var students []Student

var departmentMap = map[string]string{
	"A1": "Chemical",
	"A2": "Civil",
	"A3": "EEE",
	"A4": "Mechanical",
	"A5": "Pharmacy",
	"A7": "CSIS",
	"A8": "EnI",
	"AA": "ECE",
	"AB": "Manufacturing",
	"B1": "Biology",
	"B2": "Chemistry",
	"B3": "Economics",
	"B4": "Chemistry",
	"B5": "Physics",
}

// Student describes all atrributes associated with a student
type Student struct {
	Name       string
	ID         string
	Email      string
	Department string
}

func refCellDoer(c *xlsx.Cell) error {
	value, err := c.FormattedValue()
	if err != nil {
		fmt.Println(err.Error())
	} else if value != "" {
		rbData = append(rbData, value)
	}
	return err
}

func refRowDoer(r *xlsx.Row) error {
	return r.ForEachCell(refCellDoer)
}

// CreateNewXLSX will create a new .xlsx file
func CreateNewXLSX(fileName string, sheetName string) {
	wb := xlsx.NewFile()

	_, err := wb.AddSheet(sheetName)
	if err != nil {
		panic(err)
	}

	err = wb.Save(fileName)
	if err != nil {
		panic(err)
	}

	fmt.Println("A new XLSX with name '" + fileName + "' was created succesfully\n")
}

func retParams(fullid string) (string, string, string, string, string) {

	year := fullid[:4]
	stream := fullid[4:6]
	id := fullid[8:12]
	campus := string(fullid[12])
	email := ""
	if campus == "P" {
		email = "f" + year + id + "@pilani.bits-pilani.ac.in"
	} else if campus == "G" {
		email = "f" + year + id + "@goa.bits-pilani.ac.in"
	} else if campus == "H" {
		email = "f" + year + id + "@hyderabad.bits-pilani.ac.in"
	}

	return year, stream, id, campus, email
}

// FillXLSXValues will fill in new values inside the spreasheet
func FillXLSXValues(refFileName string, fileName string, refSheetName string, sheetName string) {
	rwb, err := xlsx.OpenFile(refFileName)
	if err != nil {
		panic(err)
	}
	wb, err := xlsx.OpenFile(fileName)
	if err != nil {
		panic(err)
	}

	rsh, ok := rwb.Sheet[refSheetName]
	if !ok {
		fmt.Println(refSheetName + " does not exist")
		return
	}

	sh, ok := wb.Sheet[sheetName]
	if !ok {
		fmt.Println(sheetName + " does not exist")
		return
	}

	rsh.ForEachRow(refRowDoer)
	arrCount := 0

	for ; arrCount < 2; arrCount++ {
		cell, _ := sh.Cell(0, arrCount)
		cell.Value = rbData[arrCount]
	}

	cell, _ := sh.Cell(0, 2)
	cell.Value = "Email"
	cell, _ = sh.Cell(0, 3)
	cell.Value = "Branch"

	rowIndex := 1
	colIndex := 0

	for ; arrCount < len(rbData); arrCount++ {
		if arrCount%2 != 0 {
			_, stream, _, _, email := retParams(rbData[arrCount])
			for ; colIndex <= 3; colIndex++ {
				cell, _ = sh.Cell(rowIndex, colIndex)
				switch colIndex {
				case 1:
					cell.Value = rbData[arrCount]
				case 2:
					cell.Value = email
				case 3:
					cell.Value = departmentMap[stream]
				}
			}
			rowIndex++
			colIndex = 0
		} else {
			cell, _ = sh.Cell(rowIndex, colIndex)
			cell.Value = rbData[arrCount]
			colIndex++
		}
	}

	err = wb.Save(fileName)
	if err != nil {
		panic(err)
	}
}

func readNewSheet(fileName string, sheetName string) {

	wb, err := xlsx.OpenFile(fileName)
	if err != nil {
		panic(err)
	}

	sh, ok := wb.Sheet[sheetName]
	if !ok {
		fmt.Println(sheetName + " does not exist")
		return
	}
	rbData = nil
	sh.ForEachRow(refRowDoer)
}

func createStudentStructs() {
	for i := 4; i < len(rbData); i += 4 {
		if i%4 == 0 {
			s := Student{
				Name:       rbData[i],
				ID:         rbData[i+1],
				Email:      rbData[i+2],
				Department: rbData[i+3],
			}
			students = append(students, s)
		}
	}
}

// PrintDetails prints all the student details from the new sheet
func PrintDetails(fileName string, sheetName string) {
	readNewSheet(fileName, sheetName)
	createStudentStructs()
	fmt.Println("Reading contents of newly created '" + sheetName + "' in '" + fileName + "'")
	fmt.Println("--------------------------")
	for i := 0; i < len(students); i++ {
		s := students[i]
		fmt.Println("Name:", s.Name)
		fmt.Println("ID: ", s.ID)
		fmt.Println("Email: ", s.Email)
		fmt.Println("Department: ", s.Department)
		fmt.Println("--------------------------")
	}
}
