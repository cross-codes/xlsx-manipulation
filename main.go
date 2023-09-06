package main

import "xlsx-manipulation/helpers"

func main() {
	helpers.CreateNewXLSX("createdsheet.xlsx", "Sheet1")
	helpers.FillXLSXValues("./spreadsheet.xlsx", "./createdsheet.xlsx", "Sheet1", "Sheet1")
	helpers.PrintDetails("./createdsheet.xlsx", "Sheet1")
}
