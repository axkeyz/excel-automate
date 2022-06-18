package main

import (
	"bufio"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func MakeFile() {
	f := excelize.NewFile()
	// Create a new sheet.
	index := f.NewSheet("Sheet2")
	// Set value of a cell.
	f.SetCellValue("Sheet2", "A2", "Hello world.")
	f.SetCellValue("Sheet1", "B2", 100)
	// Set active sheet of the workbook.
	f.SetActiveSheet(index)
	// Save spreadsheet by the given path.
	if err := f.SaveAs("user/results/Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}

func OpenUpload() {
	f, err := excelize.OpenFile("user/uploads/SampleData.xlsx")
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

	for index, name := range f.GetSheetMap() {
		fmt.Println(index, name)
	}
}

func main() {
	// Open Data
	data, err := excelize.OpenFile("user/uploads/SampleData.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := data.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Open Template
	template, err := os.Open("user/templates/test.txt")

	if err != nil {
		fmt.Println(err)
	}

	defer template.Close()

	scanner := bufio.NewScanner(template)
	var sheet string

	for scanner.Scan() {
		command := strings.Split(scanner.Text(), " ")

		if command[0] == "SetActiveSheet" {
			// Set active sheet
			sheet = command[1]
			SetActiveSheet(data, sheet)
		} else if command[0] == "OrderCol" {
			OrderCols(data, sheet, command)
		} else if command[0] == "InsertNewCol" {
			InsertNewCol(data, sheet, command[1])
		} else if command[0] == "UpdateColFormula" {
			UpdateColFormula(data, sheet, command)
		} else {
			fmt.Printf("%s command doesn't exist (yet?)\n", command[0])
		}
	}

	if err := scanner.Err(); err != nil {
		fmt.Println(err)
	}

	rows, _ := data.GetRows(sheet)
	fmt.Println("\n\nSample Output...")
	for i := 0; i < 2; i++ {
		fmt.Println(rows[i])
	}

	if err := data.SaveAs("user/results/Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}

// SetActiveSheet sets a sheet as the active sheet (first sheet to be
// opened) and will be used for the next automation steps until a new
// active sheet is set (if any).
func SetActiveSheet(data *excelize.File, sheet string) {
	data.SetActiveSheet(data.GetSheetIndex(sheet))
	fmt.Printf("Set active sheet as %s...\n", sheet)
}

// OrderCols orders the columns of an excel sheet, given an alphabetical
// ordering slice. The first value of the ordering slice is irrelevant.
func OrderCols(data *excelize.File, sheet string, ordering []string) {
	var col_indexes []int
	var cell string

	// Get rows and counts
	rows, _ := data.GetRows(sheet)
	num_rows := len(rows)
	num_cols := len(rows[0])

	for i := 0; i < num_cols; i++ {
		// Convert the new ordering alphabeticals into numeric form
		col_index, _ := excelize.ColumnNameToNumber(ordering[i+1])
		col_indexes = append(col_indexes, col_index)
	}

	for i := 0; i < num_rows; i++ {
		for j := 0; j < num_cols; j++ {
			// Get alphanumerical name of destination cell
			cell, _ = excelize.CoordinatesToCellName(j+1, i+1)

			// Set new value of destination cell with swapped cell value
			data.SetCellValue(sheet, cell, rows[i][col_indexes[j]-1])
		}
	}

	fmt.Printf("Order columns in order %x...\n", col_indexes)
}

// InsertNewCol inserts a new column before the given column in a sheet.
func InsertNewCol(data *excelize.File, sheet string, next_col string) {
	data.InsertCol(sheet, next_col)
	fmt.Printf("Inserted column before %s...\n", next_col)
}

// UpdateColFormula iterates through a column and sets all values as
// the given formula.
func UpdateColFormula(data *excelize.File, sheet string, command []string) {
	// Get rows and counts
	rows, _ := data.GetRows(sheet)
	num_rows := len(rows)

	// Generate formula
	formula := strings.Join(command[2:], " ")

	for i := 0; i < num_rows; i++ {
		// Add formula to each cell, replacing formula row placeholder if necessary
		cell := strconv.Itoa(i + 1)
		data.SetCellFormula(sheet, command[1]+cell, "="+ReplaceFormulaRows(formula, cell))
	}

	fmt.Printf("Column %s updated with formula %s...\n", command[1], formula)
}

// ReplaceFormulaRows inserts row values into a formula string with placeholders.
func ReplaceFormulaRows(formula string, replacement string) string {
	return strings.Replace(formula, "[RowNum]", replacement, -1)
}
