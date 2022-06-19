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

// SetExcelFile opens up an excel file by its path.
func SetExcelFile(file_path string) *excelize.File {
	// Open Data
	data, err := excelize.OpenFile(file_path)
	if err != nil {
		fmt.Println(err)
		return nil
	}
	defer func() {
		// Close the spreadsheet.
		if err := data.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	return data
}

func main() {
	// Open Data
	var data *excelize.File

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

		if command[0] == "SetExcelFile" {
			data = SetExcelFile(command[1])
		} else if command[0] == "SetActiveSheet" {
			// Set active sheet
			sheet = command[1]
			SetActiveSheet(data, sheet)
		} else if command[0] == "OrderCol" {
			OrderCols(data, sheet, command)
		} else if command[0] == "InsertNewCol" {
			InsertNewCol(data, sheet, command[1])
		} else if command[0] == "UpdateColFormula" {
			UpdateColFormula(data, sheet, command)
		} else if command[0] == "SaveFileAs" {
			SaveFileAs(data, command[1])
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
func ReplaceFormulaRows(formula string, row_number string) string {
	return strings.Replace(formula, "[RowNum]", row_number, -1)
}

// UpdateRowFormula iterates through a column and sets all values as
// the given formula.
func UpdateRowFormula(data *excelize.File, sheet string, command []string) {
	// Get rows and counts
	cols, _ := data.GetCols(sheet)
	num_cols := len(cols)

	// Generate formula
	formula := strings.Join(command[2:], " ")

	for i := 0; i < num_cols; i++ {
		// Add formula to each cell, replacing formula row placeholder if necessary
		cell := strconv.Itoa(i + 1)
		data.SetCellFormula(sheet, command[1]+cell, "="+ReplaceFormulaCols(formula, cell))
	}

	fmt.Printf("Row %s updated with formula %s...\n", command[1], formula)
}

// ReplaceFormulaCols inserts col values into a formula string with placeholders.
func ReplaceFormulaCols(formula string, column_number string) string {
	return strings.Replace(formula, "[ColNum]", column_number, -1)
}

// SaveFileAs saves the data according to the file name.
func SaveFileAs(data *excelize.File, file_name string) {
	// Update cell formulas.
	data.UpdateLinkedValue()

	if err := data.SaveAs(file_name); err != nil {
		fmt.Println(err)
	}
}

// SetSingleCellValue sets a single cell value to the given value or formula.
func SetSingleCellValue(data *excelize.File, sheet string, cell string,
	value string, is_formula bool) {
	if is_formula {
		data.SetCellFormula(sheet, cell, "="+value)
	} else {
		data.SetCellValue(sheet, cell, value)
	}
}
