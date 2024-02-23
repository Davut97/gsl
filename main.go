package main

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/jung-kurt/gofpdf"
	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("data.xlsx")
	if err != nil {
		fmt.Println("Error opening Excel file:", err)
		return
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println("Error getting rows from worksheet:", err)
		return
	}

	for _, row := range rows {
		name := row[0]
		newSalary := row[1]

		err := generatePDF(name, newSalary)
		if err != nil {
			fmt.Println("Error generating PDF letter:", err)
			return
		}
	}

	fmt.Println("PDF letters generated successfully!")
}

func generatePDF(name, newSalary string) error {
	pdf := gofpdf.New("P", "mm", "A4", "")
	pdf.AddPage()

	pdf.SetFont("Arial", "", 14)

	pdf.CellFormat(0, 10, fmt.Sprintf("Dear %s,", name), "", 0, "L", false, 0, "")
	pdf.Ln(10)
	pdf.CellFormat(0, 10, fmt.Sprintf("We are pleased to inform you that your new salary is $%s.", newSalary), "", 0, "L", false, 0, "")
	pdf.Ln(10)
	pdf.CellFormat(0, 10, "Sincerely,", "", 0, "L", false, 0, "")
	pdf.Ln(10)
	pdf.CellFormat(0, 10, "Your Company Name", "", 0, "L", false, 0, "")

	err := os.MkdirAll("newSalaries", os.ModePerm)
	if err != nil {
		return err
	}

	outputFile := filepath.Join("newSalaries", fmt.Sprintf("%s.pdf", name))
	err = pdf.OutputFileAndClose(outputFile)
	if err != nil {
		return err
	}

	return nil
}
