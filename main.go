package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"path"
	"strings"

	"github.com/xuri/excelize/v2"
	"golang.org/x/exp/slices"
)

type csvInfos struct {
	path      string
	csvName   string
	sheetName string
}

// create map to change keys
var (
	mapKeys = map[string]string{
		"created_at": "date",
		"id":         "reference",
	}
)

// inventory_mouvements -> InventoryMouvement
func toPascaleCase(name string) string {
	if name == "" {
		return ""
	}
	var result string = ""
	if strings.Contains(name, "_") {
		for _, substring := range strings.Split(name, "_") {
			result = result + toPascaleCase(substring)
		}
	} else {
		firstChart := name[0:1]
		result = strings.ToUpper(firstChart) + name[1:]
	}
	return result
}

func main() {
	// get all args
	commandLineArgs := os.Args[1:]
	// check if path exists
	if len(commandLineArgs) == 0 {
		fmt.Println("Please provide a path")
		return
	}
	// data folder path
	dataFolderPath := commandLineArgs[0]
	// read files
	files, err := os.ReadDir(dataFolderPath)
	if err != nil {
		fmt.Println("error reading the folder", err)
		return
	}
	// declar
	var csvPathsAndNames []csvInfos
	// iterate over files
	for _, file := range files {
		if !file.IsDir() {
			if strings.Contains(file.Name(), ".csv") {
				fileNameWoExt := strings.Split(file.Name(), ".")[0]
				// collect needed data
				csvPathsAndNames = append(csvPathsAndNames, csvInfos{
					sheetName: toPascaleCase(fileNameWoExt),
					csvName:   fileNameWoExt,
					path:      path.Join(dataFolderPath, file.Name()),
				})
			}
		}
	}
	// create excel file
	excelFile := excelize.NewFile(excelize.Options{})
	//
	for _, tableInfos := range csvPathsAndNames {
		printTableInExcel(excelFile, tableInfos)
	}
	// create excel
	excelFile.SaveAs("test.xlsx")
}

func printTableInExcel(excelFile *excelize.File, infos csvInfos) {
	// re arrange headers
	var rearrangeInfos struct {
		image struct {
			exists bool
			cell   string
		}
	}
	//
	err := excelFile.DeleteSheet("Sheet1")
	if err != nil {
		fmt.Println("error deleting sheet", err)
	}
	// get file data
	csvFile, err := os.Open(infos.path)
	if err != nil {
		fmt.Println("error opening file", err)
		return
	}
	defer csvFile.Close()
	// get reader
	reader := csv.NewReader(csvFile)
	// get sheet
	_, err = excelFile.NewSheet(infos.sheetName)
	if err != nil {
		fmt.Println("error creating sheet", err)
		return
	}
	// iterate over rows
	row := 1
	for {
		// get data
		record, err := reader.Read()
		if err == io.EOF {
			fmt.Println("no more input to read", err)
			break
		}
		if err != nil {
			fmt.Println("error reading record 2:", err)
			break
		}
		// check for images
		imageExists := slices.Contains(record, "image")
		rearrangeInfos.image.exists = true
		if imageExists {
			// get image cell
			imageIndex := slices.Index[[]string, string](record, "image")
			imageCell, err := excelize.CoordinatesToCellName(imageIndex+1, 1)
			if err != nil {
				fmt.Println("error getting cell name :", err)
				break
			}
			rearrangeInfos.image.cell = imageCell
		}
		// get cell name from cords
		cell, err := excelize.CoordinatesToCellName(1, row)
		if err != nil {
			fmt.Println("error getting cell name :", err)
			break
		}
		// write row
		if row == 1 {
			// update headers
			for i, header := range record {
				// get if it exists
				updatedKey, ok := mapKeys[header]
				if ok {
					record[i] = toPascaleCase(updatedKey)
					continue
				}
				record[i] = toPascaleCase(header)
			}
			// write headers
			if err := excelFile.SetSheetRow(infos.sheetName, cell, &record); err != nil {
				fmt.Println("coudnt write first row : ", err)
				break
			}
			row++
			continue
		}
		// write headers
		if err := excelFile.SetSheetRow(infos.sheetName, cell, &record); err != nil {
			fmt.Println("coudnt write row : ", err)
			break
		}
		row++
		// write other records
	}
	// re arrange cells
	if rearrangeInfos.image.exists {
		fmt.Println("image exists", rearrangeInfos.image.cell)
	}
}
