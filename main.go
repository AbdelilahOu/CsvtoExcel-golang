package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"path"
	"strings"

	"github.com/xuri/excelize/v2"
)

type csvInfos struct {
	path      string
	csvName   string
	sheetName string
}

// inventory_mouvements -> InventoryMouvement
func getSheetName(name string) string {
	var result string = ""
	if strings.Contains(name, "_") {
		for _, substring := range strings.Split(name, "_") {
			result = result + getSheetName(substring)
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
					sheetName: getSheetName(fileNameWoExt),
					csvName:   fileNameWoExt,
					path:      path.Join(dataFolderPath, file.Name()),
				})
			}
		}
	}
	fmt.Println(csvPathsAndNames[1])
	for i,tableInfos := range csvPathsAndNames {
		printTableInExcel(&)
	}
	// create excel
	// file := excelize.NewFile(excelize.Options{})
}

func printTableInExcel(excelFile *excelize.File, infos csvInfos) {
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
			fmt.Println("error reading records", err)
			break
		}
		if err != nil {
			fmt.Println("error reading record 2:", err)
			break
		}
		// get cell name from cords
		cell, err := excelize.CoordinatesToCellName(1, row)
		if err != nil {
			fmt.Println("error getting cell name :", err)
			break
		}
		// write row
		if row == 1 {
			if err := excelFile.SetSheetRow(infos.sheetName, cell, &record); err != nil {
				fmt.Println("coudnt write first row : ", err)
				break
			}
			row++
			continue
		}
	}
}
