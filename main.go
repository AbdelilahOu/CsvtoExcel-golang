package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"path"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
	"golang.org/x/exp/slices"
)

type csvInfos struct {
	path      string
	sheetName string
}

// create map to change keys
var (
	mapKeys = map[string]string{
		"created_at": "date",
		"id":         "reference",
	}
)

func main() {
	start := time.Now()
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
	year, month, day := time.Now().Date()
	excelFile.SaveAs(fmt.Sprintf("%d-%d-%d-%d.xlsx", year, month, day, time.Now().UnixMilli()))
	end := time.Now()
	fmt.Println("time elapsed :", end.Sub(start))
}

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

// rearrange an array
func rearrangeArray(array []string, from int, to int) []string {
	if from == to {
		return array
	}
	// get first element
	firstElement := array[from]
	// remove first element
	array = slices.Delete(array, from, from+1)
	// insert first element
	array = slices.Insert(array, to, firstElement)
	return array
}

// csv file to excel sheet
func printTableInExcel(excelFile *excelize.File, infos csvInfos) {
	// re arrange headers
	var rearrangeInfos struct {
		image struct {
			exists     bool
			mouvements struct {
				from int
				to   int
			}
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
			// fmt.Println("no more input to read", err)
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
			// change image to first col
			// check for images
			imageExists := slices.Contains(record, "image")
			rearrangeInfos.image.exists = true
			if imageExists {
				// get image cell
				imageIndex := slices.Index[[]string, string](record, "image")
				rearrangeInfos.image.mouvements.from = imageIndex
				rearrangeInfos.image.mouvements.to = 0
				record = rearrangeArray(record, imageIndex, 0)
			}
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
		if rearrangeInfos.image.exists {
			// check for images
			record = rearrangeArray(record, rearrangeInfos.image.mouvements.from, rearrangeInfos.image.mouvements.to)
		}

		// write headers
		if err := excelFile.SetSheetRow(infos.sheetName, cell, &record); err != nil {
			fmt.Println("coudnt write row : ", err)
			break
		}
		row++
		// write other records
	}
}
