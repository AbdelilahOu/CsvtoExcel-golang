package main

import (
	"fmt"
	"os"
	"path"
	"strings"
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
	// create excel
	// file := excelize.NewFile(excelize.Options{})
}
