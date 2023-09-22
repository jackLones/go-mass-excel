package main

import (
	"database/sql"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"time"

	_ "github.com/go-sql-driver/mysql"
)

func main() {
	db, err := sql.Open("mysql", "root:lyx@20161111@tcp(10.23.141.98:3306)/lyx_data_center")
	if err != nil {
		log.Printf("连接到数据库时发生错误: %v\n", err)
		return
	}
	defer db.Close()

	if err := db.Ping(); err != nil {
		log.Printf("数据库连接失败: %v\n", err)
		return
	}

	db.SetMaxIdleConns(10)
	db.SetMaxOpenConns(100)

	fileName, file := createExcelFile()
	if file == nil {
		return
	}

	sheet := "Sheet1"
	currentRow := 1

	outputExcelFile(db, file, sheet, &currentRow)

	saveExcelFile(fileName, file)
}

func createExcelFile() (string, *excelize.File) {
	now := time.Now().Format("2006-01-02_15-04-05")
	outputPath, _ := os.Executable()
	outputPath = filepath.Dir(outputPath)
	outputPath = filepath.Join(outputPath, "output")
	if err := os.MkdirAll(outputPath, os.ModePerm); err != nil {
		log.Printf("创建目录失败: %v\n", err)
		return "", nil
	}

	filename := filepath.Join(outputPath, fmt.Sprintf("lyx_data_%s.xlsx", now))

	file := excelize.NewFile()

	if err := file.SaveAs(filename); err != nil {
		log.Printf("保存 Excel 文件 '%s' 时发生错误: %v\n", filename, err)
		return "", nil
	}
	return filename, file
}

func outputExcelFile(db *sql.DB, file *excelize.File, sheet string, currentRow *int) {
	pageSize := 1000
	offset := 0
	for {
		rows, err := db.Query("select a.id,a.name,a.province_name,a.city_name,a.area_name,a.address,a.industry_fir,a.reg_cap,b.mobile, b.email from lyx_company as a LEFT JOIN lyx_company_contact as b on a.id = b.company_id where a.province_code = ? and a.industry_fir_code = ? and a.id < ? LIMIT ? OFFSET ?", 33, "C", 2000, pageSize, offset)
		if err != nil {
			log.Printf("执行查询时发生错误: %v\n", err)
			continue
		}

		hasData := batchRows(rows, file, sheet, currentRow)

		rows.Close()

		if !hasData {
			break
		}

		offset += pageSize
	}
}

func batchRows(rows *sql.Rows, file *excelize.File, sheet string, currentRow *int) bool {
	batchSize := 1000
	var rowValues [][]interface{}
	hasData := false
	for rows.Next() {
		hasData = true
		var id, name, provinceName, cityName, areaName, address, industryFir, regCap string
		var mobile, email sql.NullString
		err := rows.Scan(&id, &name, &provinceName, &cityName, &areaName, &address, &industryFir, &regCap, &mobile, &email)
		if err != nil {
			log.Printf("读取行数据时发生错误: %v\n", err)
		}

		rowValues = append(rowValues, []interface{}{id, name, provinceName, cityName, areaName, address, industryFir, regCap, mobile.String, email.String})

		if len(rowValues) >= batchSize {
			writeDataToExcel(file, sheet, *currentRow, rowValues)
			*currentRow += len(rowValues)
			rowValues = [][]interface{}{}
		}
	}

	if len(rowValues) > 0 {
		writeDataToExcel(file, sheet, *currentRow, rowValues)
		*currentRow += len(rowValues)
	}
	return hasData
}

func writeDataToExcel(file *excelize.File, sheet string, row int, data [][]interface{}) {
	startCol := 'A'
	for rowIndex, rowData := range data {
		for colIndex, value := range rowData {
			cellName := fmt.Sprintf("%c%d", startCol+rune(colIndex), row+rowIndex)
			cellValue := ""
			switch v := value.(type) {
			case string:
				cellValue = v
			case int:
				cellValue = strconv.Itoa(v)
			case float64:
				cellValue = strconv.FormatFloat(v, 'f', -1, 64)
			}
			file.SetCellValue(sheet, cellName, cellValue)
		}
	}
}

func saveExcelFile(filename string, file *excelize.File) {
	if err := file.SaveAs(filename); err != nil {
		log.Printf("保存 Excel 文件时发生错误: %v\n", err)
	}
}
