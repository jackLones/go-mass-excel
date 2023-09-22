package main

import (
	"database/sql"
	"fmt"
	_ "github.com/go-sql-driver/mysql"
	"github.com/xuri/excelize/v2"
)

func main() {
	// 连接到 MySQL 数据库
	db, err := sql.Open("mysql", "root:lyx@20161111@tcp(10.23.141.98:3306)/lyx_data_center")
	if err != nil {
		fmt.Println("连接到数据库时发生错误:", err)
		return
	}
	defer db.Close()

	// 执行查询获取数据
	rows, err := db.Query("select a.id,a.name,a.province_name,a.city_name,a.area_name,a.address,a.industry_fir,a.reg_cap,b.mobile, b.email from lyx_company as a LEFT JOIN lyx_company_contact as b on a.id = b.company_id where a.province_code = 33 and a.industry_fir_code = 'C' and a.id < 20000")
	if err != nil {
		fmt.Println("执行查询时发生错误:", err)
		return
	}
	defer rows.Close()

	// 每个 Excel 文件最大行数
	maxRows := 1000000
	// 当前行数
	currentRow := 1
	// 当前 Excel 文件编号
	fileIndex := 1

	// 创建第一个 Excel 文件
	file := createExcelFile(fileIndex, currentRow)
	sheet := "Sheet1"

	for rows.Next() {
		var id, name, provinceName, cityName, areaName, address, industryFir, regCap string
		var mobile, email sql.NullString
		err = rows.Scan(&id, &name, &provinceName, &cityName, &areaName, &address, &industryFir, &regCap, &mobile, &email)
		if err != nil {
			fmt.Println("读取行数据时发生错误:", err)
			return
		}

		// 写入数据到当前 Excel 文件的当前行
		writeDataToExcel(file, sheet, currentRow, id, name, provinceName, cityName, areaName, address, industryFir, regCap, mobile.String, email.String)

		currentRow++

		// 如果当前行达到最大行数，则创建新的 Excel 文件
		if currentRow > maxRows {
			// 保存当前 Excel 文件
			saveExcelFile(file, fileIndex)

			// 创建新的 Excel 文件
			fileIndex++
			file = createExcelFile(fileIndex, 1)
			currentRow = 2
		}
	}

	// 保存最后一个 Excel 文件
	saveExcelFile(file, fileIndex)

	fmt.Println("Excel 文件导出成功。")
}

func createExcelFile(index, startRow int) *excelize.File {
	file := excelize.NewFile()
	sheet := "Sheet1"
	filename := fmt.Sprintf("output_stream_%d.xlsx", index)

	// 写入表头
	file.SetCellValue(sheet, "A1", "ID")
	file.SetCellValue(sheet, "B1", "Name")
	file.SetCellValue(sheet, "C1", "Province Name")
	file.SetCellValue(sheet, "D1", "City Name")
	file.SetCellValue(sheet, "E1", "Area Name")
	file.SetCellValue(sheet, "F1", "Address")
	file.SetCellValue(sheet, "G1", "Industry Fir")
	file.SetCellValue(sheet, "H1", "Reg Cap")
	file.SetCellValue(sheet, "I1", "Mobile")
	file.SetCellValue(sheet, "J1", "Email")

	// 设置起始行号
	file.SetSheetRow(sheet, "A2", &[]interface{}{startRow})

	// 保存文件
	err := file.SaveAs(filename)
	if err != nil {
		fmt.Printf("保存 Excel 文件 '%s' 时发生错误: %v\n", filename, err)
		return nil
	}

	return file
}

func writeDataToExcel(file *excelize.File, sheet string, row int, data ...interface{}) {
	for i, value := range data {
		cellName := fmt.Sprintf("%c%d", 'A'+i, row)
		file.SetCellValue(sheet, cellName, value)
	}
}

func saveExcelFile(file *excelize.File, index int) {
	filename := fmt.Sprintf("output_stream_%d.xlsx", index)
	err := file.SaveAs(filename)
	if err != nil {
		fmt.Printf("保存 Excel 文件 '%s' 时发生错误: %v\n", filename, err)
	}
}
