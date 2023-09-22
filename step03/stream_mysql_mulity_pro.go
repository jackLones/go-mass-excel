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

	// 定义分页大小和页码
	pageSize := 1000
	page := 0

	// 创建一个新的 Excel 文件
	file := excelize.NewFile()
	sheet := "Sheet1"

	// 写入表头
	file.SetCellValue(sheet, "A1", "ID")
	file.SetCellValue(sheet, "B1", "名称")
	file.SetCellValue(sheet, "C1", "省份")
	file.SetCellValue(sheet, "D1", "城市")
	file.SetCellValue(sheet, "E1", "区域")
	file.SetCellValue(sheet, "F1", "地址")
	file.SetCellValue(sheet, "G1", "行业")
	file.SetCellValue(sheet, "H1", "注册资本")
	file.SetCellValue(sheet, "I1", "手机号")
	file.SetCellValue(sheet, "J1", "邮箱")

	// 分页读取并写入数据
	row := 2
	for {
		// 执行查询获取当前页数据
		rows, err := db.Query(fmt.Sprintf("select a.id,a.name,a.province_name,a.city_name,a.area_name,a.address,a.industry_fir,a.reg_cap,b.mobile, b.email from lyx_company as a LEFT JOIN lyx_company_contact as b on a.id = b.company_id where a.province_code = 33 and a.industry_fir_code = 'C' LIMIT %d OFFSET %d", pageSize, page*pageSize))
		if err != nil {
			fmt.Println("执行查询时发生错误:", err)
			return
		}

		// 检查是否有更多数据
		hasData := false
		for rows.Next() {
			hasData = true

			var id, name, provinceName, cityName, areaName, address, industryFir, regCap string
			var mobile, email sql.NullString
			err := rows.Scan(&id, &name, &provinceName, &cityName, &areaName, &address, &industryFir, &regCap, &mobile, &email)
			if err != nil {
				fmt.Println("读取行数据时发生错误:", err)
				return
			}

			// 写入数据到 Excel 文件中
			file.SetCellValue(sheet, fmt.Sprintf("A%d", row), id)
			file.SetCellValue(sheet, fmt.Sprintf("B%d", row), name)
			file.SetCellValue(sheet, fmt.Sprintf("C%d", row), provinceName)
			file.SetCellValue(sheet, fmt.Sprintf("D%d", row), cityName)
			file.SetCellValue(sheet, fmt.Sprintf("E%d", row), areaName)
			file.SetCellValue(sheet, fmt.Sprintf("F%d", row), address)
			file.SetCellValue(sheet, fmt.Sprintf("G%d", row), industryFir)
			file.SetCellValue(sheet, fmt.Sprintf("H%d", row), regCap)
			file.SetCellValue(sheet, fmt.Sprintf("I%d", row), mobile.String)
			file.SetCellValue(sheet, fmt.Sprintf("J%d", row), email.String)

			row++
		}

		// 关闭当前页的结果集
		rows.Close()

		// 如果没有更多数据，则退出循环
		if !hasData {
			break
		}

		// 增加页码
		page++
		fmt.Printf("导入到第%d页:\n", page)
		if page > 100 {
			break
		}
	}

	// 保存 Excel 文件
	err = file.SaveAs("output.xlsx")
	if err != nil {
		fmt.Println("保存 Excel 文件时发生错误:", err)
		return
	}

	fmt.Println("Excel 文件导出成功。")
}
