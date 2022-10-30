package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	// 開啟 file
	fileName := "golang-xlsx-example.xlsx" // @@@@@ 這邊需要改動 @@@@@
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// 獲取分頁上所有存儲格
	sheetName := "sheet1" // 標籤分頁名稱 @@@@@ 這邊需要改動 @@@@@
	// sheetName := "sheet2" // 標籤分頁名稱

	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println(err)
		return
	}

	for _, v := range rows[1] { // excel 第二橫列
		// 業務邏輯
		fmt.Println(v)
	}
}
