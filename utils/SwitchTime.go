// Package utils
/**
@author WS
@date 2024年10月19日 22:13:38
@packageName
@className SwitchTime
@version 1.0.0
@describe SwitchTime
**/
package utils

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"math"
	"strconv"
	"strings"
	"time"
)

func TimeToSeconds(timeStr string) (float64, error) {
	// 使用 strings.Split 来拆分时间部分
	parts := strings.Split(timeStr, ":")
	//log.Println(parts)
	if len(parts) != 3 {
		return 0, fmt.Errorf("invalid time format")
	}

	// 解析小时、分钟
	hours, err := strconv.Atoi(parts[0])
	if err != nil {
		return 0, fmt.Errorf("invalid hours: %v", err)
	}

	minutes, err := strconv.Atoi(parts[1])
	if err != nil {
		return 0, fmt.Errorf("invalid minutes: %v", err)
	}

	// 解析秒和小数部分
	seconds, err := strconv.ParseFloat(parts[2], 64)
	if err != nil {
		return 0, fmt.Errorf("invalid seconds: %v", err)
	}

	// 计算总秒数
	totalSeconds := float64(hours*3600+minutes*60) + seconds
	return totalSeconds, nil
}

func SwitchTime(filename, sheetName string) {

	times, err := ReadFirstColumnAsString(filename, sheetName)
	if err != nil {
		log.Fatalln("数据错误,请检查", err)
	}
	zeroTime := 0.0
	var data [][]interface{}

	for _, time := range times {

		seconds, err := TimeToSeconds(time)
		if err != nil {
			log.Fatalln("计算错误", err)
		}
		if zeroTime == 0.0 {
			zeroTime = math.Floor(seconds)
		}
		//log.Println("zeroTime===>", zeroTime)
		data = append(data, []interface{}{
			time, seconds - zeroTime,
		})
	}

	err = WriteTimeToExcel(filename, sheetName, data)
	if err != nil {
		log.Fatalln("写入错误")
	}
	log.Println("处理完成......2秒后自动关闭")
	time.Sleep(time.Second * 2)

}

func WriteTimeToExcel(filename, sheetName string, data [][]interface{}) error {
	var f *excelize.File
	var err error

	// 打开存在的文件，如果文件不存在则创建新的文件
	if f, err = excelize.OpenFile(filename); err != nil {
		f = excelize.NewFile() // 如果文件不存在则创建一个新文件
	}

	// 检查工作表是否存在
	num, err := f.GetSheetIndex(sheetName)
	if err != nil {
		index, _ := f.NewSheet(sheetName)
		f.SetActiveSheet(index)
	}
	if num == -1 {
		index, _ := f.NewSheet(sheetName)
		f.SetActiveSheet(index)

		//// 写入表头
		//headers := []interface{}{"初始电位(V)", "高电位(V)", "低电位(V)", "扫描速度(V/s)", "扫描方向(+/-)", "时间 (s)", "计算电位 (V)"}
		//if err := f.SetSheetRow(sheetName, "A1", &headers); err != nil {
		//	return err
		//}
	}

	//// 找到当前数据的最后一行
	//lastRow, err := f.GetRows(sheetName)
	//if err != nil {
	//	return err
	//}
	//nextRow := len(lastRow) + 1 // 下一行
	nextRow := 1 // 下一行

	// 写入数据
	for i, row := range data {
		cell := fmt.Sprintf("A%d", nextRow+i) // 从最后一行开始写入
		if err := f.SetSheetRow(sheetName, cell, &row); err != nil {
			return err
		}
	}
	// 保存文件
	if err := f.SaveAs(filename); err != nil {
		return err
	}

	return nil
}

// WriteToSecondColumn 将一列数据写入Excel的第二列
func WriteToSecondColumn(filename string, data []interface{}) error {
	var f *excelize.File
	var err error

	// 尝试打开文件，文件不存在则创建新文件
	if f, err = excelize.OpenFile(filename); err != nil {
		f = excelize.NewFile() // 如果文件不存在则创建新文件
	}

	sheetName := "Sheet1"

	// 检查工作表是否存在，不存在则创建
	num, _ := f.GetSheetIndex(sheetName)
	if num == -1 {
		index, _ := f.NewSheet(sheetName)
		f.SetActiveSheet(index)
	}

	// 获取当前工作表的已有行数，确定从哪一行开始写入
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return err
	}
	nextRow := len(rows) + 1 // 从下一行开始写入

	// 写入数据到第二列(B列)
	for i, value := range data {
		cell := fmt.Sprintf("B%d", nextRow+i) // 写入第二列的单元格，如 B1, B2, B3...
		if err := f.SetCellValue(sheetName, cell, value); err != nil {
			return err
		}
	}

	// 保存文件
	if err := f.SaveAs(filename); err != nil {
		return err
	}

	return nil
}
