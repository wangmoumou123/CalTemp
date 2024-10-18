// Package utils
/**
@author WS
@date 2024年10月17日 16:47:22
@packageName
@className Camtt
@version 1.0.0
@describe Camtt
**/
package utils

import (
	"bufio"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"math"
	"os"
	"strconv"
	"strings"
)

// PotentialCalculator 定义电位转换的结构体
type PotentialCalculator struct {
	InitialPotential float64 // 初始电位
	HighPotential    float64 // 高电位
	LowPotential     float64 // 低电位
	ScanDirection    string  // 扫描方向（正向/反向）
	ScanSpeed        float64 // 扫描速度 (V/s)
}

func (pc *PotentialCalculator) calculatePotential(time float64) float64 {
	var val float64
	t := time * pc.ScanSpeed
	w := pc.HighPotential - pc.LowPotential
	//x - math.Floor(x/y)*y
	total := t - math.Floor(t/w)*w
	log.Println(t, "====", w, "====", total)
	if pc.ScanDirection == "+" {
		p := pc.HighPotential - pc.InitialPotential
		if total >= p {
			total = total - math.Floor(total/p)*p
		}
		val = pc.InitialPotential + total
		if val >= pc.HighPotential {
			val = pc.HighPotential
			pc.ScanDirection = "-"
		}
	} else {
		q := pc.InitialPotential - pc.LowPotential
		if total >= q {
			total = q - (total - math.Floor(total/q)*q)
		}
		val = pc.InitialPotential - total
		if val <= pc.LowPotential {
			val = pc.LowPotential
			pc.ScanDirection = "+"
		}

	}
	return val
}

// 用户输入解析函数
func getUserInput(prompt string) float64 {
	var input string
	for {
		fmt.Print(prompt)
		_, err := fmt.Scanln(&input)
		if err != nil {
			fmt.Println("输入有误: ", err)
			continue
		}
		value, err := strconv.ParseFloat(strings.TrimSpace(input), 64)
		if err != nil {
			fmt.Println("输入的值无法转换为数字: ", err)
			continue
		}
		return value
	}

}

// Camm 主函数
func Camm(fileName string) {
	//// 获取用户输入
	//fmt.Println()
	//initialPotential := getUserInput("请输入初始电位 (V): ")
	//fmt.Println()
	//highPotential := getUserInput("请输入高电位 (V): ")
	//fmt.Println()
	//lowPotential := getUserInput("请输入低电位 (V): ")
	//fmt.Println()
	//scanSpeed := getUserInput("请输入扫描速度 (V/s): ")
	//fmt.Println()
	//
	//// 获取并验证扫描方向
	//fmt.Print("请输入扫描方向 (+/-): ")
	//var scanDirection string
	//_, err := fmt.Scanln(&scanDirection)
	//if err != nil {
	//	fmt.Println("输入扫描方向时发生错误: ", err)
	//	return
	//}
	//
	//// 验证扫描方向输入
	//if scanDirection != "+" && scanDirection != "-" {
	//	fmt.Println("无效的扫描方向，必须是 '+' 或 '-'")
	//	return
	//}

	//// 初始化电位转换器
	//calculator := &PotentialCalculator{
	//	InitialPotential: initialPotential,
	//	HighPotential:    highPotential,
	//	LowPotential:     lowPotential,
	//	ScanDirection:    scanDirection,
	//	ScanSpeed:        scanSpeed,
	//}
	initialPotential := 0.0
	// 初始化电位转换器
	calculator := &PotentialCalculator{
		InitialPotential: initialPotential,
		HighPotential:    1.6,
		LowPotential:     -2.8,
		ScanDirection:    "-",
		ScanSpeed:        0.05,
	}

	//calculator := &PotentialCalculator{
	//	InitialPotential: initialPotential,
	//	HighPotential:    5,
	//	LowPotential:     0,
	//	ScanDirection:    "-",
	//	ScanSpeed:        1,
	//}
	for {
		fmt.Println()
		// 获取时间序列数据
		fmt.Print("请输入时间数据（以空格分隔的秒数列表，例如: 0 1 2 3, 按q返回上一级菜单）: ")
		var timeInput string
		reader := bufio.NewReader(os.Stdin)
		timeInput, err := reader.ReadString('\n')
		if err != nil {
			fmt.Println("获取时间输入时发生错误: ", err)
			return
		}
		if strings.TrimSpace(timeInput) == "q" {
			break
		}

		// 解析时间数据
		timeStrings := strings.Fields(timeInput)
		if len(timeStrings) == 0 {
			fmt.Println("未提供有效的时间数据")
			return
		}
		var data [][]interface{}
		fmt.Println("======================================================")
		for _, timeStr := range timeStrings {
			timeValue, err := strconv.ParseFloat(timeStr, 64)
			if err != nil {
				fmt.Printf("时间数据格式错误: '%s' 无法转换为数字: %v", timeStr, err)
				return
			}
			// 计算电位
			potential := calculator.calculatePotential(timeValue)
			data = append(data, []interface{}{
				initialPotential,
				calculator.HighPotential,
				calculator.LowPotential,
				calculator.ScanSpeed,
				calculator.ScanDirection,
				timeValue,
				potential,
			})
			fmt.Printf("时间: %.4f 秒 ===> 初始电位: %.4fV === 计算电位: %.4f V\n", timeValue, initialPotential, potential)
		}
		err = writeToExcel(fileName, data)
		if err != nil {
			fmt.Println(err)
			fmt.Println("======error_data====")
			fmt.Println(data)
		}
		fmt.Println("======================================================")

	}
}

// 将数据写入 Excel
func writeToExcel(filename string, data [][]interface{}) error {
	var f *excelize.File
	var err error

	// 打开存在的文件，如果文件不存在则创建新的文件
	if f, err = excelize.OpenFile(filename); err != nil {
		f = excelize.NewFile() // 如果文件不存在则创建一个新文件
	}

	sheetName := "Results"

	// 检查工作表是否存在
	num, err := f.GetSheetIndex(sheetName)
	if err != nil {
		index, _ := f.NewSheet(sheetName)
		f.SetActiveSheet(index)
	}
	if num == -1 {
		index, _ := f.NewSheet(sheetName)
		f.SetActiveSheet(index)

		// 写入表头
		headers := []interface{}{"初始电位(V)", "高电位(V)", "低电位(V)", "扫描速度(V/s)", "扫描方向(+/-)", "时间 (s)", "计算电位 (V)"}
		if err := f.SetSheetRow(sheetName, "A1", &headers); err != nil {
			return err
		}
	}

	// 找到当前数据的最后一行
	lastRow, err := f.GetRows(sheetName)
	if err != nil {
		return err
	}
	nextRow := len(lastRow) + 1 // 下一行

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
