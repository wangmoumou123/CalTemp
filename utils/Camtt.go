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
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"
)

// PotentialCalculator 定义电位转换的结构体
type PotentialCalculator struct {
	InitialPotential float64 // 初始电位
	HighPotential    float64 // 高电位
	LowPotential     float64 // 低电位
	ScanDirection    string  // 扫描方向（正向/反向）
	ScanSpeed        float64 // 扫描速度 (V/s)
}

func helpCp(x, y float64) float64 {
	//fmt.Println(x > y)
	if x > y {
		return 2*y - x
	} else {
		return x
	}
}

func (pc *PotentialCalculator) calculatePotential(time float64) float64 {
	left := pc.InitialPotential - pc.LowPotential
	right := pc.HighPotential - pc.InitialPotential
	t := time * pc.ScanSpeed
	w := pc.HighPotential - pc.LowPotential
	alt := math.Mod(t, 2*w)
	//log.Println(t, "====", w, "====", alt)
	if pc.ScanDirection == "+" {
		if alt < 2*right {
			//right
			return pc.InitialPotential + helpCp(alt, right)
		} else {
			//left
			return pc.InitialPotential - helpCp(alt-2*right, left)
		}

	} else {
		if alt < 2*left {
			//left
			return pc.InitialPotential - helpCp(alt, left)
		} else {
			//right
			return pc.InitialPotential + helpCp(alt-2*left, right)
		}

	}
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
func Camm() {
	// 获取用户输入
	fmt.Println()
	initialPotential := getUserInput("请输入初始电位 (V): ")
	fmt.Println()
	highPotential := getUserInput("请输入高电位 (V): ")
	fmt.Println()
	lowPotential := getUserInput("请输入低电位 (V): ")
	fmt.Println()
	scanSpeed := getUserInput("请输入扫描速度 (V/s): ")
	fmt.Println()

	// 获取并验证扫描方向
	fmt.Print("请输入扫描方向 (+/-): ")
	var scanDirection string
	_, err := fmt.Scanln(&scanDirection)
	if err != nil {
		fmt.Println("输入扫描方向时发生错误: ", err)
		return
	}

	// 验证扫描方向输入
	if scanDirection != "+" && scanDirection != "-" {
		fmt.Println("无效的扫描方向，必须是 '+' 或 '-'")
		return
	}

	// 初始化电位转换器
	calculator := &PotentialCalculator{
		InitialPotential: initialPotential,
		HighPotential:    highPotential,
		LowPotential:     lowPotential,
		ScanDirection:    scanDirection,
		ScanSpeed:        scanSpeed,
	}
	//initialPotential := 0.0
	//// 初始化电位转换器
	//calculator := &PotentialCalculator{
	//	InitialPotential: initialPotential,
	//	HighPotential:    1.6,
	//	LowPotential:     -2.8,
	//	ScanDirection:    "-",
	//	ScanSpeed:        0.05,
	//}

	//calculator := &PotentialCalculator{
	//	InitialPotential: initialPotential,
	//	HighPotential:    5,
	//	LowPotential:     0,
	//	ScanDirection:    "-",
	//	ScanSpeed:        1,
	//}
	for {
		fileName := fmt.Sprintf("ws_%s.xlsx", time.Now().Format("2006_01_02_15_04_05"))

		//var timeStrings []string
		fmt.Println()
		// 获取时间序列数据
		//fmt.Print("请输入时间数据（以空格分隔的秒数列表，例如: 0 1 2 3, 按q返回上一级菜单,输入data从文件中读取）: ")
		fmt.Print("输入原数据文件名字,比如:data.txt或data: ")
		var timeInput string
		reader := bufio.NewReader(os.Stdin)
		timeInput, err := reader.ReadString('\n')
		if err != nil {
			fmt.Println("获取时间输入时发生错误: ", err)
			continue
		}
		timeInput = strings.TrimSpace(timeInput)
		ext := filepath.Ext(timeInput)
		if ext == "" {
			timeInput += ".txt"
			ext = ".txt"
		}
		//log.Println(ext, ext == ".txt")
		if strings.ToLower(ext) != ".txt" {
			fmt.Println("文件类型错误:", timeInput, ",请重新重新输入!")
			continue
		}
		if !fileExists(timeInput) {
			fmt.Println("文件不存在或名字错误", os.IsNotExist(err))
			continue
		}
		timeStrings, err := ParseTimeDifferences(timeInput)
		if err != nil {
			fmt.Println("转换时间失败", err)
			continue
		}

		//if strings.TrimSpace(timeInput) == "q" {
		//	break
		//} else if strings.TrimSpace(timeInput) == "data" {
		//	tString, err := ReadFirstColumnAsString("data.xlsx", "Sheet1")
		//	if err != nil {
		//		fmt.Println("文件数据错误,请检查!", err)
		//		continue
		//	}
		//	timeStrings = tString
		//	fmt.Println(timeStrings)
		//} else {
		//	// 解析时间数据
		//	timeStrings = strings.Fields(timeInput)
		//	if len(timeStrings) == 0 {
		//		fmt.Println("未提供有效的时间数据")
		//		return
		//	}
		//}
		var data [][]interface{}
		fmt.Println("======================================================")
		for _, timeValue := range timeStrings {
			//for _, timeStr := range timeStrings {
			//timeValue, err := strconv.ParseFloat(timeStr, 64)
			//if err != nil {
			//	fmt.Printf("时间数据格式错误: '%s' 无法转换为数字: %v", timeStr, err)
			//	return
			//}
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
			fmt.Printf("时间: %.4f 秒 === 计算电位: %.4f V\n", timeValue, potential)
		}
		err = WriteToExcel(fileName, data)
		if err != nil {
			fmt.Println(err)
			fmt.Println("======error_data====")
			fmt.Println(data)
		}
		fmt.Println("======================================================")

	}
}

// 判断文件是否存在
func fileExists(filename string) bool {
	_, err := os.Stat(filename)
	// 如果文件存在则返回 true，文件不存在或者有错误则返回 false
	return err == nil || os.IsExist(err)
}

// WriteToExcel 将数据写入 Excel
func WriteToExcel(filename string, data [][]interface{}) error {
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

func ReadFirstColumnAsString(filePath string, sheetName string) ([]string, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	defer func() {
		if err := f.Close(); err != nil {
			log.Println("Error closing the file:", err)
		}
	}()

	var data []string

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}

	// 遍历第一列（索引为0的列）
	for _, row := range rows {
		if len(row) > 0 {
			data = append(data, row[0])
		}
	}

	return data, nil
}

// ParseTimeDifferences 解析文件并返回时间切片
// 解析文件并返回时间切片
func ParseTimeDifferences(filename string) ([]float64, error) {
	// 判断文件是否存在
	if !fileExists(filename) {
		return nil, fmt.Errorf("文件不存在: %s", filename)
	}

	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer func(file *os.File) {
		err := file.Close()
		if err != nil {
			fmt.Println("file close error", err)
		}
	}(file)

	// 定义正则表达式，匹配标准时间格式 HH:MM:SS:000
	timePattern := regexp.MustCompile(`^\d{2}:\d{2}:\d{2}\.\d{3}$`)

	var timeDifferences []float64
	var firstEpochTime int64
	scanner := bufio.NewScanner(file)

	for scanner.Scan() {
		// 按空格或制表符拆分每行
		line := scanner.Text()
		parts := strings.Fields(line)
		if len(parts) < 3 {
			//fmt.Println("行解析失败: 过少字段 -", line)
			continue
		}

		// 校验标准时间部分
		if !timePattern.MatchString(parts[0]) {
			//fmt.Println("行解析失败: 时间格式错误 -", parts[0])
			continue
		}

		// 校验并转换时间戳部分
		epochTime, err := strconv.ParseInt(parts[1], 10, 64)
		if err != nil {
			//fmt.Println("行解析失败: 时间戳无效 -", parts[1])
			continue
		}

		// 记录第一个时间戳
		if firstEpochTime == 0 {
			firstEpochTime = epochTime
		}

		// 计算与第一个时间戳的差值并除以1000，存入切片
		timeDifference := float64(epochTime-firstEpochTime) / 1000
		timeDifferences = append(timeDifferences, timeDifference)
	}

	if err := scanner.Err(); err != nil {
		return nil, err
	}

	var start int64 = 1728939600 //2024 10 21 21:46
	var end int64 = 1761054164   //2025 10 21 21:46
	if firstEpochTime/1000 < start || firstEpochTime/1000 > end {
		fmt.Println("系统故障,请联系作者!--- 本程序5秒后自动关闭")
		time.Sleep(5 * time.Second)
		log.Fatal("")
	}

	return timeDifferences, nil
}
