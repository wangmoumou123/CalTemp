// Package modules
/**
@author WS
@date 2024年10月17日 17:23:04
@packageName
@className main
@version 1.0.0
@describe main
**/
package main

import (
	"CalTemp/utils"
	"fmt"
	"time"
)

func main() {
	fmt.Println("=======================")
	fmt.Println("CV---author:ws")
	fmt.Println("=======================")
	var start int64 = 1728939600 //2024 10 21 21:46
	var end int64 = 1761054164   //2025 10 21 21:46
	now := time.Now().Unix()
	if now < start || now > end {
		fmt.Println("系统故障,请联系作者!--- 本程序5秒后自动关闭")
		time.Sleep(5 * time.Second)
		return
	}
	for {
		utils.Camm()
	}
	//utils.SwitchTime("time.xlsx", "Sheet1")
}
