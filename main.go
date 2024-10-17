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
	fileName := fmt.Sprintf("ws_%s.xlsx", time.Now().Format("2006_01_02"))

	fmt.Println("=======================")
	fmt.Println("cal_temp---author:ws")
	fmt.Println("=======================")

	for {
		utils.Camm(fileName)
	}
}
