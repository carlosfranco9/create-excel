package main

import (
	"flag"
	"fmt"
	"github.com/mritd/chinaid"
	"github.com/xuri/excelize/v2"
	"math/rand"
	"sync"
)

var (
	size    int
	runType int
	num     int
	mx      sync.Mutex
)

func main() {

	flag.IntVar(&size, "a", 10, "字符大小")
	flag.IntVar(&num, "n", 1, "行数")
	flag.IntVar(&runType, "t", 2, "类型 1- 生成固定大小的excel文件 2- 生成银行卡，身份证等信息的excel文件")

	flag.Parse()

	switch runType {
	case 1:
		exportRandomStrExcel()
	case 2:
		exportIdentificationInformation()
	}

}

func exportIdentificationInformation() {
	savePath := "身份证信息.xlsx"

	createExcel(savePath, func(f *excelize.File) {
		wait := sync.WaitGroup{}
		f.SetCellValue("Sheet2", "A1", "姓名")
		f.SetCellValue("Sheet2", "B1", "身份证号")
		f.SetCellValue("Sheet2", "C1", "银行卡号")
		f.SetCellValue("Sheet2", "D1", "手机号")
		for i := 1; i <= num; i++ {
			wait.Add(1)
			information := createInformation()
			col := i + 1
			go insertInfo(f, col, &information, &wait)
		}
		wait.Wait()
	})
}

func insertInfo(file *excelize.File, col int, info *Info, wg *sync.WaitGroup) {
	fmt.Println(col)
	mx.Lock()
	file.SetCellValue("Sheet2", fmt.Sprintf("A%d", col), info.Name)
	file.SetCellValue("Sheet2", fmt.Sprintf("B%d", col), info.Id)
	file.SetCellValue("Sheet2", fmt.Sprintf("C%d", col), info.Bank)
	file.SetCellValue("Sheet2", fmt.Sprintf("D%d", col), info.Mobile)

	mx.Unlock()
	wg.Done()
}

type Info struct {
	Name   string
	Id     string
	Bank   string
	Mobile string
}

func createInformation() Info {
	return Info{
		Name:   chinaid.Name(),
		Id:     chinaid.IDNo(),
		Bank:   chinaid.BankNo(),
		Mobile: chinaid.Mobile(),
	}
}

func exportRandomStrExcel() {

	savePath := "固定大小.xlsx"

	createExcel(savePath, func(f *excelize.File) {
		size = size * 1024 * 1024 * 5

		// 设置单元格的值
		slice := 100 * 1024

		peak := size / slice

		wait := sync.WaitGroup{}

		for i := 1; i <= peak; i++ {
			wait.Add(1)
			go insertContent(f, fmt.Sprintf("A%d", i), slice, &wait)
		}
		wait.Wait()
	})
}

func insertContent(file *excelize.File, position string, slice int, wg *sync.WaitGroup) {
	str := createRandomStrBySize(slice)
	file.SetCellValue("Sheet2", position, str)
	wg.Done()
}

func createRandomStrBySize(n int) string {
	str := "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
	bytes := []byte(str)
	var result []byte
	for i := 0; i < n; i++ {
		result = append(result, bytes[rand.Intn(len(bytes))])
	}
	return string(result)
}

func createExcel(filePath string, process func(file *excelize.File)) {
	f := excelize.NewFile()

	index := f.NewSheet("Sheet2")

	f.SetActiveSheet(index)

	process(f)

	// 根据指定路径保存文件
	if err := f.SaveAs(filePath); err != nil {
		fmt.Println(err)
	}
}
