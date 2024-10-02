package excel

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

// Получить строки колонки с книги
func ExcelParseColumnCells(file *excelize.File, sheet string, cellName string, row int) ([]string, error) {
	links := make([]string, 0)
	for {
		value, err := file.GetCellValue(sheet, fmt.Sprintf("%s%d", cellName, row))
		row++
		if err != nil {
			log.Fatalf("Ошибка при чтении ячейки %s: %s", value, err)
		}
		if value == "" {
			break
		}
		links = append(links, value)
	}
	return links, nil
}

// Установить значения для строк колонки
func ExcelSetCellsValue(file *excelize.File, sheet string, cellName []string, row int, orgName string, orgLink string, counts [5]int) {
	file.SetCellStr(sheet, fmt.Sprintf("%s%d", cellName[0], row), orgName)
	file.SetCellStr(sheet, fmt.Sprintf("%s%d", cellName[1], row), orgLink)
	for i := 0; i < 5; i++ {
		file.SetCellInt(sheet, fmt.Sprintf("%s%d", cellName[i+2], row), counts[i])
	}
	file.SetCellFormula(sheet, fmt.Sprintf("%s%d", cellName[7], row), fmt.Sprintf("SUM(%s%d:%s%d)/4", cellName[2], row, cellName[5], row))
	file.SetCellFormula(sheet, fmt.Sprintf("%s%d", cellName[8], row), fmt.Sprintf("SUM(%s%d:%s%d)", cellName[2], row, cellName[5], row))
	file.SetCellFormula(sheet, fmt.Sprintf("%s%d", cellName[9], row), fmt.Sprintf("SUM(%s%d:%s%d)", cellName[2], row, cellName[6], row))
}

func ColumnNumber(s string) int {
	n := len(s)
	result := 0

	for i := 0; i < n; i++ {
		result *= 26                // Сдвигаем результат на разряд влево (умножаем на 26)
		result += int(s[i]-'A') + 1 // Добавляем значение текущей буквы (A=1, B=2, ..., Z=26)
	}

	return result
}

func ColumnName(n int) string {
	name := ""
	for n > 0 {
		n-- // уменьшаем на 1 для корректного вычисления (A=1)
		remainder := n % 26
		name = string('A'+remainder) + name // добавляем букву в начало строки
		n /= 26                             // переходим к следующему разряду
	}
	return name
}
