package main

import (
	"fmt"
	"log"
	"time"

	"github.com/Miac-mo-parser/domain"
	"github.com/Miac-mo-parser/utils"

	"github.com/xuri/excelize/v2"
)

const (
	OrgID = iota
	LinkID
	DataID
	PostsID
	LikesID
	CommentsID
	RepostsID
	ViewsID
)

func monthToRussian(month time.Month) string {
	switch month {
	case time.January:
		return "Январь"
	case time.February:
		return "Февраль"
	case time.March:
		return "Март"
	case time.April:
		return "Апрель"
	case time.May:
		return "Май"
	case time.June:
		return "Июнь"
	case time.July:
		return "Июль"
	case time.August:
		return "Август"
	case time.September:
		return "Сентябрь"
	case time.October:
		return "Октябрь"
	case time.November:
		return "Ноябрь"
	default:
		return "Декабрь"
	}
}

var (
	excelCfg = domain.ExcelConfig{
		FileName:  "Шаблон Госпаблики ВК.xlsx",
		SheetName: "За весь период",
		Cells: []domain.Cells{{Title: "Организация", CellsName: []string{"A"}},
			{
				Title:     "Ссылка",
				CellsName: []string{"B"},
			},
			{
				Title:     "Дата публикации",
				CellsName: []string{"C"},
			},
			{
				Title:     "Количество постов",
				CellsName: []string{"D", "E", "F", "G", "H"},
			},
			{
				Title:     "Количество лайков",
				CellsName: []string{"I", "J", "K", "L", "M"},
			},
			{
				Title:     "Количество комментариев",
				CellsName: []string{"N", "O", "P", "Q", "R"},
			},
			{
				Title:     "Количество репостов",
				CellsName: []string{"S", "T", "U", "V", "W"},
			},
			{
				Title:     "Количество просмотров",
				CellsName: []string{"X", "Y", "Z", "AA", "AB"},
			},
		},
	}
)

func main() {

	var err error
	file, err := excelize.OpenFile(excelCfg.FileName)
	if err != nil {
		log.Fatalln(err)
	}

	// получение ссылок из excel
	links, err := utils.ExcelParseColumnCells(file, excelCfg.SheetName, excelCfg.Cells[LinkID].CellsName[0], 2)
	if err != nil {
		log.Fatalln(err)
	}

	// получение доменов из ссылок
	domains, err := utils.DomainsFromLinks(links)
	if err != nil {
		log.Fatalln(err)
	}

	start := utils.StartNowMoth()
	end := utils.EndNowMonth()
	row := 2
	for i := 0; i < 2; i++ {
		dates := utils.SplitMonth(start.AddDate(0, -1*i, 0), end.AddDate(0, -1*i, 0))
		if i == 0 {
			for j := 0; j < 5; j++ {
				for _, cell := range excelCfg.Cells {
					title := fmt.Sprintf(cell.Title+" с %d по %d", dates[j].Start.Day(), dates[j].End.Day())
					for _, l := range cell.CellsName {
						file.SetCellStr(excelCfg.SheetName, fmt.Sprintf("%s%d", l, 1), title)
					}
				}
			}
		}

		for i, dom := range domains {
			vkPosts, err := utils.GetVkPost(dom, dates[0].Start, dates[4].End)
			if err != nil {
				log.Fatalln(err)
			}
			vkCount := utils.VkCountInMonth(vkPosts.Response.Items, dates)
			fmt.Println(i+1, dom, vkCount)
			for _, cell := range excelCfg.Cells {
				var arr [5]int
				switch cell.Title {
				case "Дата публикации":
					for _, name := range cell.CellsName {
						file.SetCellStr(excelCfg.SheetName, fmt.Sprintf("%s%d", name, row), fmt.Sprintf("%s-%d", monthToRussian(dates[0].Start.Month()), dates[0].Start.Year()))
					}
					continue
				case "Количество постов":
					arr = vkCount.PostsCount
					break
				case "Количество лайков":
					arr = vkCount.LikesCount
					break
				case "Количество комментариев":
					arr = vkCount.CommentsCount
					break
				case "Количество репостов":
					arr = vkCount.RepostsCount
					break
				case "Количество просмотров":
					arr = vkCount.ViewsCount
					break
				case "Организация":
					continue
				case "Ссылка":
					continue
				}
				for i, name := range cell.CellsName {
					file.SetCellInt(excelCfg.SheetName, fmt.Sprintf("%s%d", name, row), arr[i])
				}
			}
			row++
		}

	}
	file.AddPivotTable(&excelize.PivotTableOptions{DataRange: fmt.Sprintf("%s!%s%d:%s%d", excelCfg.SheetName, excelCfg.Cells[0].CellsName[0], 2, excelCfg.Cells[0].CellsName[0], row-1),
		PivotTableRange: fmt.Sprintf("Свод!A1:Z%d", row-1),
		Filter:          []excelize.PivotTableField{{Data: "Дата публикации"}}})
	file.Save()
	// file.SaveAs(fmt.Sprintf(fileName+"_"+"%s", utils.StartNowMoth().Format(time.DateOnly)))
}
