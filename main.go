package main

import (
	"fmt"
	"log"
	"time"

	"github.com/Miac-mo-parser/domain"
	"github.com/Miac-mo-parser/utils"
	"github.com/xuri/excelize/v2"
)

func mustConfig() string {
	return ""
}

const (
	Org = iota
	Link
	Date
	PostsOne
	PostsTwo
	PostsThree
	PostsFour
	PostsLast
	LikesOne
	LikesTwo
	LikesThree
	LikesFour
	LikesLast
	CommsOne
	CommsTwo
	CommsThree
	CommsFour
	CommsLast
	RepostsOne
	RepostsTwo
	RepostsThree
	RepostsFour
	RepostsLast
	ViewsOne
	ViewsTwo
	ViewsThree
	ViewsFour
	ViewsLast
)

const (
	titleRow = 1
	valueRow = 2
)

var (
	filename = "Шаблон Госпаблики ВК.xlsx"
	sheet1   = "За_весь_период"
	cells    = []domain.Cells{
		{Title: "Организация", CellName: "A"},
		{Title: "Ссылка", CellName: "B"},
		{Title: "Дата публикации", CellName: "C"},
		{Title: "Количество постов за 1 неделю", CellName: "D"},
		{Title: "Количество постов за 2 неделю", CellName: "E"},
		{Title: "Количество постов за 3 неделю", CellName: "F"},
		{Title: "Количество постов за 4 неделю", CellName: "G"},
		{Title: "Количество постов за ост.дни", CellName: "H"},
		{Title: "Количество лайков за 1 неделю", CellName: "I"},
		{Title: "Количество лайков за 2 неделю", CellName: "J"},
		{Title: "Количество лайков за 3 неделю", CellName: "K"},
		{Title: "Количество лайков за 4 неделю", CellName: "L"},
		{Title: "Количество лайков за ост. дни", CellName: "M"},
		{Title: "Количество комментариев за 1 неделю", CellName: "N"},
		{Title: "Количество комментариев за 2 неделю", CellName: "O"},
		{Title: "Количество комментариев за 3 неделю", CellName: "P"},
		{Title: "Количество комментариев за 4 неделю", CellName: "Q"},
		{Title: "Количество комментариев за ост.дни", CellName: "R"},
		{Title: "Количество репостов за 1 неделю", CellName: "S"},
		{Title: "Количество репостов за 2 неделю", CellName: "T"},
		{Title: "Количество репостов за 3 неделю", CellName: "U"},
		{Title: "Количество репостов за 4 неделю", CellName: "V"},
		{Title: "Количество репостов за ост.дни", CellName: "W"},
		{Title: "Количество просмотров за 1 неделю", CellName: "X"},
		{Title: "Количество просмотров за 2 неделю", CellName: "Y"},
		{Title: "Количество просмотров за 3 неделю", CellName: "Z"},
		{Title: "Количество просмотров за 4 неделю", CellName: "AA"},
		{Title: "Количество просмотров за ост.дни", CellName: "AB"},
	}
)

func main() {

	excelCfg := domain.NewExcelConfig(filename, sheet1, cells)

	var err error
	file, err := excelize.OpenFile(excelCfg.FileName)
	if err != nil {
		log.Fatalln(err)
	}

	for _, cell := range excelCfg.Cells {
		file.SetCellStr(excelCfg.Sheet, fmt.Sprintf("%s%d", cell.CellName, titleRow), cell.Title)
	}

	// получение ссылок из excel
	links, err := utils.ExcelParseColumnCells(file, excelCfg.Sheet, excelCfg.Cells[Link].CellName, valueRow)
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
	for i := 0; i < 12; i++ {
		newStart := start.AddDate(0, -1*i, 0)
		newEnd := end.AddDate(0, -1*i, 0)
		if newStart.Month() == time.February {
			newEnd = newEnd.AddDate(0, 0, -1)
		}
		dates := utils.SplitMonth(newStart, newEnd)
		for _, dom := range domains {
			vkPosts, err := utils.GetVkPost(dom, dates[0].Start, dates[4].End)
			if err != nil {
				log.Fatalln(err)
			}
			vkCount := utils.VkCountInMonth(vkPosts.Response.Items, dates)
			for _, cell := range excelCfg.Cells {
				//var arr [5]int
				switch cell.Title {
				case excelCfg.Cells[Date].Title:
					file.SetCellStr(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[Date].CellName, row), fmt.Sprintf("%s-%d", utils.MonthToRussian(dates[0].Start.Month()), dates[0].Start.Year()))
					break
				case excelCfg.Cells[PostsOne].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[PostsOne].CellName, row), vkCount.PostsCount[0])
					break
				case excelCfg.Cells[PostsTwo].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[PostsTwo].CellName, row), vkCount.PostsCount[1])
					break
				case excelCfg.Cells[PostsThree].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[PostsThree].CellName, row), vkCount.PostsCount[2])
					break
				case excelCfg.Cells[PostsFour].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[PostsFour].CellName, row), vkCount.PostsCount[3])
					break
				case excelCfg.Cells[PostsLast].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[PostsLast].CellName, row), vkCount.PostsCount[4])
					break
				case excelCfg.Cells[LikesOne].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[LikesOne].CellName, row), vkCount.LikesCount[0])
					break
				case excelCfg.Cells[LikesTwo].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[LikesTwo].CellName, row), vkCount.LikesCount[1])
					break
				case excelCfg.Cells[LikesThree].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[LikesThree].CellName, row), vkCount.LikesCount[2])
					break
				case excelCfg.Cells[LikesFour].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[LikesFour].CellName, row), vkCount.LikesCount[3])
					break
				case excelCfg.Cells[LikesLast].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[LikesLast].CellName, row), vkCount.LikesCount[4])
					break
				case excelCfg.Cells[CommsOne].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[CommsOne].CellName, row), vkCount.CommentsCount[0])
					break
				case excelCfg.Cells[CommsTwo].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[CommsTwo].CellName, row), vkCount.CommentsCount[1])
					break
				case excelCfg.Cells[CommsThree].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[CommsThree].CellName, row), vkCount.CommentsCount[2])
					break
				case excelCfg.Cells[CommsFour].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[CommsFour].CellName, row), vkCount.CommentsCount[3])
					break
				case excelCfg.Cells[CommsLast].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[CommsLast].CellName, row), vkCount.CommentsCount[4])
					break
				case excelCfg.Cells[RepostsOne].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[RepostsOne].CellName, row), vkCount.RepostsCount[0])
					break
				case excelCfg.Cells[RepostsTwo].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[RepostsTwo].CellName, row), vkCount.RepostsCount[1])
					break
				case excelCfg.Cells[RepostsThree].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[RepostsThree].CellName, row), vkCount.RepostsCount[2])
					break
				case excelCfg.Cells[RepostsFour].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[RepostsFour].CellName, row), vkCount.RepostsCount[3])
					break
				case excelCfg.Cells[RepostsLast].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[RepostsLast].CellName, row), vkCount.RepostsCount[4])
					break
				case excelCfg.Cells[ViewsOne].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[ViewsOne].CellName, row), vkCount.ViewsCount[0])
					break
				case excelCfg.Cells[ViewsTwo].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[ViewsTwo].CellName, row), vkCount.ViewsCount[1])
					break
				case excelCfg.Cells[ViewsThree].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[ViewsThree].CellName, row), vkCount.ViewsCount[2])
					break
				case excelCfg.Cells[ViewsFour].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[ViewsFour].CellName, row), vkCount.ViewsCount[3])
					break
				case excelCfg.Cells[ViewsLast].Title:
					file.SetCellInt(excelCfg.Sheet, fmt.Sprintf("%s%d", excelCfg.Cells[ViewsLast].CellName, row), vkCount.ViewsCount[4])
					break
				}
			}
			row++
		}
	}
	file.NewSheet("Общий свод")
	index, err := file.GetSheetIndex(excelCfg.Sheet)
	if err != nil {
		log.Fatalln(err)
	}
	fmt.Println(index)
	index, err = file.GetSheetIndex("Общий свод")
	if err != nil {
		log.Fatalln(err)
	}
	fmt.Println(index)

	// tableName := "Фильтр_даты_публикации"
	// tableRange := fmt.Sprintf("%s%d:%s%d", excelCfg.Cells[Date].CellName, titleRow, excelCfg.Cells[Date].CellName, row-1)
	// table := excelize.Table{
	// 	Name:              tableName,
	// 	Range:             tableRange,
	// 	ShowFirstColumn:   false,
	// 	ShowLastColumn:    false,
	// 	ShowColumnStripes: false,
	// }
	// if err := file.AddTable(excelCfg.Sheet, &table); err != nil {
	// 	log.Fatalf("failed to add table: %v", err)
	// }

	// err = file.AddSlicer("Свод", &excelize.SlicerOptions{
	// 	Name:       "Дата публикации",
	// 	Cell:       "H4",
	// 	TableSheet: "За_весь_период",
	// 	TableName:  "Фильтр_даты_публикации",
	// 	Width:      200,
	// 	Height:     200,
	// })
	// if err != nil {
	// 	log.Fatalln(err)
	// }
	file.Save()
	// file.SaveAs(fmt.Sprintf(fileName+"_"+"%s", utils.StartNowMoth().Format(time.DateOnly)))
}
