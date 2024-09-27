package main

import (
	"fmt"
	"log"

	"github.com/Miac-mo-parser/domain"
	"github.com/Miac-mo-parser/utils"

	"github.com/xuri/excelize/v2"
)

var (
	fileName = "Шаблон Госпаблики ВК.xlsx"
	sheets   = []domain.SheetCells{
		{Name: "Организации", CellNames: []string{"A", "B"}, Row: 2},
		{Name: "Количество постов", CellNames: []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"}, Row: 4},
		{Name: "Количество лайков", CellNames: []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"}, Row: 4},
		{Name: "Количество комментариев", CellNames: []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"}, Row: 4},
		{Name: "Количество репостов", CellNames: []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"}, Row: 4},
		{Name: "Количество просмотров", CellNames: []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"}, Row: 4},
	}
)

func main() {

	var err error
	file, err := excelize.OpenFile(fileName)
	if err != nil {
		log.Fatalln(err)
	}

	orgs := domain.MedOrg{}
	orgs.Names, err = utils.ExcelParseColumnCells(file, sheets[0].Name, sheets[0].CellNames[0], sheets[0].Row)
	if err != nil {
		log.Fatalln(err)
	}
	orgs.Links, err = utils.ExcelParseColumnCells(file, sheets[0].Name, sheets[0].CellNames[1], sheets[0].Row)
	if err != nil {
		log.Fatalln(err)
	}

	domains, err := utils.DomainsFromLinks(orgs.Links)
	if err != nil {
		log.Fatalln(err)
	}

	dates := utils.SplitMonth()
	for i := 1; i < len(sheets); i++ {
		title := fmt.Sprintf(sheets[i].Name+" с %d по %d", dates[4].Start.Day(), dates[4].End.Day())
		file.SetCellStr(sheets[i].Name, fmt.Sprintf("%s%d", sheets[i].CellNames[6], sheets[i].Row-1), title)
	}
	for i, dom := range domains {
		vkPosts, err := utils.GetVkPost(dom, utils.StartNowMoth(), utils.EndNowMonth())
		if err != nil {
			log.Fatalln(err)
		}
		vkCount := utils.VkCountInMonth(vkPosts.Response.Items, dates)
		utils.ExcelSetCellsValue(file, sheets[1].Name, sheets[1].CellNames, sheets[1].Row, orgs.Names[i], orgs.Links[i], vkCount.PostsCount)
		utils.ExcelSetCellsValue(file, sheets[2].Name, sheets[2].CellNames, sheets[2].Row, orgs.Names[i], orgs.Links[i], vkCount.CommentsCount)
		utils.ExcelSetCellsValue(file, sheets[3].Name, sheets[3].CellNames, sheets[3].Row, orgs.Names[i], orgs.Links[i], vkCount.PostsCount)
		utils.ExcelSetCellsValue(file, sheets[4].Name, sheets[4].CellNames, sheets[4].Row, orgs.Names[i], orgs.Links[i], vkCount.RepostsCount)
		utils.ExcelSetCellsValue(file, sheets[5].Name, sheets[5].CellNames, sheets[5].Row, orgs.Names[i], orgs.Links[i], vkCount.LikesCount)
		sheets[1].Row++
		sheets[2].Row++
		sheets[3].Row++
		sheets[4].Row++
		sheets[5].Row++
	}
	for _, sheet := range sheets[1:] {
		file.SetCellStr(sheet.Name, fmt.Sprintf("%s%d", sheet.CellNames[0], sheet.Row), "Общий итог")
		file.SetCellFormula(sheet.Name, fmt.Sprintf("%s%d", sheet.CellNames[2], sheet.Row), fmt.Sprintf("SUM(%s%d:%s%d)", sheet.CellNames[2], 4, sheet.CellNames[2], sheet.Row-1))
		file.SetCellFormula(sheet.Name, fmt.Sprintf("%s%d", sheet.CellNames[3], sheet.Row), fmt.Sprintf("SUM(%s%d:%s%d)", sheet.CellNames[3], 4, sheet.CellNames[3], sheet.Row-1))
		file.SetCellFormula(sheet.Name, fmt.Sprintf("%s%d", sheet.CellNames[4], sheet.Row), fmt.Sprintf("SUM(%s%d:%s%d)", sheet.CellNames[4], 4, sheet.CellNames[4], sheet.Row-1))
		file.SetCellFormula(sheet.Name, fmt.Sprintf("%s%d", sheet.CellNames[5], sheet.Row), fmt.Sprintf("SUM(%s%d:%s%d)", sheet.CellNames[5], 4, sheet.CellNames[5], sheet.Row-1))
		file.SetCellFormula(sheet.Name, fmt.Sprintf("%s%d", sheet.CellNames[6], sheet.Row), fmt.Sprintf("SUM(%s%d:%s%d)", sheet.CellNames[6], 4, sheet.CellNames[6], sheet.Row-1))
		file.SetCellFormula(sheet.Name, fmt.Sprintf("%s%d", sheet.CellNames[7], sheet.Row), fmt.Sprintf("SUM(%s%d:%s%d)", sheet.CellNames[7], 4, sheet.CellNames[7], sheet.Row-1))
		file.SetCellFormula(sheet.Name, fmt.Sprintf("%s%d", sheet.CellNames[8], sheet.Row), fmt.Sprintf("SUM(%s%d:%s%d)", sheet.CellNames[8], 4, sheet.CellNames[8], sheet.Row-1))
		file.SetCellFormula(sheet.Name, fmt.Sprintf("%s%d", sheet.CellNames[9], sheet.Row), fmt.Sprintf("SUM(%s%d:%s%d)", sheet.CellNames[9], 4, sheet.CellNames[9], sheet.Row-1))
	}
	file.Save()
	// file.SaveAs(fmt.Sprintf(fileName+"_"+"%s", utils.StartNowMoth().Format(time.DateOnly)))
}
