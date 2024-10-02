package main

import (
	"fmt"
	"log"
	"time"

	"github.com/Miac-mo-parser/domain"
	"github.com/Miac-mo-parser/utils"
	"github.com/Miac-mo-parser/utils/dates"
	"github.com/Miac-mo-parser/utils/excel"
	"github.com/xuri/excelize/v2"
)

func main() {

	excelCfg := domain.NewExcelConfig(filename, []domain.SheetCells{{Sheet: totalSheet, Title: []string{
		org,
		link,
		date,
		postsOne,
		postsTwo,
		postsThree,
		postsFour,
		postsLast,
		totalPostsFourWeek,
		postsAvgFourWeek,
		totalPostsMonth,
		likesOne,
		likesTwo,
		likesThree,
		likesFour,
		likesLast,
		totalLikesFourWeek,
		likesAvgFourWeek,
		totalLikesMonth,
		commsOne,
		commsTwo,
		commsThree,
		commsFour,
		commsLast,
		totalCommsFourWeek,
		commsAvgFourWeek,
		totalCommsMonth,
		repostsOne,
		repostsTwo,
		repostsThree,
		repostsFour,
		repostsLast,
		totalRepostsFourWeek,
		repostsAvgFourWeek,
		totalRepostsMonth,
		viewsOne,
		viewsTwo,
		viewsThree,
		viewsFour,
		viewsLast,
		totalViewsFourWeek,
		viewsAvgFourWeek,
		totalViewsMonth}}, {Sheet: postsSheet, Title: []string{
		org,
		link,
		date,
		postsOne,
		postsTwo,
		postsThree,
		postsFour,
		postsLast,
		totalPostsFourWeek,
		postsAvgFourWeek,
		totalPostsMonth,
	}}, {Sheet: likesSheet, Title: []string{
		org,
		link,
		date,
		likesOne,
		likesTwo,
		likesThree,
		likesFour,
		likesLast,
		totalLikesFourWeek,
		likesAvgFourWeek,
		totalLikesMonth,
	}}, {Sheet: commsSheet, Title: []string{
		org,
		link,
		date,
		commsOne,
		commsTwo,
		commsThree,
		commsFour,
		commsLast,
		totalCommsFourWeek,
		commsAvgFourWeek,
		totalCommsMonth,
	}},
		{Sheet: repostsSheet, Title: []string{
			org,
			link,
			date,
			repostsOne,
			repostsTwo,
			repostsThree,
			repostsFour,
			repostsLast,
			totalRepostsFourWeek,
			repostsAvgFourWeek,
			totalRepostsMonth,
		}},
		{Sheet: viewsSheet, Title: []string{
			org,
			link,
			date,
			viewsOne,
			viewsTwo,
			viewsThree,
			viewsFour,
			viewsLast,
			totalViewsFourWeek,
			viewsAvgFourWeek,
			totalViewsMonth}}})

	var err error
	file, err := excelize.OpenFile(excelCfg.FileName)
	if err != nil {
		log.Fatalln(err)
	}

	// // удалить листы
	// for _, sCell := range excelCfg.SheetCells[1:] {
	// 	if err := file.DeleteSheet(sCell.Sheet); err != nil {
	// 		log.Println(err)
	// 	}
	// }

	// // создать листы
	// for _, sCell := range excelCfg.SheetCells[1:] {
	// 	if _, err := file.NewSheet(sCell.Sheet); err != nil {
	// 		log.Println(err)
	// 	}
	// }

	// // установить заголовки для листов
	// for _, sCell := range excelCfg.SheetCells[1:] {
	// 	for _, cell := range sCell.Cells {
	// 		file.SetCellStr(sCell.Sheet, fmt.Sprintf("%s%d", cell.CellName, titleRow), cell.Title)
	// 	}
	// }

	// получение названий МО из листа организаций
	_, err = excel.ExcelParseColumnCells(file, orgSheet, columnName(1), valueRow)
	if err != nil {
		log.Fatalln(err)
	}
	// получение ссылок из листа организаций
	links, err := excel.ExcelParseColumnCells(file, orgSheet, columnName(2), valueRow)
	if err != nil {
		log.Fatalln(err)
	}
	// получение доменов из ссылок
	_, err = utils.DomainsFromLinks(links)
	if err != nil {
		log.Fatalln(err)
	}
	// все даты за год
	start := dates.StartNowMoth()
	end := dates.EndNowMonth()
	var dates [12][5]dates.MonthBorders
	for i := 0; i < len(dates); i++ {
		newStart := start.AddDate(0, -1*i, 0)
		newEnd := end.Add(time.Second).AddDate(0, -1*i, 0).Add(-time.Second)
		fmt.Println(newStart, newEnd)
	}

	// row := valueRow
	// for i, dom := range domains {
	// 	for _, date := range dates {
	// 		vkPosts, err := utils.GetVkPost(dom, date[0].Start, date[4].End)
	// 		if err != nil {
	// 			log.Fatalln(err)
	// 		}
	// 		vkCount := utils.VkCountInMonth(vkPosts.Response.Items, date)

	// 		file.SetCellValue(totalSheet, "A1", orgs[i])
	// 		file.SetCellValue(totalSheet, "A2", dom)
	// 		file.SetCellValue(totalSheet, "A3", fmt.Sprintf("%s-%d", utils.MonthToRussian(date[0].Start.Month()), date[0].Start.Year()))
	// 		for i, j := columnNumber("D"), 1; i <= columnNumber("AQ"); i, j = i+1, j+1 {
	// 			if j == 4 {
	// 				file.SetCellFormula(totalSheet, fmt.Sprintf("%s%d", columnName(i+1), row), fmt.Sprintf("SUM(%s%d:%s%d)", columnName(i-4)))
	// 				file.SetCellFormula(totalSheet, fmt.Sprintf("%s%d", columnName(i+2), row), fmt.Sprintf("SUM(%s%d:%s%d)"))
	// 				file.SetCellFormula(totalSheet, fmt.Sprintf("%s%d", columnName(i+3), row), fmt.Sprintf("SUM(%s%d:%s%d)"))
	// 				j = -1
	// 				i += 2
	// 			}
	// 			file.SetCellValue(totalSheet, fmt.Sprintf("%s%d", columnName(i), i), vkCount.PostsCount[i-1])
	// 		}
	// 	}
	// }
	// file.Save()

	// row := valueRow
	// for _, dom := range domains {
	// 	for _, date := range dates {
	// 		vkPosts, err := utils.GetVkPost(dom, date[0].Start, date[4].End)
	// 		if err != nil {
	// 			log.Fatalln(err)
	// 		}
	// 		vkCount := utils.VkCountInMonth(vkPosts.Response.Items, date)
	// 		for _, sCell := range excelCfg.SheetCells[1:] {
	// 			switch sCell.Sheet {
	// 			case excelCfg.SheetCells[TotalSheet].Sheet:
	// 				for _, cell := range sCell.Cells {
	// 					switch cell.Title {
	// 					case excelCfg.SheetCells[TotalSheet].Cells[Date].Title:
	// 						file.SetCellStr(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[Date].CellName, row), fmt.Sprintf("%s-%d", utils.MonthToRussian(date[0].Start.Month()), date[0].Start.Year()))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[PostsOne].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[PostsOne].CellName, row), vkCount.PostsCount[0])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[PostsTwo].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[PostsTwo].CellName, row), vkCount.PostsCount[1])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[PostsThree].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[PostsThree].CellName, row), vkCount.PostsCount[2])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[PostsFour].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[PostsFour].CellName, row), vkCount.PostsCount[3])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[PostsLast].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[PostsLast].CellName, row), vkCount.PostsCount[4])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[TotalPostsFourWeek].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[TotalPostsFourWeek].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[TotalSheet].Cells[PostsOne].CellName, row, excelCfg.SheetCells[TotalSheet].Cells[PostsFour].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[PostsAvgFourWeek].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[PostsAvgFourWeek].CellName, row), fmt.Sprintf("%s%d/4", excelCfg.SheetCells[TotalSheet].Cells[TotalPostsFourWeek].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[TotalPostsMonth].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[TotalPostsMonth].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[TotalSheet].Cells[PostsOne].CellName, row, excelCfg.SheetCells[TotalSheet].Cells[PostsLast].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[LikesOne].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[LikesOne].CellName, row), vkCount.LikesCount[0])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[LikesTwo].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[LikesTwo].CellName, row), vkCount.LikesCount[1])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[LikesThree].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[LikesThree].CellName, row), vkCount.LikesCount[2])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[LikesFour].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[LikesFour].CellName, row), vkCount.LikesCount[3])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[LikesLast].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[LikesLast].CellName, row), vkCount.LikesCount[4])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[TotalLikesFourWeek].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[TotalLikesFourWeek].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[TotalSheet].Cells[LikesOne].CellName, row, excelCfg.SheetCells[TotalSheet].Cells[LikesFour].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[LikesAvgFourWeek].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[LikesAvgFourWeek].CellName, row), fmt.Sprintf("%s%d/4", excelCfg.SheetCells[TotalSheet].Cells[TotalLikesFourWeek].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[TotalLikesMonth].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[TotalLikesMonth].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[TotalSheet].Cells[LikesOne].CellName, row, excelCfg.SheetCells[TotalSheet].Cells[LikesLast].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[CommsOne].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[CommsOne].CellName, row), vkCount.CommentsCount[0])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[CommsTwo].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[CommsTwo].CellName, row), vkCount.CommentsCount[1])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[CommsThree].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[CommsThree].CellName, row), vkCount.CommentsCount[2])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[CommsFour].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[CommsFour].CellName, row), vkCount.CommentsCount[3])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[CommsLast].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[CommsLast].CellName, row), vkCount.CommentsCount[4])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[TotalCommsFourWeek].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[TotalCommsFourWeek].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[TotalSheet].Cells[CommsOne].CellName, row, excelCfg.SheetCells[TotalSheet].Cells[CommsFour].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[CommsAvgFourWeek].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[CommsAvgFourWeek].CellName, row), fmt.Sprintf("%s%d/4", excelCfg.SheetCells[TotalSheet].Cells[TotalCommsFourWeek].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[TotalCommsMonth].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[TotalCommsMonth].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[TotalSheet].Cells[CommsOne].CellName, row, excelCfg.SheetCells[TotalSheet].Cells[CommsLast].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[RepostsOne].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[RepostsOne].CellName, row), vkCount.RepostsCount[0])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[RepostsTwo].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[RepostsTwo].CellName, row), vkCount.RepostsCount[1])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[RepostsThree].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[RepostsThree].CellName, row), vkCount.RepostsCount[2])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[RepostsFour].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[RepostsFour].CellName, row), vkCount.RepostsCount[3])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[RepostsLast].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[RepostsLast].CellName, row), vkCount.RepostsCount[4])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[TotalRepostsFourWeek].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[TotalRepostsFourWeek].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[TotalSheet].Cells[RepostsOne].CellName, row, excelCfg.SheetCells[TotalSheet].Cells[RepostsFour].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[RepostsAvgFourWeek].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[RepostsAvgFourWeek].CellName, row), fmt.Sprintf("%s%d/4", excelCfg.SheetCells[TotalSheet].Cells[TotalRepostsFourWeek].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[TotalRepostsMonth].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[TotalRepostsMonth].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[TotalSheet].Cells[RepostsOne].CellName, row, excelCfg.SheetCells[TotalSheet].Cells[RepostsLast].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[ViewsOne].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[ViewsOne].CellName, row), vkCount.ViewsCount[0])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[ViewsTwo].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[ViewsTwo].CellName, row), vkCount.ViewsCount[1])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[ViewsThree].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[ViewsThree].CellName, row), vkCount.ViewsCount[2])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[ViewsFour].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[ViewsFour].CellName, row), vkCount.ViewsCount[3])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[ViewsLast].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[ViewsLast].CellName, row), vkCount.ViewsCount[4])
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[TotalViewsFourWeek].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[TotalViewsFourWeek].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[TotalSheet].Cells[ViewsOne].CellName, row, excelCfg.SheetCells[TotalSheet].Cells[ViewsFour].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[ViewsAvgFourWeek].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[ViewsAvgFourWeek].CellName, row), fmt.Sprintf("%s%d/4", excelCfg.SheetCells[TotalSheet].Cells[TotalViewsFourWeek].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[TotalSheet].Cells[TotalViewsMonth].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[TotalSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[TotalSheet].Cells[TotalViewsMonth].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[TotalSheet].Cells[ViewsOne].CellName, row, excelCfg.SheetCells[TotalSheet].Cells[ViewsLast].CellName, row))
	// 						break
	// 					}
	// 				}
	// 				break
	// 			case excelCfg.SheetCells[PostsSheet].Sheet:
	// 				for _, cell := range sCell.Cells {
	// 					switch cell.Title {
	// 					case excelCfg.SheetCells[PostsSheet].Cells[Date].Title:
	// 						file.SetCellStr(excelCfg.SheetCells[PostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[PostsSheet].Cells[Date].CellName, row), fmt.Sprintf("%s-%d", utils.MonthToRussian(date[0].Start.Month()), date[0].Start.Year()))
	// 						break
	// 					case excelCfg.SheetCells[PostsSheet].Cells[PostsOne].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[PostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[PostsSheet].Cells[PostsOne].CellName, row), vkCount.PostsCount[0])
	// 						break
	// 					case excelCfg.SheetCells[PostsSheet].Cells[PostsTwo].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[PostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[PostsSheet].Cells[PostsTwo].CellName, row), vkCount.PostsCount[1])
	// 						break
	// 					case excelCfg.SheetCells[PostsSheet].Cells[PostsThree].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[PostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[PostsSheet].Cells[PostsThree].CellName, row), vkCount.PostsCount[2])
	// 						break
	// 					case excelCfg.SheetCells[PostsSheet].Cells[PostsFour].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[PostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[PostsSheet].Cells[PostsFour].CellName, row), vkCount.PostsCount[3])
	// 						break
	// 					case excelCfg.SheetCells[PostsSheet].Cells[PostsLast].Title:
	// 						file.SetCellInt(excelCfg.SheetCells[PostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[PostsSheet].Cells[PostsLast].CellName, row), vkCount.PostsCount[4])
	// 						break
	// 					case excelCfg.SheetCells[PostsSheet].Cells[TotalPostsFourWeek].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[PostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[PostsSheet].Cells[TotalPostsFourWeek].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[PostsSheet].Cells[PostsOne].CellName, row, excelCfg.SheetCells[PostsSheet].Cells[PostsFour].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[PostsSheet].Cells[PostsAvgFourWeek].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[PostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[PostsSheet].Cells[PostsAvgFourWeek].CellName, row), fmt.Sprintf("%s%d/4", excelCfg.SheetCells[PostsSheet].Cells[TotalPostsFourWeek].CellName, row))
	// 						break
	// 					case excelCfg.SheetCells[PostsSheet].Cells[TotalPostsMonth].Title:
	// 						file.SetCellFormula(excelCfg.SheetCells[PostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[PostsSheet].Cells[TotalPostsMonth].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[PostsSheet].Cells[PostsOne].CellName, row, excelCfg.SheetCells[PostsSheet].Cells[PostsLast].CellName, row))
	// 						break
	// 					}
	// 				}
	// 				break
	// 			case excelCfg.SheetCells[LikesSheet].Sheet:
	// 				for _, cell := range sCell.Cells {
	// 					switch cell.Title {
	// 					case excelCfg.SheetCells[LikesSheet].Cells[Date].Title:
	// 						file.SetCellStr(excelCfg.SheetCells[LikesSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[LikesSheet].Cells[Date].CellName, row), fmt.Sprintf("%s-%d", utils.MonthToRussian(date[0].Start.Month()), date[0].Start.Year()))
	// 						break
	// 						// case excelCfg.SheetCells[LikesSheet].Cells[LikesOne].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[LikesSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[LikesSheet].Cells[LikesOne].CellName, row), vkCount.LikesCount[0])
	// 						// 	break
	// 						// case excelCfg.SheetCells[LikesSheet].Cells[LikesTwo].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[LikesSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[LikesSheet].Cells[LikesTwo].CellName, row), vkCount.LikesCount[1])
	// 						// 	break
	// 						// case excelCfg.SheetCells[LikesSheet].Cells[LikesThree].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[LikesSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[LikesSheet].Cells[LikesThree].CellName, row), vkCount.LikesCount[2])
	// 						// 	break
	// 						// case excelCfg.SheetCells[LikesSheet].Cells[LikesFour].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[LikesSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[LikesSheet].Cells[LikesFour].CellName, row), vkCount.LikesCount[3])
	// 						// 	break
	// 						// case excelCfg.SheetCells[LikesSheet].Cells[LikesLast].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[LikesSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[LikesSheet].Cells[LikesLast].CellName, row), vkCount.LikesCount[4])
	// 						// 	break
	// 						// case excelCfg.SheetCells[LikesSheet].Cells[TotalLikesFourWeek].Title:
	// 						// 	file.SetCellFormula(excelCfg.SheetCells[LikesSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[LikesSheet].Cells[TotalLikesFourWeek].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[LikesSheet].Cells[LikesOne].CellName, row, excelCfg.SheetCells[LikesSheet].Cells[LikesFour].CellName, row))
	// 						// 	break
	// 						// case excelCfg.SheetCells[LikesSheet].Cells[LikesAvgFourWeek].Title:
	// 						// 	file.SetCellFormula(excelCfg.SheetCells[LikesSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[LikesSheet].Cells[LikesAvgFourWeek].CellName, row), fmt.Sprintf("%s%d/4", excelCfg.SheetCells[LikesSheet].Cells[TotalLikesFourWeek].CellName, row))
	// 						// 	break
	// 						// case excelCfg.SheetCells[LikesSheet].Cells[TotalLikesMonth].Title:
	// 						// 	file.SetCellFormula(excelCfg.SheetCells[LikesSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[LikesSheet].Cells[TotalLikesMonth].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[LikesSheet].Cells[LikesOne].CellName, row, excelCfg.SheetCells[LikesSheet].Cells[LikesLast].CellName, row))
	// 						// 	break
	// 					}
	// 				}
	// 				break
	// 			case excelCfg.SheetCells[CommsSheet].Sheet:
	// 				for _, cell := range sCell.Cells {
	// 					switch cell.Title {
	// 					case excelCfg.SheetCells[LikesSheet].Cells[Date].Title:
	// 						file.SetCellStr(excelCfg.SheetCells[LikesSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[LikesSheet].Cells[Date].CellName, row), fmt.Sprintf("%s-%d", utils.MonthToRussian(date[0].Start.Month()), date[0].Start.Year()))
	// 						break
	// 						// case excelCfg.SheetCells[CommsSheet].Cells[CommsOne].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[CommsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[CommsSheet].Cells[CommsOne].CellName, row), vkCount.CommentsCount[0])
	// 						// 	break
	// 						// case excelCfg.SheetCells[CommsSheet].Cells[CommsTwo].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[CommsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[CommsSheet].Cells[CommsTwo].CellName, row), vkCount.CommentsCount[1])
	// 						// 	break
	// 						// case excelCfg.SheetCells[CommsSheet].Cells[CommsThree].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[CommsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[CommsSheet].Cells[CommsThree].CellName, row), vkCount.CommentsCount[2])
	// 						// 	break
	// 						// case excelCfg.SheetCells[CommsSheet].Cells[CommsFour].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[CommsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[CommsSheet].Cells[CommsFour].CellName, row), vkCount.CommentsCount[3])
	// 						// 	break
	// 						// case excelCfg.SheetCells[CommsSheet].Cells[CommsLast].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[CommsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[CommsSheet].Cells[CommsLast].CellName, row), vkCount.CommentsCount[4])
	// 						// 	break
	// 						// case excelCfg.SheetCells[CommsSheet].Cells[TotalCommsFourWeek].Title:
	// 						// 	file.SetCellFormula(excelCfg.SheetCells[CommsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[CommsSheet].Cells[TotalCommsFourWeek].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[CommsSheet].Cells[CommsOne].CellName, row, excelCfg.SheetCells[CommsSheet].Cells[CommsFour].CellName, row))
	// 						// 	break
	// 						// case excelCfg.SheetCells[CommsSheet].Cells[CommsAvgFourWeek].Title:
	// 						// 	file.SetCellFormula(excelCfg.SheetCells[CommsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[CommsSheet].Cells[CommsAvgFourWeek].CellName, row), fmt.Sprintf("%s%d/4", excelCfg.SheetCells[CommsSheet].Cells[TotalCommsFourWeek].CellName, row))
	// 						// 	break
	// 						// case excelCfg.SheetCells[CommsSheet].Cells[TotalCommsMonth].Title:
	// 						// 	file.SetCellFormula(excelCfg.SheetCells[CommsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[CommsSheet].Cells[TotalCommsMonth].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[CommsSheet].Cells[CommsOne].CellName, row, excelCfg.SheetCells[CommsSheet].Cells[CommsLast].CellName, row))
	// 						// 	break
	// 					}
	// 				}
	// 				break
	// 			case excelCfg.SheetCells[RepostsSheet].Sheet:
	// 				for _, cell := range sCell.Cells {
	// 					switch cell.Title {
	// 					case excelCfg.SheetCells[RepostsSheet].Cells[Date].Title:
	// 						file.SetCellStr(excelCfg.SheetCells[RepostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[RepostsSheet].Cells[Date].CellName, row), fmt.Sprintf("%s-%d", utils.MonthToRussian(date[0].Start.Month()), date[0].Start.Year()))
	// 						break
	// 						// case excelCfg.SheetCells[RepostsSheet].Cells[RepostsOne].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[RepostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[RepostsSheet].Cells[RepostsOne].CellName, row), vkCount.RepostsCount[0])
	// 						// 	break
	// 						// case excelCfg.SheetCells[RepostsSheet].Cells[RepostsTwo].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[RepostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[RepostsSheet].Cells[RepostsTwo].CellName, row), vkCount.RepostsCount[1])
	// 						// 	break
	// 						// case excelCfg.SheetCells[RepostsSheet].Cells[RepostsThree].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[RepostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[RepostsSheet].Cells[RepostsThree].CellName, row), vkCount.RepostsCount[2])
	// 						// 	break
	// 						// case excelCfg.SheetCells[RepostsSheet].Cells[RepostsFour].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[RepostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[RepostsSheet].Cells[RepostsFour].CellName, row), vkCount.RepostsCount[3])
	// 						// 	break
	// 						// case excelCfg.SheetCells[RepostsSheet].Cells[RepostsLast].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[RepostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[RepostsSheet].Cells[RepostsLast].CellName, row), vkCount.RepostsCount[4])
	// 						// 	break
	// 						// case excelCfg.SheetCells[RepostsSheet].Cells[TotalRepostsFourWeek].Title:
	// 						// 	file.SetCellFormula(excelCfg.SheetCells[RepostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[RepostsSheet].Cells[TotalRepostsFourWeek].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[RepostsSheet].Cells[RepostsOne].CellName, row, excelCfg.SheetCells[RepostsSheet].Cells[RepostsFour].CellName, row))
	// 						// 	break
	// 						// case excelCfg.SheetCells[RepostsSheet].Cells[RepostsAvgFourWeek].Title:
	// 						// 	file.SetCellFormula(excelCfg.SheetCells[RepostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[RepostsSheet].Cells[RepostsAvgFourWeek].CellName, row), fmt.Sprintf("%s%d/4", excelCfg.SheetCells[RepostsSheet].Cells[TotalRepostsFourWeek].CellName, row))
	// 						// 	break
	// 						// case excelCfg.SheetCells[RepostsSheet].Cells[TotalRepostsMonth].Title:
	// 						// 	file.SetCellFormula(excelCfg.SheetCells[RepostsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[RepostsSheet].Cells[TotalRepostsMonth].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[RepostsSheet].Cells[RepostsOne].CellName, row, excelCfg.SheetCells[RepostsSheet].Cells[RepostsLast].CellName, row))
	// 						// 	break
	// 					}
	// 				}
	// 				break
	// 			case excelCfg.SheetCells[ViewsSheet].Sheet:
	// 				for _, cell := range sCell.Cells {
	// 					switch cell.Title {
	// 					case excelCfg.SheetCells[ViewsSheet].Cells[Date].Title:
	// 						file.SetCellStr(excelCfg.SheetCells[ViewsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[ViewsSheet].Cells[Date].CellName, row), fmt.Sprintf("%s-%d", utils.MonthToRussian(date[0].Start.Month()), date[0].Start.Year()))
	// 						break
	// 						// case excelCfg.SheetCells[ViewsSheet].Cells[ViewsOne].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[ViewsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[ViewsSheet].Cells[ViewsOne].CellName, row), vkCount.ViewsCount[0])
	// 						// 	break
	// 						// case excelCfg.SheetCells[ViewsSheet].Cells[ViewsTwo].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[ViewsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[ViewsSheet].Cells[ViewsTwo].CellName, row), vkCount.ViewsCount[1])
	// 						// 	break
	// 						// case excelCfg.SheetCells[ViewsSheet].Cells[ViewsThree].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[ViewsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[ViewsSheet].Cells[ViewsThree].CellName, row), vkCount.ViewsCount[2])
	// 						// 	break
	// 						// case excelCfg.SheetCells[ViewsSheet].Cells[ViewsFour].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[ViewsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[ViewsSheet].Cells[ViewsFour].CellName, row), vkCount.ViewsCount[3])
	// 						// 	break
	// 						// case excelCfg.SheetCells[ViewsSheet].Cells[ViewsLast].Title:
	// 						// 	file.SetCellInt(excelCfg.SheetCells[ViewsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[ViewsSheet].Cells[ViewsLast].CellName, row), vkCount.ViewsCount[4])
	// 						// 	break
	// 						// case excelCfg.SheetCells[ViewsSheet].Cells[TotalViewsFourWeek].Title:
	// 						// 	file.SetCellFormula(excelCfg.SheetCells[ViewsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[ViewsSheet].Cells[TotalViewsFourWeek].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[ViewsSheet].Cells[ViewsOne].CellName, row, excelCfg.SheetCells[ViewsSheet].Cells[ViewsFour].CellName, row))
	// 						// 	break
	// 						// case excelCfg.SheetCells[ViewsSheet].Cells[ViewsAvgFourWeek].Title:
	// 						// 	file.SetCellFormula(excelCfg.SheetCells[ViewsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[ViewsSheet].Cells[ViewsAvgFourWeek].CellName, row), fmt.Sprintf("%s%d/4", excelCfg.SheetCells[ViewsSheet].Cells[TotalViewsFourWeek].CellName, row))
	// 						// 	break
	// 						// case excelCfg.SheetCells[ViewsSheet].Cells[TotalViewsMonth].Title:
	// 						// 	file.SetCellFormula(excelCfg.SheetCells[ViewsSheet].Sheet, fmt.Sprintf("%s%d", excelCfg.SheetCells[ViewsSheet].Cells[TotalViewsMonth].CellName, row), fmt.Sprintf("SUM(%s%d:%s%d)", excelCfg.SheetCells[ViewsSheet].Cells[ViewsOne].CellName, row, excelCfg.SheetCells[ViewsSheet].Cells[ViewsLast].CellName, row))
	// 						// 	break
	// 					}
	// 				}
	// 				break
	// 			}
	// 		}
	// 		row++
	// 	}
	// }
	// // file.NewSheet("Общий свод")
	// // index1, err := file.GetSheetIndex(excelCfg.SheetCells[TotalSheet].Sheet)
	// // if err != nil {
	// // 	log.Fatalln(err)
	// // }
	// // index2, err := file.GetSheetIndex("Общий свод")
	// // if err != nil {
	// // 	log.Fatalln(err)
	// // }
	// // file.CopySheet(index1, index2)

	// // tableName := "Фильтр_даты_публикации"
	// // tableRange := fmt.Sprintf("%s%d:%s%d", excelCfg.SheetCells[TotalSheet].Cells[Date].CellName, titleRow, excelCfg.SheetCells[TotalSheet].Cells[Date].CellName, row-1)
	// // table := excelize.Table{
	// // 	Name:              tableName,
	// // 	Range:             tableRange,
	// // 	ShowFirstColumn:   false,
	// // 	ShowLastColumn:    false,
	// // 	ShowColumnStripes: false,
	// // }
	// // if err := file.AddTable("Общий свод", &table); err != nil {
	// // 	log.Fatalf("failed to add table: %v", err)
	// // }

	// // err = file.AddSlicer("Общий свод", &excelize.SlicerOptions{
	// // 	Name:          "Дата публикации",
	// // 	Cell:          "B1",
	// // 	TableSheet:    "Общий свод",
	// // 	TableName:     "Фильтр_даты_публикации",
	// // 	Width:         700,
	// // 	Height:        260,
	// // 	ItemDesc:      true,
	// // 	DisplayHeader: nil,
	// // })
	// // if err != nil {
	// // 	log.Fatalln(err)
	// // }
	// file.Save()
	// // file.SaveAs(fmt.Sprintf(fileName+"_"+"%s", utils.StartNowMoth().Format(time.DateOnly)))
}
