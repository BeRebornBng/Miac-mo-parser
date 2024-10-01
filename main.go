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
	org                  = "Организация"
	link                 = "Ссылка"
	date                 = "Дата публикации"
	postsOne             = "Количество постов за 1 неделю"
	postsTwo             = "Количество постов за 2 неделю"
	postsThree           = "Количество постов за 3 неделю"
	postsFour            = "Количество постов за 4 неделю"
	postsLast            = "Количество постов за ост.дни"
	totalPostsFourWeek   = "Всего постов за 4 недели"
	postsAvgFourWeek     = "Постов в среднем за 4 недели"
	totalPostsMonth      = "Всего постов за месяц"
	likesOne             = "Количество лайков за 1 неделю"
	likesTwo             = "Количество лайков за 2 неделю"
	likesThree           = "Количество лайков за 3 неделю"
	likesFour            = "Количество лайков за 4 неделю"
	likesLast            = "Количество лайков за ост. дни"
	totalLikesFourWeek   = "Всего лайков за 4 недели"
	likesAvgFourWeek     = "Лайков в среднем за 4 недели"
	totalLikesMonth      = "Всего лайков за месяц"
	commsOne             = "Количество комментариев за 1 неделю"
	commsTwo             = "Количество комментариев за 2 неделю"
	commsThree           = "Количество комментариев за 3 неделю"
	commsFour            = "Количество комментариев за 4 неделю"
	commsLast            = "Количество комментариев за ост.дни"
	totalCommsFourWeek   = "Всего комментариев за 4 недели"
	commsAvgFourWeek     = "Комментариев в среднем за 4 недели"
	totalCommsMonth      = "Всего комментариев за месяц"
	repostsOne           = "Количество репостов за 1 неделю"
	repostsTwo           = "Количество репостов за 2 неделю"
	repostsThree         = "Количество репостов за 3 неделю"
	repostsFour          = "Количество репостов за 4 неделю"
	repostsLast          = "Количество репостов за ост.дни"
	totalRepostsFourWeek = "Всего репостов за 4 недели"
	repostsAvgFourWeek   = "Репостов в среднем за 4 недели"
	totalRepostsMonth    = "Всего репостов за месяц"
	viewsOne             = "Количество просмотров за 1 неделю"
	viewsTwo             = "Количество просмотров за 2 неделю"
	viewsThree           = "Количество просмотров за 3 неделю"
	viewsFour            = "Количество просмотров за 4 неделю"
	viewsLast            = "Количество просмотров за ост.дни"
	totalViewsFourWeek   = "Всего просмотров за 4 недели"
	viewsAvgFourWeek     = "Просмотров в среднем за 4 недели"
	totalViewsMonth      = "Всего просмотров за месяц"
)

const (
	filename = "Шаблон Госпаблики ВК.xlsx"
)

const (
	orgSheet     = "Организации"
	totalSheet   = "Общий свод"
	postsSheet   = "Посты свод"
	viewsSheet   = "Просмотры свод"
	repostsSheet = "Репосты свод"
	commsSheet   = "Комментарии свод"
	likesSheet   = "Лайки свод"
)

const (
	titleRow = 10
	valueRow = 11
)

func columnNumber(s string) int {
	n := len(s)
	result := 0

	for i := 0; i < n; i++ {
		result *= 26                // Сдвигаем результат на разряд влево (умножаем на 26)
		result += int(s[i]-'A') + 1 // Добавляем значение текущей буквы (A=1, B=2, ..., Z=26)
	}

	return result
}

func columnName(n int) string {
	name := ""
	for n > 0 {
		n-- // уменьшаем на 1 для корректного вычисления (A=1)
		remainder := n % 26
		name = string('A'+remainder) + name // добавляем букву в начало строки
		n /= 26                             // переходим к следующему разряду
	}
	return name
}

func main() {

	excelCfg := domain.NewExcelConfig(filename, []domain.SheetCells{{Sheet: totalSheet, Title: []string{org,
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
		totalViewsMonth}}, {Sheet: postsSheet, Title: []string{org,
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
	}}, {Sheet: likesSheet, Title: []string{org,
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
	}}, {Sheet: commsSheet, Title: []string{org,
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
	}}})

	var err error
	file, err := excelize.OpenFile(filename)
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
	orgs, err := utils.ExcelParseColumnCells(file, orgSheet, columnName(1), valueRow)
	if err != nil {
		log.Fatalln(err)
	}
	// получение ссылок из листа организаций
	links, err := utils.ExcelParseColumnCells(file, orgSheet, columnName(2), valueRow)
	if err != nil {
		log.Fatalln(err)
	}
	// получение доменов из ссылок
	domains, err := utils.DomainsFromLinks(links)
	if err != nil {
		log.Fatalln(err)
	}
	// все даты за год
	start := utils.StartNowMoth()
	end := utils.EndNowMonth()
	var dates [12][5]utils.Dates
	for i := 0; i < len(dates); i++ {
		newStart := start.AddDate(0, -1*i, 0)
		newEnd := end.AddDate(0, -1*i, 0)
		if newStart.Month() == time.February {
			newEnd = newEnd.AddDate(0, 0, -1)
		}
		dates[i] = utils.SplitMonth(newStart, newEnd)
	}

	row := valueRow
	for i, dom := range domains {
		for _, date := range dates {
			vkPosts, err := utils.GetVkPost(dom, date[0].Start, date[4].End)
			if err != nil {
				log.Fatalln(err)
			}
			vkCount := utils.VkCountInMonth(vkPosts.Response.Items, date)

			file.SetCellValue(totalSheet, "A1", orgs[i])
			file.SetCellValue(totalSheet, "A2", dom)
			file.SetCellValue(totalSheet, "A3", fmt.Sprintf("%s-%d", utils.MonthToRussian(date[0].Start.Month()), date[0].Start.Year()))
			for i, j := columnNumber("D"), 1; i <= columnNumber("AQ"); i, j = i+1, j+1 {
				if j == 4 {
					file.SetCellFormula(totalSheet, fmt.Sprintf("%s%d", columnName(i+1), row), fmt.Sprintf("SUM(%s%d:%s%d)", columnName(i-4)))
					file.SetCellFormula(totalSheet, fmt.Sprintf("%s%d", columnName(i+2), row), fmt.Sprintf("SUM(%s%d:%s%d)"))
					file.SetCellFormula(totalSheet, fmt.Sprintf("%s%d", columnName(i+3), row), fmt.Sprintf("SUM(%s%d:%s%d)"))
					j = -1
					i += 2
				}
				file.SetCellValue(totalSheet, fmt.Sprintf("%s%d", columnName(i), i), vkCount.PostsCount[i-1])
			}
		}
	}
	file.Save()

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
