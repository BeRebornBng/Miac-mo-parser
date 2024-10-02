package main

import (
	"fmt"
	"log"
	"strconv"
	"time"

	"github.com/Miac-mo-parser/domain"
	"github.com/Miac-mo-parser/utils"
	"github.com/Miac-mo-parser/utils/dates"
	"github.com/Miac-mo-parser/utils/excel"
	vkClient "github.com/Miac-mo-parser/utils/vk"
	"github.com/xuri/excelize/v2"
)

// const (
// 	templateFileName = "Шаблон Госпаблики ВК.xlsx"
// 	reportFileName   = "Госпаблики ВК.xlsx"
// )

// const (
// 	orgSheet     = "Организации"
// 	totalSheet   = "Общий_свод"
// 	postsSheet   = "Посты_свод"
// 	viewsSheet   = "Просмотры_свод"
// 	repostsSheet = "Репосты_свод"
// 	commsSheet   = "Комментарии_свод"
// 	likesSheet   = "Лайки_свод"
// )

// const (
// 	titleRow = 10
// 	valueRow = 11
// )

// const (
// 	org                  = "Организация"
// 	link                 = "Ссылка"
// 	date                 = "Дата публикации"
// 	postsOne             = "Количество постов за 1 неделю"
// 	postsTwo             = "Количество постов за 2 неделю"
// 	postsThree           = "Количество постов за 3 неделю"
// 	postsFour            = "Количество постов за 4 неделю"
// 	postsLast            = "Количество постов за ост.дни"
// 	totalPostsFourWeek   = "Всего постов за 4 недели"
// 	postsAvgFourWeek     = "Постов в среднем за 4 недели"
// 	totalPostsMonth      = "Всего постов за месяц"
// 	likesOne             = "Количество лайков за 1 неделю"
// 	likesTwo             = "Количество лайков за 2 неделю"
// 	likesThree           = "Количество лайков за 3 неделю"
// 	likesFour            = "Количество лайков за 4 неделю"
// 	likesLast            = "Количество лайков за ост. дни"
// 	totalLikesFourWeek   = "Всего лайков за 4 недели"
// 	likesAvgFourWeek     = "Лайков в среднем за 4 недели"
// 	totalLikesMonth      = "Всего лайков за месяц"
// 	commsOne             = "Количество комментариев за 1 неделю"
// 	commsTwo             = "Количество комментариев за 2 неделю"
// 	commsThree           = "Количество комментариев за 3 неделю"
// 	commsFour            = "Количество комментариев за 4 неделю"
// 	commsLast            = "Количество комментариев за ост.дни"
// 	totalCommsFourWeek   = "Всего комментариев за 4 недели"
// 	commsAvgFourWeek     = "Комментариев в среднем за 4 недели"
// 	totalCommsMonth      = "Всего комментариев за месяц"
// 	repostsOne           = "Количество репостов за 1 неделю"
// 	repostsTwo           = "Количество репостов за 2 неделю"
// 	repostsThree         = "Количество репостов за 3 неделю"
// 	repostsFour          = "Количество репостов за 4 неделю"
// 	repostsLast          = "Количество репостов за ост.дни"
// 	totalRepostsFourWeek = "Всего репостов за 4 недели"
// 	repostsAvgFourWeek   = "Репостов в среднем за 4 недели"
// 	totalRepostsMonth    = "Всего репостов за месяц"
// 	viewsOne             = "Количество просмотров за 1 неделю"
// 	viewsTwo             = "Количество просмотров за 2 неделю"
// 	viewsThree           = "Количество просмотров за 3 неделю"
// 	viewsFour            = "Количество просмотров за 4 неделю"
// 	viewsLast            = "Количество просмотров за ост.дни"
// 	totalViewsFourWeek   = "Всего просмотров за 4 недели"
// 	viewsAvgFourWeek     = "Просмотров в среднем за 4 недели"
// 	totalViewsMonth      = "Всего просмотров за месяц"
// )

func main() {

	templateCfg := domain.NewExcelConfig("Шаблон Госпаблики ВК.xlsx", []domain.SheetCells{
		{Sheet: "Организации", Titles: []string{"Наименование", "Ссылка", "Кол-во месяцев"}},
	})

	reportCfg := domain.NewExcelConfig("Госпаблики ВК.xlsx", []domain.SheetCells{
		{Sheet: "Общий_свод", Titles: []string{"Организация",
			"Ссылка",
			"Дата публикации",
			"Количество постов за 1 неделю",
			"Количество постов за 2 неделю",
			"Количество постов за 3 неделю",
			"Количество постов за 4 неделю",
			"Количество постов за ост.дни",
			"Всего постов за 4 недели",
			"Постов в среднем за 4 недели",
			"Всего постов за месяц",
			"Количество лайков за 1 неделю",
			"Количество лайков за 2 неделю",
			"Количество лайков за 3 неделю",
			"Количество лайков за 4 неделю",
			"Количество лайков за ост. дни",
			"Всего лайков за 4 недели",
			"Лайков в среднем за 4 недели",
			"Всего лайков за месяц",
			"Количество комментариев за 1 неделю",
			"Количество комментариев за 2 неделю",
			"Количество комментариев за 3 неделю",
			"Количество комментариев за 4 неделю",
			"Количество комментариев за ост.дни",
			"Всего комментариев за 4 недели",
			"Комментариев в среднем за 4 недели",
			"Всего комментариев за месяц",
			"Количество репостов за 1 неделю",
			"Количество репостов за 2 неделю",
			"Количество репостов за 3 неделю",
			"Количество репостов за 4 неделю",
			"Количество репостов за ост.дни",
			"Всего репостов за 4 недели",
			"Репостов в среднем за 4 недели",
			"Всего репостов за месяц",
			"Количество просмотров за 1 неделю",
			"Количество просмотров за 2 неделю",
			"Количество просмотров за 3 неделю",
			"Количество просмотров за 4 неделю",
			"Количество просмотров за ост.дни",
			"Всего просмотров за 4 недели",
			"Просмотров в среднем за 4 недели",
			"Всего просмотров за месяц"}},
		{Sheet: "Посты_свод", Titles: []string{"Организация",
			"Ссылка",
			"Дата публикации",
			"Количество постов за 1 неделю",
			"Количество постов за 2 неделю",
			"Количество постов за 3 неделю",
			"Количество постов за 4 неделю",
			"Количество постов за ост.дни",
			"Всего постов за 4 недели",
			"Постов в среднем за 4 недели",
			"Всего постов за месяц",
		}},
		{Sheet: "Лайки_свод", Titles: []string{"Организация",
			"Ссылка",
			"Дата публикации",
			"Количество лайков за 1 неделю",
			"Количество лайков за 2 неделю",
			"Количество лайков за 3 неделю",
			"Количество лайков за 4 неделю",
			"Количество лайков за ост. дни",
			"Всего лайков за 4 недели",
			"Лайков в среднем за 4 недели",
			"Всего лайков за месяц",
		}},
		{Sheet: "Репосты_свод", Titles: []string{"Организация",
			"Ссылка",
			"Дата публикации",
			"Количество репостов за 1 неделю",
			"Количество репостов за 2 неделю",
			"Количество репостов за 3 неделю",
			"Количество репостов за 4 неделю",
			"Количество репостов за ост.дни",
			"Всего репостов за 4 недели",
			"Репостов в среднем за 4 недели",
			"Всего репостов за месяц",
		}},
		{Sheet: "Комментарии_свод", Titles: []string{"Организация",
			"Ссылка",
			"Дата публикации",
			"Количество комментариев за 1 неделю",
			"Количество комментариев за 2 неделю",
			"Количество комментариев за 3 неделю",
			"Количество комментариев за 4 неделю",
			"Количество комментариев за ост.дни",
			"Всего комментариев за 4 недели",
			"Комментариев в среднем за 4 недели",
			"Всего комментариев за месяц",
		}},
		{Sheet: "Просмотры_свод", Titles: []string{"Организация",
			"Ссылка",
			"Дата публикации",
			"Количество просмотров за 1 неделю",
			"Количество просмотров за 2 неделю",
			"Количество просмотров за 3 неделю",
			"Количество просмотров за 4 неделю",
			"Количество просмотров за ост.дни",
			"Всего просмотров за 4 недели",
			"Просмотров в среднем за 4 недели",
			"Всего просмотров за месяц",
		}},
	})

	var err error
	file, err := excelize.OpenFile(templateCfg.FileName)
	if err != nil {
		log.Fatalln(err)
	}
	monthStrCount, err := file.GetCellValue(templateCfg.SheetCells[0].Sheet, "C2")
	if err != nil {
		log.Fatalln(err)
	}
	monthCount, err := strconv.Atoi(monthStrCount)
	if err != nil {
		log.Fatalln(err)
	}
	if monthCount <= 0 && monthCount >= 12 {
		log.Fatalln("количество месяцев > 0 и <= 12")
	}
	// получение названий МО из листа организаций
	orgs, err := excel.ExcelParseColumnCells(file, templateCfg.SheetCells[0].Sheet, "A", 2)
	if err != nil {
		log.Fatalln(err)
	}
	// получение ссылок из листа организаций
	links, err := excel.ExcelParseColumnCells(file, templateCfg.SheetCells[0].Sheet, "B", 2)
	if err != nil {
		log.Fatalln(err)
	}
	// получение доменов из ссылок
	domains, err := utils.DomainsFromLinks(links)
	if err != nil {
		log.Fatalln(err)
	}
	fmt.Println(domains, orgs)

	file = excelize.NewFile()
	for _, rs := range reportCfg.SheetCells {
		file.NewSheet(rs.Sheet)
		file.SetSheetRow(rs.Sheet, "A15", &rs.Titles)
	}
	// Создаем стиль для заголовков
	titleStyle := excelize.Style{
		Font: &excelize.Font{
			Family: "Calibri",
			Size:   12,
			Color:  "#000000", // Черный цвет
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
			WrapText:   true, // Перенос слов
		},
	}
	// Добавляем стили в файл Excel и получаем их ID
	titleStyleID, err := file.NewStyle(&titleStyle)
	if err != nil {
		panic(err)
	}
	// Создаем стиль для значений
	valueStyle := excelize.Style{
		Font: &excelize.Font{
			Family: "Calibri",
			Size:   12,
			Color:  "#000000", // Черный цвет
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
			WrapText:   true, // Перенос слов
		},
	}
	// Добавляем стили в файл Excel и получаем их ID
	valueStyleID, err := file.NewStyle(&valueStyle)
	if err != nil {
		panic(err)
	}
	fmt.Println(titleStyleID, valueStyleID)
	// все даты за период
	start := dates.StartNowMoth()
	end := dates.EndNowMonth()
	var monthBorders [][5]dates.MonthBorders
	for i := 0; i < monthCount; i++ {
		newStart := start.AddDate(0, -1*i, 0)
		newEnd := end.Add(time.Second).AddDate(0, -1*i, 0).Add(-time.Second)
		monthBorders = append(monthBorders, dates.SplitMonth(newStart, newEnd))
	}

	//row := 16
	for _, dom := range domains {
		for _, date := range monthBorders {
			vkPosts, err := vkClient.GetVkPost(dom, date[0].Start, date[4].End)
			if err != nil {
				log.Fatalln(err)
			}
			_ = vkClient.VkCountInMonth(vkPosts.Response.Items, date)
			for _, rs := range reportCfg.SheetCells {
				file.SetSheetRow(rs.Sheet, "A1", &rs.Titles)
			}
		}
	}

	// row := valueRow
	// for i, dom := range domains {
	// 	for _, date := range monthBorders {
	// 		vkPosts, err := vkClient.GetVkPost(dom, date[0].Start, date[4].End)
	// 		if err != nil {
	// 			log.Fatalln(err)
	// 		}
	// 		vkCount := vkClient.VkCountInMonth(vkPosts.Response.Items, date)
	// 		var values []float64
	// 		for j := 0; j < 8; j++ {
	// 			values = append(values, vkCount.PostsCount[j])
	// 		}
	// 		for j := 0; j < 8; j++ {
	// 			values = append(values, vkCount.LikesCount[j])
	// 		}
	// 		for j := 0; j < 8; j++ {
	// 			values = append(values, vkCount.CommentsCount[j])
	// 		}
	// 		for j := 0; j < 8; j++ {
	// 			values = append(values, vkCount.RepostsCount[j])
	// 		}
	// 		for j := 0; j < 8; j++ {
	// 			values = append(values, vkCount.ViewsCount[j])
	// 		}
	// 		for t, j := excel.ColumnNumber("D"), 0; t <= excel.ColumnNumber("AQ"); t, j = t+1, j+1 {
	// 			file.SetCellValue(totalSheet, fmt.Sprintf("%s%d", excel.ColumnName(t), row), values[j])
	// 			file.SetCellValue(totalSheet, fmt.Sprintf("%s%d", excel.ColumnName(1), row), orgs[i])
	// 			file.SetCellValue(totalSheet, fmt.Sprintf("%s%d", excel.ColumnName(2), row), links[i])
	// 			month, err := dates.MonthToRussian(date[0].Start.Month())
	// 			if err != nil {
	// 				log.Fatalln(err)
	// 			}
	// 			file.SetCellValue(totalSheet, fmt.Sprintf("%s%d", excel.ColumnName(3), row), fmt.Sprintf("%s-%d", month, date[0].Start.Year()))
	// 		}
	// 		for t, j := excel.ColumnNumber("D"), 0; t <= excel.ColumnNumber("K"); t, j = t+1, j+1 {
	// 			file.SetCellValue(postsSheet, fmt.Sprintf("%s%d", excel.ColumnName(t), row), vkCount.PostsCount[j])
	// 			file.SetCellValue(postsSheet, fmt.Sprintf("%s%d", excel.ColumnName(1), row), orgs[i])
	// 			file.SetCellValue(postsSheet, fmt.Sprintf("%s%d", excel.ColumnName(2), row), links[i])
	// 			month, err := dates.MonthToRussian(date[0].Start.Month())
	// 			if err != nil {
	// 				log.Fatalln(err)
	// 			}
	// 			file.SetCellValue(postsSheet, fmt.Sprintf("%s%d", excel.ColumnName(3), row), fmt.Sprintf("%s-%d", month, date[0].Start.Year()))
	// 		}
	// 		for t, j := excel.ColumnNumber("D"), 0; t <= excel.ColumnNumber("K"); t, j = t+1, j+1 {
	// 			file.SetCellValue(likesSheet, fmt.Sprintf("%s%d", excel.ColumnName(t), row), vkCount.LikesCount[j])
	// 			file.SetCellValue(likesSheet, fmt.Sprintf("%s%d", excel.ColumnName(1), row), orgs[i])
	// 			file.SetCellValue(likesSheet, fmt.Sprintf("%s%d", excel.ColumnName(2), row), links[i])
	// 			month, err := dates.MonthToRussian(date[0].Start.Month())
	// 			if err != nil {
	// 				log.Fatalln(err)
	// 			}
	// 			file.SetCellValue(likesSheet, fmt.Sprintf("%s%d", excel.ColumnName(3), row), fmt.Sprintf("%s-%d", month, date[0].Start.Year()))
	// 		}
	// 		for t, j := excel.ColumnNumber("D"), 0; t <= excel.ColumnNumber("K"); t, j = t+1, j+1 {
	// 			file.SetCellValue(commsSheet, fmt.Sprintf("%s%d", excel.ColumnName(t), row), vkCount.CommentsCount[j])
	// 			file.SetCellValue(commsSheet, fmt.Sprintf("%s%d", excel.ColumnName(1), row), orgs[i])
	// 			file.SetCellValue(commsSheet, fmt.Sprintf("%s%d", excel.ColumnName(2), row), links[i])
	// 			month, err := dates.MonthToRussian(date[0].Start.Month())
	// 			if err != nil {
	// 				log.Fatalln(err)
	// 			}
	// 			file.SetCellValue(commsSheet, fmt.Sprintf("%s%d", excel.ColumnName(3), row), fmt.Sprintf("%s-%d", month, date[0].Start.Year()))
	// 		}
	// 		for t, j := excel.ColumnNumber("D"), 0; t <= excel.ColumnNumber("K"); t, j = t+1, j+1 {
	// 			file.SetCellValue(repostsSheet, fmt.Sprintf("%s%d", excel.ColumnName(t), row), vkCount.RepostsCount[j])
	// 			file.SetCellValue(repostsSheet, fmt.Sprintf("%s%d", excel.ColumnName(1), row), orgs[i])
	// 			file.SetCellValue(repostsSheet, fmt.Sprintf("%s%d", excel.ColumnName(2), row), links[i])
	// 			month, err := dates.MonthToRussian(date[0].Start.Month())
	// 			if err != nil {
	// 				log.Fatalln(err)
	// 			}
	// 			file.SetCellValue(repostsSheet, fmt.Sprintf("%s%d", excel.ColumnName(3), row), fmt.Sprintf("%s-%d", month, date[0].Start.Year()))
	// 		}
	// 		for t, j := excel.ColumnNumber("D"), 0; t <= excel.ColumnNumber("K"); t, j = t+1, j+1 {
	// 			file.SetCellValue(viewsSheet, fmt.Sprintf("%s%d", excel.ColumnName(t), row), vkCount.ViewsCount[j])
	// 			file.SetCellValue(viewsSheet, fmt.Sprintf("%s%d", excel.ColumnName(1), row), orgs[i])
	// 			file.SetCellValue(viewsSheet, fmt.Sprintf("%s%d", excel.ColumnName(2), row), links[i])
	// 			month, err := dates.MonthToRussian(date[0].Start.Month())
	// 			if err != nil {
	// 				log.Fatalln(err)
	// 			}
	// 			file.SetCellValue(viewsSheet, fmt.Sprintf("%s%d", excel.ColumnName(3), row), fmt.Sprintf("%s-%d", month, date[0].Start.Year()))
	// 		}
	// 		// Устанавливаем высоту для строк
	// 		height := 30.0
	// 		if err := file.SetRowHeight(totalSheet, row, height); err != nil {
	// 			panic(err)
	// 		}
	// 		if err := file.SetRowHeight(postsSheet, row, height); err != nil {
	// 			panic(err)
	// 		}
	// 		if err := file.SetRowHeight(likesSheet, row, height); err != nil {
	// 			panic(err)
	// 		}
	// 		if err := file.SetRowHeight(repostsSheet, row, height); err != nil {
	// 			panic(err)
	// 		}
	// 		if err := file.SetRowHeight(commsSheet, row, height); err != nil {
	// 			panic(err)
	// 		}
	// 		if err := file.SetRowHeight(viewsSheet, row, height); err != nil {
	// 			panic(err)
	// 		}
	// 		row++
	// 	}
	// }

	// file.SetCellStyle(totalSheet, fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("A")), 1), fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("AQ")), row), styleID1)
	// // Устанавливаем ширину для колонок A и B
	// if err := file.SetColWidth(totalSheet, "A", "A", 60); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(totalSheet, "B", "B", 40); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(totalSheet, "C", "C", 30); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(totalSheet, "D", "AQ", 20); err != nil {
	// 	panic(err)
	// }
	// file.SetCellStyle(postsSheet, fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("A")), 1), fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("AQ")), row), styleID1)
	// // Устанавливаем ширину для колонок A и B
	// if err := file.SetColWidth(postsSheet, "A", "A", 60); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(postsSheet, "B", "B", 40); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(postsSheet, "C", "C", 30); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(postsSheet, "D", "AQ", 20); err != nil {
	// 	panic(err)
	// }
	// file.SetCellStyle(likesSheet, fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("A")), 1), fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("AQ")), row), styleID1)
	// // Устанавливаем ширину для колонок A и B
	// if err := file.SetColWidth(likesSheet, "A", "A", 60); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(likesSheet, "B", "B", 40); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(likesSheet, "C", "C", 30); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(likesSheet, "D", "AQ", 20); err != nil {
	// 	panic(err)
	// }
	// file.SetCellStyle(viewsSheet, fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("A")), 1), fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("AQ")), row), styleID1)
	// // Устанавливаем ширину для колонок A и B
	// if err := file.SetColWidth(viewsSheet, "A", "A", 60); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(viewsSheet, "B", "B", 40); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(viewsSheet, "C", "C", 30); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(viewsSheet, "D", "AQ", 20); err != nil {
	// 	panic(err)
	// }

	// file.SetCellStyle(repostsSheet, fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("A")), 1), fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("AQ")), row), styleID1)
	// // Устанавливаем ширину для колонок A и B
	// if err := file.SetColWidth(repostsSheet, "A", "A", 60); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(repostsSheet, "B", "B", 40); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(repostsSheet, "C", "C", 30); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(repostsSheet, "D", "AQ", 20); err != nil {
	// 	panic(err)
	// }
	// file.SetCellStyle(commsSheet, fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("A")), 1), fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("AQ")), row), styleID1)
	// // Устанавливаем ширину для колонок A и B
	// if err := file.SetColWidth(commsSheet, "A", "A", 60); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(commsSheet, "B", "B", 40); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(commsSheet, "C", "C", 20); err != nil {
	// 	panic(err)
	// }
	// if err := file.SetColWidth(commsSheet, "D", "AQ", 20); err != nil {
	// 	panic(err)
	// }

	// sheetsArr := []string{totalSheet, likesSheet, postsSheet, repostsSheet, commsSheet, viewsSheet}
	// for i, sheet := range sheetsArr {
	// 	tableName := fmt.Sprintf("%s_%d", "Фильтр_даты_публикации", i)
	// 	tableRange := fmt.Sprintf("%s%d:%s%d", excel.ColumnName(3), titleRow, excel.ColumnName(3), row-1)
	// 	table := excelize.Table{
	// 		Name:              tableName,
	// 		Range:             tableRange,
	// 		ShowFirstColumn:   false,
	// 		ShowLastColumn:    false,
	// 		ShowColumnStripes: false,
	// 	}
	// 	if err := file.AddTable(sheet, &table); err != nil {
	// 		log.Fatalf("failed to add table: %v", err)
	// 	}

	// 	tableName2 := fmt.Sprintf("%s_%d", "Фильтр_организаций", i)
	// 	tableRange2 := fmt.Sprintf("%s%d:%s%d", excel.ColumnName(1), titleRow, excel.ColumnName(1), row-1)
	// 	table2 := excelize.Table{
	// 		Name:              tableName2,
	// 		Range:             tableRange2,
	// 		ShowFirstColumn:   false,
	// 		ShowLastColumn:    false,
	// 		ShowColumnStripes: false,
	// 	}
	// 	if err := file.AddTable(sheet, &table2); err != nil {
	// 		log.Fatalf("failed to add table: %v", err)
	// 	}
	// }
	// series := []excelize.ChartSeries{}
	// for i := 11; i < row-1; i++ {
	// 	series = append(series, excelize.ChartSeries{Name: "'" + postsSheet + "'" + fmt.Sprintf("!$C$%d", i), Categories: "'" + postsSheet + "'" + fmt.Sprintf("!$C$%d", i), Values: postsSheet + fmt.Sprintf("!$J$%d", i)})
	// }
	// err = file.AddChart(postsSheet, "A1", &excelize.Chart{
	// 	Type:   excelize.Col3DClustered,
	// 	Title:  []excelize.RichTextRun{{Text: "Наываыва"}},
	// 	Series: series,
	// })
	// if err != nil {
	// 	log.Fatalln(err)
	// }
	file.SaveAs(reportCfg.FileName)
	if err := file.DeleteSheet("Sheet1"); err != nil {
		log.Println("Sheet 1 not founded")
	}
	file.SaveAs(reportCfg.FileName)
	log.Println("Отчет успешно сохранен в файл " + reportCfg.FileName)
}
