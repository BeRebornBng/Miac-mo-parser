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

const (
	orgSheet     = "Организации"
	totalSheet   = "Общий_свод"
	postsSheet   = "Посты_свод"
	viewsSheet   = "Просмотры_свод"
	repostsSheet = "Репосты_свод"
	commsSheet   = "Комментарии_свод"
	likesSheet   = "Лайки_свод"
)

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

	// смотрим шаблонный файл
	var err error
	file, err := excelize.OpenFile(templateCfg.FileName)
	if err != nil {
		log.Fatalln(err)
	}
	// получаем количество месяцев
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

	// новый файл
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
	// все даты за период
	start := dates.StartNowMoth()
	end := dates.EndNowMonth()
	var monthBorders [][5]dates.MonthBorders
	for i := 0; i < monthCount; i++ {
		newStart := start.AddDate(0, -1*i, 0)
		newEnd := end.Add(time.Second).AddDate(0, -1*i, 0).Add(-time.Second)
		monthBorders = append(monthBorders, dates.SplitMonth(newStart, newEnd))
	}

	row := 16
	for i, dom := range domains {
		fmt.Println(dom)
		for _, date := range monthBorders {

			vkPosts, err := vkClient.GetVkPost(dom, date[0].Start, date[4].End)
			if err != nil {
				log.Fatalln(err)
			}
			vkCount := vkClient.VkCountInMonth(vkPosts.Response.Items, date)
			for _, item := range vkPosts.Response.Items {
				time.Unix(int64(item.Date), 0)
			}
			for _, rs := range reportCfg.SheetCells {
				height := 30.0
				if err := file.SetRowHeight(rs.Sheet, row, height); err != nil {
					panic(err)
				}
				file.SetCellStyle(rs.Sheet, fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("A")), 1), fmt.Sprintf("%s%d", excel.ColumnName(excel.ColumnNumber("AQ")), row), titleStyleID)
				if err := file.SetColWidth(rs.Sheet, "A", "A", 60); err != nil {
					panic(err)
				}
				if err := file.SetColWidth(rs.Sheet, "B", "B", 40); err != nil {
					panic(err)
				}
				if err := file.SetColWidth(rs.Sheet, "C", "C", 30); err != nil {
					panic(err)
				}
				if err := file.SetColWidth(rs.Sheet, "D", "AQ", 20); err != nil {
					panic(err)
				}
				month, err := dates.MonthToRussian(date[0].Start.Month())
				if err != nil {
					log.Fatalln(err)
				}
				file.SetSheetRow(rs.Sheet, fmt.Sprintf("A%d", row), &[]interface{}{orgs[i], links[i], month})
				switch rs.Sheet {
				case totalSheet:
					file.SetSheetRow(rs.Sheet, fmt.Sprintf("A%d", row), &[]interface{}{orgs[i], links[i], fmt.Sprintf("%s-%d", month, date[0].Start.Year())})
					file.SetSheetRow(rs.Sheet, fmt.Sprintf("D%d", row), &vkCount.PostsCount)
					file.SetSheetRow(rs.Sheet, fmt.Sprintf("L%d", row), &vkCount.LikesCount)
					file.SetSheetRow(rs.Sheet, fmt.Sprintf("T%d", row), &vkCount.CommentsCount)
					file.SetSheetRow(rs.Sheet, fmt.Sprintf("AB%d", row), &vkCount.RepostsCount)
					file.SetSheetRow(rs.Sheet, fmt.Sprintf("AJ%d", row), &vkCount.ViewsCount)
					break
				case postsSheet:
					file.SetSheetRow(rs.Sheet, fmt.Sprintf("D%d", row), &vkCount.PostsCount)
					break
				case viewsSheet:
					file.SetSheetRow(rs.Sheet, fmt.Sprintf("D%d", row), &vkCount.ViewsCount)
					break
				case repostsSheet:
					file.SetSheetRow(rs.Sheet, fmt.Sprintf("D%d", row), &vkCount.RepostsCount)
					break
				case commsSheet:
					file.SetSheetRow(rs.Sheet, fmt.Sprintf("D%d", row), &vkCount.CommentsCount)
					break
				case likesSheet:
					file.SetSheetRow(rs.Sheet, fmt.Sprintf("D%d", row), &vkCount.LikesCount)
					break
				}
			}
			row++
		}
	}

	file.NewSheet("Динамика")

	for _, rs := range reportCfg.SheetCells {
		err := file.AutoFilter(rs.Sheet, fmt.Sprintf("A%d:AQ%d", 15, row-1), nil)
		if err != nil {
			log.Fatalln(err)
		}
		// series := []excelize.ChartSeries{}
		// for i := 15; i < row-1; i = i + monthCount {
		// 	if i+monthCount > row-1 {
		// 		break
		// 	}
		// 	series = append(series, excelize.ChartSeries{Name: "'" + rs.Sheet + "'" + fmt.Sprintf("!$C$%d:$C$%d", i, i+monthCount), Categories: "'" + rs.Sheet + "'" + fmt.Sprintf("!$C$%d:$C$%d", i, i+monthCount), Values: rs.Sheet + fmt.Sprintf("!$J$%d:$J$%d", i, i+monthCount)})
		// }
		// err = file.AddChart(rs.Sheet, "A1", &excelize.Chart{
		// 	Type:   excelize.Col3DClustered,
		// 	Title:  []excelize.RichTextRun{{Text: "В среднем лайков"}},
		// 	Series: series,
		// })
	}

	if err := file.AddPivotTable(&excelize.PivotTableOptions{
		Name:            "Свод",
		DataRange:       fmt.Sprintf("%s!A%d:AQ%d", totalSheet, 15, 15+monthCount),
		PivotTableRange: fmt.Sprintf("%s!A%d:AQ%d", "Динамика", 15, 15+monthCount),
		Rows:            []excelize.PivotTableField{{Name: "Дата публикации", Data: "Дата публикации", DefaultSubtotal: true}},
		Filter:          []excelize.PivotTableField{{Name: "Выбор организации", Data: "Организация", DefaultSubtotal: true}},
		//Columns:         []excelize.PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data: []excelize.PivotTableField{{Data: "Постов в среднем за 4 недели", Name: fmt.Sprintf("Постов в среднем за %d месяцев", monthCount)},
			{Data: "Лайков в среднем за 4 недели", Name: fmt.Sprintf("Лайков в среднем за %d месяцев", monthCount)},
			{Data: "Репостов в среднем за 4 недели", Name: fmt.Sprintf("Репостов в среднем за %d месяцев", monthCount)},
			{Data: "Комментариев в среднем за 4 недели", Name: fmt.Sprintf("Комментариев в среднем за %d месяцев", monthCount)},
			{Data: "Просмотров в среднем за 4 недели", Name: fmt.Sprintf("Просмотров в среднем за %d месяцев", monthCount)}},
		RowGrandTotals: true,
		ColGrandTotals: true,
		ShowDrill:      true,
		ShowRowHeaders: true,
		ShowColHeaders: true,
		ShowLastColumn: true,
	}); err != nil {
		fmt.Println(err)
	}

	err = file.AddChart("Динамика", "H1", &excelize.Chart{
		Type:   excelize.Line,
		Title:  []excelize.RichTextRun{{Text: fmt.Sprintf("Постов в среднем за %d месяцев", monthCount)}},
		Series: []excelize.ChartSeries{{Name: "", Categories: fmt.Sprintf("'Динамика'!$A$15:$A$%d", 15+monthCount), Values: fmt.Sprintf("'Динамика'!$B$15:$B$%d", 15+monthCount)}},
	})
	err = file.AddChart("Динамика", "H15", &excelize.Chart{
		Type:   excelize.Line,
		Title:  []excelize.RichTextRun{{Text: fmt.Sprintf("Лайков в среднем за %d месяцев", monthCount)}},
		Series: []excelize.ChartSeries{{Name: "", Categories: fmt.Sprintf("'Динамика'!$A$15:$A$%d", 15+monthCount), Values: fmt.Sprintf("'Динамика'!$C$15:$C$%d", 15+monthCount)}},
	})
	err = file.AddChart("Динамика", "H30", &excelize.Chart{
		Type:   excelize.Line,
		Title:  []excelize.RichTextRun{{Text: fmt.Sprintf("Репостов в среднем за %d месяцев", monthCount)}},
		Series: []excelize.ChartSeries{{Name: "", Categories: fmt.Sprintf("'Динамика'!$A$15:$A$%d", 15+monthCount), Values: fmt.Sprintf("'Динамика'!$D$15:$D$%d", 15+monthCount)}},
	})
	err = file.AddChart("Динамика", "H45", &excelize.Chart{
		Type:   excelize.Line,
		Title:  []excelize.RichTextRun{{Text: fmt.Sprintf("Комментариев в среднем за %d месяцев", monthCount)}},
		Series: []excelize.ChartSeries{{Name: "", Categories: fmt.Sprintf("'Динамика'!$A$15:$A$%d", 15+monthCount), Values: fmt.Sprintf("'Динамика'!$E$15:$E$%d", 15+monthCount)}},
	})
	err = file.AddChart("Динамика", "H60", &excelize.Chart{
		Type:   excelize.Line,
		Title:  []excelize.RichTextRun{{Text: fmt.Sprintf("Просмотров в среднем за %d месяцев", monthCount)}},
		Series: []excelize.ChartSeries{{Name: "", Categories: fmt.Sprintf("'Динамика'!$A$15:$A$%d", 15+monthCount), Values: fmt.Sprintf("'Динамика'!$F$15:$F$%d", 15+monthCount)}},
	})

	file.SaveAs(reportCfg.FileName)
	if err := file.DeleteSheet("Sheet1"); err != nil {
		log.Println("Sheet 1 not founded")
	}
	file.SaveAs(reportCfg.FileName)
	log.Println("Отчет успешно сохранен в файл " + reportCfg.FileName)
}
