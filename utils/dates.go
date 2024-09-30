package utils

import (
	"time"
)

type Dates struct {
	Start time.Time
	End   time.Time
}

func StartNowMoth() time.Time {
	now := time.Now()
	startOfMonth := time.Date(now.Year(), now.Month(), 1, 0, 0, 0, 0, now.Location())
	return startOfMonth
}

func EndNowMonth() time.Time {
	now := time.Now()
	endOfMonth := time.Date(now.Year(), now.Month()+1, 1, 0, 0, 0, 0, now.Location()).Add(-time.Second)
	return endOfMonth
}

const (
	splitWeeks = 5
)

func SplitMonth(start time.Time, end time.Time) [splitWeeks]Dates {
	dates := [splitWeeks]Dates{}
	for i := 0; i < splitWeeks; i++ {
		if i == splitWeeks-1 {
			dates[i].Start = start
			dates[i].End = end
			break
		}
		dates[i].Start = start
		start = start.AddDate(0, 0, 6)
		dates[i].End = start
		start = start.AddDate(0, 0, 1)
	}
	return dates
}

func MonthToRussian(month time.Month) string {
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
