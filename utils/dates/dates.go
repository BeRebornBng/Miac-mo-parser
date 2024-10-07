package dates

import (
	"errors"
	"time"
)

type MonthBorders struct {
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
	monthParts = 5
)

func SplitMonth(start time.Time, end time.Time) [monthParts]MonthBorders {
	dates := [monthParts]MonthBorders{}
	for i := 0; i < monthParts; i++ {
		if i == monthParts-1 {
			dates[i].Start = start
			dates[i].End = end
			break
		}
		dates[i].Start = start
		start = start.AddDate(0, 0, 6)
		dates[i].End = start.AddDate(0, 0, 1).Add(-time.Second)
		start = start.AddDate(0, 0, 1)
	}
	return dates
}

func MonthToRussian(month time.Month) (string, error) {
	switch month {
	case time.January:
		return "Январь", nil
	case time.February:
		return "Февраль", nil
	case time.March:
		return "Март", nil
	case time.April:
		return "Апрель", nil
	case time.May:
		return "Май", nil
	case time.June:
		return "Июнь", nil
	case time.July:
		return "Июль", nil
	case time.August:
		return "Август", nil
	case time.September:
		return "Сентябрь", nil
	case time.October:
		return "Октябрь", nil
	case time.November:
		return "Ноябрь", nil
	case time.December:
		return "Декабрь", nil
	}
	return "", errors.New("месяц не найден")
}
