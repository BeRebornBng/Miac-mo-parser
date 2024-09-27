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

func SplitMonth() [splitWeeks]Dates {
	start := StartNowMoth()
	dates := [splitWeeks]Dates{}
	for i := 0; i < splitWeeks; i++ {
		if i == splitWeeks-1 {
			dates[i].Start = start
			dates[i].End = EndNowMonth()
			break
		}
		dates[i].Start = start
		start = start.AddDate(0, 0, 6)
		dates[i].End = start
		start = start.AddDate(0, 0, 1)
	}
	return dates
}
