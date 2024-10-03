package vkClient

import (
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"net/url"
	"time"

	"github.com/Miac-mo-parser/domain"
	"github.com/Miac-mo-parser/utils/dates"
)

func GetVkPosts(domains []string, start time.Time, end time.Time) ([]domain.Response, error) {
	count := 100
	offset := 0
	vks := make([]domain.Response, 0)
	for _, dom := range domains {
	nextPosts:
		fmt.Println(dom)
		data := url.Values{
			"access_token": {"9beebd669beebd669beebd663f98f01ac399bee9beebd66fd03b1233b2b354972bcd403"},
			"offset":       {fmt.Sprintf("%d", offset)},
			"domain":       {dom},
			"count":        {fmt.Sprintf("%d", count)},
			"v":            {"5.199"},
		}
		resp, err := http.PostForm("https://api.vk.com/method/wall.get", data)
		if err != nil {
			log.Fatal(err)
		}
		defer resp.Body.Close()
		body, err := io.ReadAll(resp.Body)
		if err != nil {
			log.Fatal(err)
		}
		var vkResponse domain.Response
		err = json.Unmarshal(body, &vkResponse)
		if err != nil {
			log.Fatalf("Ошибка при парсинге JSON: %v", err)
		}

		for _, item := range vkResponse.Response.Items {
			if item.Date <= int(start.Unix()) {
				break
			}
			if item.Date >= int(start.Unix()) && item.Date <= int(end.Unix()) {
				vks = append(vks, vkResponse)
			}
		}
		if vkResponse.Response.Items[len(vkResponse.Response.Items)-1].Date >= int(start.Unix()) {
			offset += count
			if offset >= vkResponse.Response.Count {
				break
			}
			goto nextPosts
		}
	}
	return vks, nil
}

func GetVkPost(dom string, start time.Time, end time.Time) (domain.Response, error) {
	count := 100
	offset := 0
	vkData := domain.ResponseData{}
	vkItems := make([]domain.Item, 0)
nextPosts:
	data := url.Values{
		"access_token": {"9beebd669beebd669beebd663f98f01ac399bee9beebd66fd03b1233b2b354972bcd403"},
		"offset":       {fmt.Sprintf("%d", offset)},
		"domain":       {dom},
		"count":        {fmt.Sprintf("%d", count)},
		"v":            {"5.199"},
	}
	resp, err := http.PostForm("https://api.vk.com/method/wall.get", data)
	if err != nil {
		log.Fatal(err)
	}
	defer resp.Body.Close()
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		log.Fatal(err)
	}
	var vkResponse domain.Response
	err = json.Unmarshal(body, &vkResponse)
	if err != nil {
		log.Fatalf("Ошибка при парсинге JSON: %v", err)
	}
	for _, item := range vkResponse.Response.Items {
		fmt.Println(item.Text)
		if item.Date <= int(start.Unix()) {
			continue
		}
		if item.Date >= int(start.Unix()) && item.Date <= int(end.Unix()) {
			vkItems = append(vkItems, item)
		}
	}

	// Проверка на то, что срез Items не пуст
	if len(vkResponse.Response.Items) > 0 {
		// Продолжаем проверку даты последнего элемента
		if vkResponse.Response.Items[len(vkResponse.Response.Items)-1].Date >= int(start.Unix()) {
			offset += count
			if offset >= vkResponse.Response.Count {
				vkData.Count = vkResponse.Response.Count
				vkData.Items = vkItems
				return domain.Response{Response: vkData}, nil
			}
			goto nextPosts
		}
	}

	vkData.Count = vkResponse.Response.Count
	vkData.Items = vkItems
	return domain.Response{Response: vkData}, nil
}

func VkCountInMonth(items []domain.Item, dates [5]dates.MonthBorders) domain.MonthPublishes {
	vkc := domain.MonthPublishes{}
	i := 0
	vkc.PostsCount = make([]float64, 8)
	vkc.LikesCount = make([]float64, 8)
	vkc.CommentsCount = make([]float64, 8)
	vkc.RepostsCount = make([]float64, 8)
	vkc.ViewsCount = make([]float64, 8)
	for _, item := range items {
		iDate := time.Unix(int64(item.Date), 0)
		fmt.Println(iDate)
		if (iDate.Compare(dates[i].Start) == 0 || iDate.Compare(dates[i].Start) == 1) && (iDate.Compare(dates[i].End) == -1 || iDate.Compare(dates[i].End) == 0) {
			fmt.Println(iDate)
			vkc.PostsCount[i] += 1
			vkc.LikesCount[i] += float64(item.Likes.Count)
			vkc.CommentsCount[i] += float64(item.Comments.Count)
			vkc.RepostsCount[i] += float64(item.Reposts.Count)
			vkc.ViewsCount[i] += float64(item.Views.Count)
			vkc.LikesCount[5] += float64(item.Likes.Count)
			vkc.CommentsCount[5] += float64(item.Comments.Count)
			vkc.RepostsCount[5] += float64(item.Reposts.Count)
			vkc.ViewsCount[5] += float64(item.Views.Count)
			vkc.PostsCount[7] += 1
			vkc.LikesCount[7] += float64(item.Likes.Count)
			vkc.CommentsCount[7] += float64(item.Comments.Count)
			vkc.RepostsCount[7] += float64(item.Reposts.Count)
			vkc.ViewsCount[7] += float64(item.Views.Count)
		} else {
			i++
			if i == 5 {
				return vkc
			}
			fmt.Println(iDate)
			vkc.PostsCount[i] += 1
			vkc.LikesCount[i] += float64(item.Likes.Count)
			vkc.CommentsCount[i] += float64(item.Comments.Count)
			vkc.RepostsCount[i] += float64(item.Reposts.Count)
			vkc.ViewsCount[i] += float64(item.Views.Count)
			vkc.PostsCount[7] += 1
			vkc.LikesCount[7] += float64(item.Likes.Count)
			vkc.CommentsCount[7] += float64(item.Comments.Count)
			vkc.RepostsCount[7] += float64(item.Reposts.Count)
			vkc.ViewsCount[7] += float64(item.Views.Count)
		}
	}
	for i := 0; i < 4; i++ {
		vkc.PostsCount[5] += vkc.PostsCount[i]
	}
	vkc.PostsCount[6] = vkc.PostsCount[5] / 4.0
	vkc.LikesCount[6] = vkc.LikesCount[5] / 4.0
	vkc.CommentsCount[6] = vkc.CommentsCount[5] / 4.0
	vkc.RepostsCount[6] = vkc.RepostsCount[5] / 4.0
	vkc.ViewsCount[6] = vkc.ViewsCount[5] / 4.0
	return vkc
}
