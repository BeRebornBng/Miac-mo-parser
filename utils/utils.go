package utils

import (
	"errors"
	"strings"
)

// получеиние доменов из ссылок
func DomainsFromLinks(links []string) ([]string, error) {

	domains := make([]string, len(links))
	for i, link := range links {
		str, err := strings.CutPrefix(link, "https://vk.com/")
		if !err {
			return nil, errors.New("некорректная ссылка")
		}
		domains[i] = str
	}

	return domains, nil

}
