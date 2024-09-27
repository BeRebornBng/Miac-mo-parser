package domain

// Структура для получения основного ответа
type Response struct {
	Response ResponseData `json:"response"`
}

// Основная структура для поля "response"
type ResponseData struct {
	Count        int           `json:"count"`
	Items        []Item        `json:"items"`
	ReactionSets []ReactionSet `json:"reaction_sets"`
}

// Структура для элемента "items"
type Item struct {
	InnerType      string       `json:"inner_type"`
	Donut          Donut        `json:"donut"`
	Comments       Comments     `json:"comments"`
	MarkedAsAds    int          `json:"marked_as_ads"`
	Hash           string       `json:"hash"`
	Type           string       `json:"type"`
	CarouselOffset int          `json:"carousel_offset"`
	Attachments    []Attachment `json:"attachments"`
	Date           int          `json:"date"`
	FromID         int          `json:"from_id"`
	ID             int          `json:"id"`
	Likes          Likes        `json:"likes"`
	ReactionSetID  string       `json:"reaction_set_id"`
	Reactions      Reactions    `json:"reactions"`
	OwnerID        int          `json:"owner_id"`
	PostType       string       `json:"post_type"`
	Reposts        Reposts      `json:"reposts"`
	Text           string       `json:"text"`
	Views          Views        `json:"views"`
}

// Структура для "donut"
type Donut struct {
	IsDonut bool `json:"is_donut"`
}

// Структура для "comments"
type Comments struct {
	Count int `json:"count"`
}

// Структура для "attachments"
type Attachment struct {
	Type  string `json:"type"`
	Photo Photo  `json:"photo"`
}

// Структура для "photo"
type Photo struct {
	AlbumID      int       `json:"album_id"`
	Date         int       `json:"date"`
	ID           int       `json:"id"`
	OwnerID      int       `json:"owner_id"`
	AccessKey    string    `json:"access_key"`
	Sizes        []Size    `json:"sizes"`
	Text         string    `json:"text"`
	WebViewToken string    `json:"web_view_token"`
	HasTags      bool      `json:"has_tags"`
	OrigPhoto    OrigPhoto `json:"orig_photo"`
}

// Структура для "sizes"
type Size struct {
	Height int    `json:"height"`
	Type   string `json:"type"`
	Width  int    `json:"width"`
	URL    string `json:"url"`
}

// Структура для "orig_photo"
type OrigPhoto struct {
	Height int    `json:"height"`
	Type   string `json:"type"`
	Width  int    `json:"width"`
	URL    string `json:"url"`
}

// Структура для "likes"
type Likes struct {
	CanLike   int `json:"can_like"`
	Count     int `json:"count"`
	UserLikes int `json:"user_likes"`
}

// Структура для "reactions"
type Reactions struct {
	Count int            `json:"count"`
	Items []ReactionItem `json:"items"`
}

// Структура для каждого элемента в "reactions"
type ReactionItem struct {
	ID    int `json:"id"`
	Count int `json:"count"`
}

// Структура для "reposts"
type Reposts struct {
	Count int `json:"count"`
}

// Структура для "views"
type Views struct {
	Count int `json:"count"`
}

// Структура для "reaction_sets"
type ReactionSet struct {
	ID    string          `json:"id"`
	Items []ReactionTitle `json:"items"`
}

// Структура для каждого элемента в "reaction_sets"
type ReactionTitle struct {
	ID    int    `json:"id"`
	Title string `json:"title"`
	Asset Asset  `json:"asset"`
}

// Структура для "asset"
type Asset struct {
	AnimationURL string     `json:"animation_url"`
	Images       []Image    `json:"images"`
	Title        Title      `json:"title"`
	TitleColor   TitleColor `json:"title_color"`
}

// Структура для "image"
type Image struct {
	URL    string `json:"url"`
	Width  int    `json:"width"`
	Height int    `json:"height"`
}

// Структура для "title"
type Title struct {
	Color Color `json:"color"`
}

// Структура для "color"
type Color struct {
	Foreground Foreground `json:"foreground"`
	Background Background `json:"background"`
}

// Структура для "foreground"
type Foreground struct {
	Light string `json:"light"`
	Dark  string `json:"dark"`
}

// Структура для "background"
type Background struct {
	Light string `json:"light"`
	Dark  string `json:"dark"`
}

// Структура для "title_color"
type TitleColor struct {
	Light string `json:"light"`
	Dark  string `json:"dark"`
}
