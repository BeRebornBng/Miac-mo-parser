package domain

type Cells struct {
	Title     string
	CellsName []string
}

type ExcelConfig struct {
	FileName  string
	SheetName string
	Cells     []Cells
}
