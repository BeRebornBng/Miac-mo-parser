package domain

type SheetCells struct {
	Sheet  string
	Titles []string
}

type ExcelConfig struct {
	FileName   string
	SheetCells []SheetCells
}

func NewExcelConfig(filename string, sheetCells []SheetCells) *ExcelConfig {
	return &ExcelConfig{
		FileName:   filename,
		SheetCells: sheetCells,
	}
}
