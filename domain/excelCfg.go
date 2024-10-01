package domain

type SheetCells struct {
	Sheet string
	Title []string
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
