package domain

type Cells struct {
	Title    string
	CellName string
}

type ExcelConfig struct {
	FileName string
	Sheet    string
	Cells    []Cells
}

func NewExcelConfig(filename string, sheet string, cells []Cells) *ExcelConfig {
	return &ExcelConfig{
		FileName: filename,
		Sheet:    sheet,
		Cells:    cells,
	}
}
