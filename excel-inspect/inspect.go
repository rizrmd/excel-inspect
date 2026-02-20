package excelinspect

import (
	"fmt"
	"os"
	"strings"

	"github.com/thedatashed/xlsxreader"
	"github.com/xuri/excelize/v2"
)

type Inspector struct {
	filePath string
	file     *excelize.File
	xl       *xlsxreader.XlsxFileCloser
}

type InspectorOption func(*Inspector)

func WithTimeout(seconds int) InspectorOption {
	return func(i *Inspector) {}
}

func WithHeaderRow(row int) InspectorOption {
	return func(i *Inspector) {}
}

func WithMaxSampleRows(rows int) InspectorOption {
	return func(i *Inspector) {}
}

func WithIncludeRowCount(include bool) InspectorOption {
	return func(i *Inspector) {}
}

type SheetInfo struct {
	Name        string `json:"name"`
	RowCount    int    `json:"row_count"`
	ColumnCount int    `json:"column_count"`
}

type ColumnInfo struct {
	Name          string        `json:"name"`
	StartPosition string        `json:"start_position"`
	SampleValues  []interface{} `json:"sample_values"`
	DataType      string        `json:"data_type"`
}

type SheetDetail struct {
	Name        string       `json:"name"`
	RowCount    int          `json:"row_count"`
	ColumnCount int          `json:"column_count"`
	Headers     []string     `json:"headers"`
	Columns     []ColumnInfo `json:"columns"`
}

type FileInfo struct {
	Sheets       []SheetInfo   `json:"sheets"`
	SheetDetails []SheetDetail `json:"sheet_details,omitempty"`
}

func New(filePath string, opts ...InspectorOption) (*Inspector, error) {
	if _, err := os.Stat(filePath); err != nil {
		return nil, fmt.Errorf("file not found: %w", err)
	}

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open excel file: %w", err)
	}

	xl, err := xlsxreader.OpenFile(filePath)
	if err != nil {
		f.Close()
		return nil, fmt.Errorf("failed to open excel file: %w", err)
	}

	ins := &Inspector{
		filePath: filePath,
		file:     f,
		xl:       xl,
	}

	for _, opt := range opts {
		opt(ins)
	}

	return ins, nil
}

func (i *Inspector) Close() error {
	if i.file != nil {
		i.file.Close()
	}
	if i.xl != nil {
		i.xl.Close()
	}
	return nil
}

func (i *Inspector) Inspect() (*FileInfo, error) {
	sheets := i.file.GetSheetList()

	info := &FileInfo{
		Sheets: make([]SheetInfo, 0, len(sheets)),
	}

	for _, sheetName := range sheets {
		visible, err := i.file.GetSheetVisible(sheetName)
		if err != nil || !visible {
			continue
		}

		rowCount := i.getRowCount(sheetName)
		colCount := i.getColumnCount(sheetName)

		info.Sheets = append(info.Sheets, SheetInfo{
			Name:        sheetName,
			RowCount:    rowCount,
			ColumnCount: colCount,
		})
	}

	return info, nil
}

func (i *Inspector) InspectWithDetails() (*FileInfo, error) {
	sheets := i.file.GetSheetList()

	info := &FileInfo{
		Sheets:       make([]SheetInfo, 0, len(sheets)),
		SheetDetails: make([]SheetDetail, 0, len(sheets)),
	}

	for _, sheetName := range sheets {
		visible, err := i.file.GetSheetVisible(sheetName)
		if err != nil || !visible {
			continue
		}

		rowCount := i.getRowCount(sheetName)
		colCount := i.getColumnCount(sheetName)

		info.Sheets = append(info.Sheets, SheetInfo{
			Name:        sheetName,
			RowCount:    rowCount,
			ColumnCount: colCount,
		})

		detail := i.inspectSheetDetail(sheetName)
		info.SheetDetails = append(info.SheetDetails, detail)
	}

	return info, nil
}

func (i *Inspector) getRowCount(sheetName string) int {
	rows := i.xl.ReadRows(sheetName)
	rowCount := 0
	for range rows {
		rowCount++
		if rowCount > 1000 {
			break
		}
	}
	return rowCount
}

func (i *Inspector) getColumnCount(sheetName string) int {
	rows := i.xl.ReadRows(sheetName)
	for row := range rows {
		return len(row.Cells)
	}
	return 0
}

func (i *Inspector) inspectSheetDetail(sheetName string) SheetDetail {
	detail := SheetDetail{
		Name: sheetName,
	}

	rows := i.xl.ReadRows(sheetName)
	rowCount := 0
	headerRow := 1

	for row := range rows {
		rowCount++

		if rowCount == 1 {
			detail.ColumnCount = len(row.Cells)
			if detail.ColumnCount == 0 {
				continue
			}
			skipRow := false
			for idx, cell := range row.Cells {
				if idx < 3 && cell.Value != "" {
					if len(cell.Value) < 5 || strings.Contains(cell.Value, "-") {
						skipRow = true
						break
					}
				}
			}
			if skipRow && detail.ColumnCount > 3 {
				headerRow = 2
				continue
			}
			for _, cell := range row.Cells {
				detail.Headers = append(detail.Headers, cell.Value)
			}
		} else if rowCount == headerRow {
			if len(detail.Headers) == 0 {
				detail.ColumnCount = len(row.Cells)
				for _, cell := range row.Cells {
					detail.Headers = append(detail.Headers, cell.Value)
				}
			}
		}

		if len(detail.Headers) > 0 && rowCount == headerRow+1 {
			maxCols := len(detail.Headers)
			for colIdx, cell := range row.Cells {
				if colIdx >= maxCols {
					break
				}
				colInfo := ColumnInfo{
					Name:          detail.Headers[colIdx],
					StartPosition: fmt.Sprintf("%s%d", columnLetter(colIdx), headerRow),
				}
				if cell.Value != "" {
					colInfo.SampleValues = append(colInfo.SampleValues, cell.Value)
					colInfo.DataType = getDataType(cell.Value)
				}
				detail.Columns = append(detail.Columns, colInfo)
			}
		} else if len(detail.Columns) > 0 && rowCount <= headerRow+5 {
			maxCols := len(detail.Headers)
			for colIdx, cell := range row.Cells {
				if colIdx >= maxCols || colIdx >= len(detail.Columns) {
					break
				}
				if cell.Value != "" {
					detail.Columns[colIdx].SampleValues = append(
						detail.Columns[colIdx].SampleValues, cell.Value)
				}
			}
		}

		if rowCount > 1000 {
			break
		}
	}

	detail.RowCount = rowCount

	for colIdx := range detail.Headers {
		if colIdx >= len(detail.Columns) {
			detail.Columns = append(detail.Columns, ColumnInfo{
				Name:          detail.Headers[colIdx],
				StartPosition: fmt.Sprintf("%s%d", columnLetter(colIdx), headerRow),
			})
		}
	}

	return detail
}

func getDataType(value string) string {
	if value == "" {
		return "empty"
	}

	for _, r := range value {
		if r >= '0' && r <= '9' || r == '.' || r == '-' || r == '+' || r == 'e' || r == 'E' {
			continue
		}
		return "string"
	}
	return "number"
}

func columnLetter(colIdx int) string {
	result := ""
	for {
		result = string(rune('A'+colIdx%26)) + result
		colIdx = colIdx/26 - 1
		if colIdx < 0 {
			break
		}
	}
	return result
}
