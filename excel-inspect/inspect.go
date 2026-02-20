package excelinspect

import (
	"fmt"
	"os"

	"github.com/thedatashed/xlsxreader"
)

type Inspector struct {
	filePath string
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

	xl, err := xlsxreader.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open excel file: %w", err)
	}

	ins := &Inspector{
		filePath: filePath,
		xl:       xl,
	}

	for _, opt := range opts {
		opt(ins)
	}

	return ins, nil
}

func (i *Inspector) Close() error {
	if i.xl != nil {
		return i.xl.Close()
	}
	return nil
}

func (i *Inspector) Inspect() (*FileInfo, error) {
	sheetInfos := getSheetInfos(i.xl)

	info := &FileInfo{
		Sheets: make([]SheetInfo, 0, len(sheetInfos)),
	}

	for _, si := range sheetInfos {
		info.Sheets = append(info.Sheets, SheetInfo{
			Name:        si.Name,
			RowCount:    si.RowCount,
			ColumnCount: si.ColumnCount,
		})
	}

	return info, nil
}

func (i *Inspector) InspectWithDetails() (*FileInfo, error) {
	sheetInfos := getSheetInfos(i.xl)

	info := &FileInfo{
		Sheets:       make([]SheetInfo, 0, len(sheetInfos)),
		SheetDetails: make([]SheetDetail, 0, len(sheetInfos)),
	}

	for _, si := range sheetInfos {
		info.Sheets = append(info.Sheets, SheetInfo{
			Name:        si.Name,
			RowCount:    si.RowCount,
			ColumnCount: si.ColumnCount,
		})

		detail := SheetDetail{
			Name:        si.Name,
			RowCount:    si.RowCount,
			ColumnCount: si.ColumnCount,
			Headers:     si.Headers,
			Columns:     si.Columns,
		}
		info.SheetDetails = append(info.SheetDetails, detail)
	}

	return info, nil
}

type sheetInfo struct {
	Name        string
	RowCount    int
	ColumnCount int
	Headers     []string
	Columns     []ColumnInfo
}

func getSheetInfos(xl *xlsxreader.XlsxFileCloser) []sheetInfo {
	var result []sheetInfo

	for _, sheetName := range xl.Sheets {
		si := sheetInfo{
			Name: sheetName,
		}

		rows := xl.ReadRows(sheetName)
		rowCount := 0
		maxCols := 0

		for row := range rows {
			rowCount++

			if rowCount == 1 {
				maxCols = len(row.Cells)
				si.ColumnCount = maxCols
				for _, cell := range row.Cells {
					si.Headers = append(si.Headers, cell.Value)
				}
			} else if rowCount == 2 {
				for colIdx, cell := range row.Cells {
					if colIdx >= maxCols {
						break
					}
					colInfo := ColumnInfo{
						Name:          si.Headers[colIdx],
						StartPosition: fmt.Sprintf("%s2", columnLetter(colIdx)),
					}
					if cell.Value != "" {
						colInfo.SampleValues = append(colInfo.SampleValues, cell.Value)
						colInfo.DataType = getDataType(cell.Value)
					}
					si.Columns = append(si.Columns, colInfo)
				}
			} else if rowCount <= 6 {
				for colIdx, cell := range row.Cells {
					if colIdx >= maxCols || colIdx >= len(si.Columns) {
						break
					}
					if cell.Value != "" {
						si.Columns[colIdx].SampleValues = append(
							si.Columns[colIdx].SampleValues, cell.Value)
					}
				}
			}

			if rowCount > 1000 {
				break
			}
		}

		si.RowCount = rowCount

		for colIdx := range si.Headers {
			if colIdx >= len(si.Columns) {
				si.Columns = append(si.Columns, ColumnInfo{
					Name:          si.Headers[colIdx],
					StartPosition: fmt.Sprintf("%s2", columnLetter(colIdx)),
				})
			}
		}

		result = append(result, si)
	}

	return result
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
