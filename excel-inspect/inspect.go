package excelinspect

import (
	"context"
	"fmt"
	"os"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

type Inspector struct {
	filePath string
	file     *excelize.File
	mu       sync.Mutex
	Timeout  int
}

type InspectorOption func(*Inspector)

func WithTimeout(seconds int) InspectorOption {
	return func(i *Inspector) {
		i.Timeout = seconds
	}
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

	ins := &Inspector{
		filePath: filePath,
		file:     f,
	}

	for _, opt := range opts {
		opt(ins)
	}

	return ins, nil
}

func (i *Inspector) Close() error {
	i.mu.Lock()
	defer i.mu.Unlock()
	if i.file != nil {
		return i.file.Close()
	}
	return nil
}

func (i *Inspector) Inspect() (*FileInfo, error) {
	i.mu.Lock()
	defer i.mu.Unlock()

	sheets := i.file.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("no sheets found in file")
	}

	info := &FileInfo{
		Sheets: make([]SheetInfo, 0, len(sheets)),
	}

	ctx, cancel := context.WithTimeout(context.Background(), time.Duration(i.Timeout)*time.Second)
	defer cancel()

	for _, sheetName := range sheets {
		visible, err := i.file.GetSheetVisible(sheetName)
		if err != nil || !visible {
			continue
		}

		select {
		case <-ctx.Done():
			return info, nil
		default:
		}

		sheetInfo := i.inspectSheetWithContext(ctx, sheetName)
		info.Sheets = append(info.Sheets, sheetInfo)
	}

	return info, nil
}

func (i *Inspector) InspectWithDetails() (*FileInfo, error) {
	i.mu.Lock()
	defer i.mu.Unlock()

	sheets := i.file.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("no sheets found in file")
	}

	info := &FileInfo{
		Sheets:       make([]SheetInfo, 0, len(sheets)),
		SheetDetails: make([]SheetDetail, 0, len(sheets)),
	}

	ctx, cancel := context.WithTimeout(context.Background(), time.Duration(i.Timeout)*time.Second)
	defer cancel()

	for _, sheetName := range sheets {
		visible, err := i.file.GetSheetVisible(sheetName)
		if err != nil || !visible {
			continue
		}

		sheetInfo := i.inspectSheetWithContext(ctx, sheetName)
		info.Sheets = append(info.Sheets, sheetInfo)

		select {
		case <-ctx.Done():
			return info, nil
		default:
		}

		detail := i.inspectSheetDetailWithContext(ctx, sheetName)
		info.SheetDetails = append(info.SheetDetails, detail)
	}

	return info, nil
}

func (i *Inspector) inspectSheetWithContext(ctx context.Context, sheetName string) SheetInfo {
	resultCh := make(chan SheetInfo, 1)

	go func() {
		resultCh <- i.inspectSheet(sheetName)
	}()

	select {
	case <-ctx.Done():
		return SheetInfo{Name: sheetName, RowCount: 0, ColumnCount: 0}
	case result := <-resultCh:
		return result
	}
}

func (i *Inspector) inspectSheet(sheetName string) SheetInfo {
	colCount := 0

	for colIdx := 0; colIdx < 100; colIdx++ {
		colLetter := columnLetter(colIdx)
		headerCell := fmt.Sprintf("%s1", colLetter)
		val, err := i.file.GetCellValue(sheetName, headerCell)
		if err != nil || val == "" {
			colCount = colIdx
			break
		}
	}

	return SheetInfo{
		Name:        sheetName,
		RowCount:    0,
		ColumnCount: colCount,
	}
}

func (i *Inspector) inspectSheetDetailWithContext(ctx context.Context, sheetName string) SheetDetail {
	resultCh := make(chan SheetDetail, 1)

	go func() {
		resultCh <- i.inspectSheetDetail(sheetName)
	}()

	select {
	case <-ctx.Done():
		return SheetDetail{Name: sheetName}
	case result := <-resultCh:
		return result
	}
}

func (i *Inspector) inspectSheetDetail(sheetName string) SheetDetail {
	detail := SheetDetail{
		Name: sheetName,
	}

	detail.Headers = make([]string, 0, 100)
	detail.Columns = make([]ColumnInfo, 0, 100)

	for colIdx := 0; colIdx < 500; colIdx++ {
		colLetter := columnLetter(colIdx)

		colInfo := ColumnInfo{
			StartPosition: fmt.Sprintf("%s1", colLetter),
		}

		headerCell := fmt.Sprintf("%s1", colLetter)
		headerVal, err := i.file.GetCellValue(sheetName, headerCell)
		if err != nil || headerVal == "" {
			detail.ColumnCount = colIdx
			break
		}

		colInfo.Name = headerVal
		detail.Headers = append(detail.Headers, colInfo.Name)

		colInfo.SampleValues = make([]interface{}, 0, 5)
		for sampleRow := 2; sampleRow <= 6; sampleRow++ {
			cell := fmt.Sprintf("%s%d", colLetter, sampleRow)
			val, _ := i.file.GetCellValue(sheetName, cell)
			if val != "" {
				colInfo.SampleValues = append(colInfo.SampleValues, val)
				if colInfo.DataType == "" {
					colInfo.DataType = getDataType(val)
				}
			}
		}

		detail.Columns = append(detail.Columns, colInfo)
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

func columnIndex(letter string) int {
	result := 0
	for _, r := range letter {
		result = result*26 + int(r-'A'+1)
	}
	return result - 1
}
