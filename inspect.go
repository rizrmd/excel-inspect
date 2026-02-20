package excelinspect

import (
	"fmt"
	"os"
	"strings"

	toon "github.com/mateuszkardas/toon-go"
	"github.com/thedatashed/xlsxreader"
	"github.com/xuri/excelize/v2"
)

type Inspector struct {
	filePath         string
	file             *excelize.File
	xl               *xlsxreader.XlsxFileCloser
	progressCallback func(ProgressInfo)
	progressChan     chan<- ProgressInfo
}

type InspectorOption func(*Inspector)

type ProgressInfo struct {
	Phase   string  `json:"phase"`
	Sheet   string  `json:"sheet,omitempty"`
	Current int     `json:"current"`
	Total   int     `json:"total"`
	Percent float64 `json:"percent"`
}

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

func WithProgressCallback(fn func(ProgressInfo)) InspectorOption {
	return func(i *Inspector) {
		i.progressCallback = fn
	}
}

func WithProgressChannel(ch chan<- ProgressInfo) InspectorOption {
	return func(i *Inspector) {
		i.progressChan = ch
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
	Sections    []Section    `json:"sections,omitempty"`
}

type FileInfo struct {
	Sheets       []SheetInfo   `json:"sheets"`
	SheetDetails []SheetDetail `json:"sheet_details,omitempty"`
}

type Section struct {
	Title      string       `json:"title"`
	HeaderRow  int          `json:"header_row"`
	StartRow   int          `json:"start_row"`
	EndRow     int          `json:"end_row"`
	Headers    []string     `json:"headers"`
	Columns    []ColumnInfo `json:"columns"`
	RowCount   int          `json:"row_count"`
	ColumnCount int         `json:"column_count"`
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
	visibleSheets := i.visibleSheets(sheets)

	info := &FileInfo{
		Sheets: make([]SheetInfo, 0, len(sheets)),
	}

	total := len(visibleSheets)
	i.emitProgress("inspect_sheets", "", 0, total)
	for idx, sheetName := range visibleSheets {
		rowCount := i.getRowCount(sheetName)
		colCount := i.getColumnCount(sheetName)

		info.Sheets = append(info.Sheets, SheetInfo{
			Name:        sheetName,
			RowCount:    rowCount,
			ColumnCount: colCount,
		})
		i.emitProgress("inspect_sheets", sheetName, idx+1, total)
	}

	return info, nil
}

func (i *Inspector) InspectWithDetails() (*FileInfo, error) {
	sheets := i.file.GetSheetList()
	visibleSheets := i.visibleSheets(sheets)

	info := &FileInfo{
		Sheets:       make([]SheetInfo, 0, len(sheets)),
		SheetDetails: make([]SheetDetail, 0, len(sheets)),
	}

	total := len(visibleSheets)
	i.emitProgress("inspect_details", "", 0, total)
	for idx, sheetName := range visibleSheets {
		rowCount := i.getRowCount(sheetName)
		colCount := i.getColumnCount(sheetName)

		info.Sheets = append(info.Sheets, SheetInfo{
			Name:        sheetName,
			RowCount:    rowCount,
			ColumnCount: colCount,
		})

		detail := i.inspectSheetDetail(sheetName)
		info.SheetDetails = append(info.SheetDetails, detail)
		i.emitProgress("inspect_details", sheetName, idx+1, total)
	}

	return info, nil
}

func (i *Inspector) InspectTOON() (string, error) {
	info, err := i.Inspect()
	if err != nil {
		return "", err
	}
	return toon.Marshal(info, nil)
}

func (i *Inspector) InspectWithDetailsTOON() (string, error) {
	info, err := i.InspectWithDetails()
	if err != nil {
		return "", err
	}
	return toon.Marshal(i.buildCompactTOONPayloadFull(info), nil)
}

func (i *Inspector) InspectWithDetailsTOONSample() (string, error) {
	info, err := i.InspectWithDetails()
	if err != nil {
		return "", err
	}
	return toon.Marshal(buildCompactTOONPayloadSample(info), nil)
}

func (i *Inspector) InspectMarkdown() (string, error) {
	info, err := i.Inspect()
	if err != nil {
		return "", err
	}
	return i.MarkdownFromInfo(info, false), nil
}

func (i *Inspector) InspectWithDetailsMarkdown() (string, error) {
	info, err := i.InspectWithDetails()
	if err != nil {
		return "", err
	}
	return i.MarkdownFromInfo(info, true), nil
}

func (i *Inspector) MarkdownFromInfo(info *FileInfo, detailed bool) string {
	return i.buildMarkdown(info, detailed)
}

func (i *Inspector) buildMarkdown(info *FileInfo, detailed bool) string {
	var b strings.Builder

	b.WriteString("# Excel Inspect Report\n\n")
	b.WriteString("## Sheets\n\n")
	b.WriteString("| Name | Rows | Columns |\n")
	b.WriteString("| --- | ---: | ---: |\n")
	for _, s := range info.Sheets {
		b.WriteString(fmt.Sprintf("| %s | %d | %d |\n", escapeMarkdownCell(s.Name), s.RowCount, s.ColumnCount))
	}

	if !detailed || len(info.SheetDetails) == 0 {
		return b.String()
	}

	totalSections := 0
	for _, d := range info.SheetDetails {
		totalSections += len(d.Sections)
	}
	doneSections := 0
	if totalSections > 0 {
		i.emitProgress("markdown_sections", "", 0, totalSections)
	}

	b.WriteString("\n## Sheet Details\n")
	for _, d := range info.SheetDetails {
		b.WriteString(fmt.Sprintf("\n### %s\n\n", escapeMarkdownCell(d.Name)))
		b.WriteString(fmt.Sprintf("- Rows: %d\n", d.RowCount))
		b.WriteString(fmt.Sprintf("- Columns: %d\n", d.ColumnCount))
		b.WriteString(fmt.Sprintf("- Headers: %d\n", len(d.Headers)))

		if len(d.Columns) > 0 {
			b.WriteString("\n#### Columns\n\n")
			b.WriteString("| # | Name | Start | Type | Samples |\n")
			b.WriteString("| ---: | --- | --- | --- | --- |\n")
			for idx, c := range d.Columns {
				samples := toSampleStrings(c.SampleValues)
				b.WriteString(fmt.Sprintf(
					"| %d | %s | %s | %s | %s |\n",
					idx+1,
					escapeMarkdownCell(c.Name),
					escapeMarkdownCell(c.StartPosition),
					escapeMarkdownCell(c.DataType),
					escapeMarkdownCell(strings.Join(samples, ", ")),
				))
			}
		}

		if len(d.Sections) > 0 {
			maxEndRow := 0
			for _, s := range d.Sections {
				if s.EndRow > maxEndRow {
					maxEndRow = s.EndRow
				}
			}
			rowsByNum := i.sheetRowsByNumber(d.Name, maxEndRow)

			b.WriteString("\n#### Sections\n\n")
			for idx, s := range d.Sections {
				b.WriteString(fmt.Sprintf("##### Section %d: %s\n\n", idx+1, escapeMarkdownCell(s.Title)))
				b.WriteString(fmt.Sprintf("- Header row: %d\n", s.HeaderRow))
				b.WriteString(fmt.Sprintf("- Start row: %d\n", s.StartRow))
				b.WriteString(fmt.Sprintf("- End row: %d\n", s.EndRow))
				b.WriteString(fmt.Sprintf("- Rows: %d\n", s.RowCount))
				b.WriteString(fmt.Sprintf("- Columns: %d\n", s.ColumnCount))
				b.WriteString("\n")

				headers := s.Headers
				if len(headers) == 0 {
					headers = make([]string, s.ColumnCount)
					for hIdx := range headers {
						headers[hIdx] = fmt.Sprintf("Column %d", hIdx+1)
					}
				}

				values := sectionValuesFromRows(rowsByNum, s, len(headers))
				if len(values) == 0 {
					b.WriteString("_No section rows found._\n\n")
					doneSections++
					i.emitProgress("markdown_sections", d.Name, doneSections, totalSections)
					continue
				}

				b.WriteString("| ")
				for hIdx, h := range headers {
					if hIdx > 0 {
						b.WriteString(" | ")
					}
					b.WriteString(escapeMarkdownCell(h))
				}
				b.WriteString(" |\n")

				b.WriteString("| ")
				for hIdx := range headers {
					if hIdx > 0 {
						b.WriteString(" | ")
					}
					b.WriteString("---")
				}
				b.WriteString(" |\n")

				for _, row := range values {
					b.WriteString("| ")
					for cIdx, cell := range row {
						if cIdx > 0 {
							b.WriteString(" | ")
						}
						b.WriteString(escapeMarkdownCell(cell))
					}
					b.WriteString(" |\n")
				}
				b.WriteString("\n")
				doneSections++
				i.emitProgress("markdown_sections", d.Name, doneSections, totalSections)
			}
		}
	}

	return b.String()
}

func sectionValuesFromRows(rowsByNum map[int][]string, section Section, width int) [][]string {
	if width <= 0 || section.StartRow <= 0 || section.EndRow < section.StartRow {
		return nil
	}

	out := make([][]string, 0, max(0, section.EndRow-section.StartRow+1))
	for rowNum := section.StartRow; rowNum <= section.EndRow; rowNum++ {
		row, ok := rowsByNum[rowNum]
		if !ok {
			continue
		}
		values := make([]string, width)
		for idx := 0; idx < width; idx++ {
			if idx < len(row) {
				values[idx] = strings.TrimSpace(row[idx])
			}
		}
		if isEmptyRow(values) {
			continue
		}
		out = append(out, values)
	}
	return out
}

func (i *Inspector) sheetRowsByNumber(sheet string, maxRow int) map[int][]string {
	if maxRow <= 0 {
		return nil
	}
	rows := i.xl.ReadRows(sheet)
	out := make(map[int][]string, maxRow)
	rowNum := 0
	for row := range rows {
		rowNum++
		if rowNum > maxRow {
			break
		}
		values := make([]string, len(row.Cells))
		for idx, cell := range row.Cells {
			values[idx] = strings.TrimSpace(cell.Value)
		}
		out[rowNum] = values
		if rowNum%100 == 0 || rowNum == maxRow {
			i.emitProgress("markdown_scan_rows", sheet, rowNum, maxRow)
		}
	}
	return out
}

func escapeMarkdownCell(v string) string {
	v = strings.TrimSpace(v)
	if v == "" {
		return ""
	}
	v = strings.ReplaceAll(v, "\\", "\\\\")
	v = strings.ReplaceAll(v, "|", "\\|")
	v = strings.ReplaceAll(v, "\n", " ")
	return v
}

func buildCompactTOONPayloadSample(info *FileInfo) map[string]interface{} {
	payload := map[string]interface{}{
		"sheet_details": nil,
		"sections":      nil,
		"columns":       nil,
	}

	sheetMeta := make([]map[string]interface{}, 0, len(info.SheetDetails))
	type compactCol struct {
		sheet      string
		columnIdx  int
		name       string
		startPos   string
		dataType   string
		samples    []string
	}
	colIndex := make(map[string]int)
	compactCols := make([]compactCol, 0)
	sections := make([]map[string]interface{}, 0)

	for _, sd := range info.SheetDetails {
		sheetMeta = append(sheetMeta, map[string]interface{}{
			"name":          sd.Name,
			"row_count":     sd.RowCount,
			"column_count":  sd.ColumnCount,
			"header_count":  len(sd.Headers),
			"section_count": len(sd.Sections),
		})

		if len(sd.Sections) > 0 {
			for sIdx, sec := range sd.Sections {
				sections = append(sections, map[string]interface{}{
					"sheet":        sd.Name,
					"section_idx":  sIdx + 1,
					"title":        sec.Title,
					"header_row":   sec.HeaderRow,
					"start_row":    sec.StartRow,
					"end_row":      sec.EndRow,
					"row_count":    sec.RowCount,
					"column_count": sec.ColumnCount,
				})
				for cIdx, col := range sec.Columns {
					key := fmt.Sprintf("%s|%d|%s", sd.Name, cIdx+1, col.Name)
					if pos, ok := colIndex[key]; ok {
						compactCols[pos].samples = mergeSampleStrings(compactCols[pos].samples, toSampleStrings(col.SampleValues), 5)
						if compactCols[pos].dataType == "" && col.DataType != "" {
							compactCols[pos].dataType = col.DataType
						}
						continue
					}
					colIndex[key] = len(compactCols)
					compactCols = append(compactCols, compactCol{
						sheet:     sd.Name,
						columnIdx: cIdx + 1,
						name:      col.Name,
						startPos:  col.StartPosition,
						dataType:  col.DataType,
						samples:   toSampleStrings(col.SampleValues),
					})
				}
			}
			continue
		}

		for cIdx, col := range sd.Columns {
			key := fmt.Sprintf("%s|%d|%s", sd.Name, cIdx+1, col.Name)
			if pos, ok := colIndex[key]; ok {
				compactCols[pos].samples = mergeSampleStrings(compactCols[pos].samples, toSampleStrings(col.SampleValues), 5)
				if compactCols[pos].dataType == "" && col.DataType != "" {
					compactCols[pos].dataType = col.DataType
				}
				continue
			}
			colIndex[key] = len(compactCols)
			compactCols = append(compactCols, compactCol{
				sheet:     sd.Name,
				columnIdx: cIdx + 1,
				name:      col.Name,
				startPos:  col.StartPosition,
				dataType:  col.DataType,
				samples:   toSampleStrings(col.SampleValues),
			})
		}
	}

	columns := make([]map[string]interface{}, 0, len(compactCols))
	for _, c := range compactCols {
		columns = append(columns, map[string]interface{}{
			"sheet":          c.sheet,
			"column_idx":     c.columnIdx,
			"name":           c.name,
			"start_position": c.startPos,
			"data_type":      c.dataType,
			"samples":        strings.Join(c.samples, "|"),
		})
	}

	payload["sheet_details"] = sheetMeta
	payload["sections"] = sections
	payload["columns"] = columns
	return payload
}

func (i *Inspector) buildCompactTOONPayloadFull(info *FileInfo) map[string]interface{} {
	payload := buildCompactTOONPayloadSample(info)
	colsRaw, ok := payload["columns"].([]map[string]interface{})
	if !ok {
		return payload
	}

	cols := make([]map[string]interface{}, 0, len(colsRaw))
	for _, c := range colsRaw {
		cp := make(map[string]interface{}, len(c))
		for k, v := range c {
			cp[k] = v
		}
		cols = append(cols, cp)
	}

	type colRef struct {
		sheet string
		idx   int
		row   map[string]interface{}
	}
	refsBySheet := make(map[string][]colRef)
	for _, c := range cols {
		sheet, _ := c["sheet"].(string)
		columnIdx, _ := c["column_idx"].(int)
		if columnIdx <= 0 {
			continue
		}
		refsBySheet[sheet] = append(refsBySheet[sheet], colRef{
			sheet: sheet,
			idx:   columnIdx - 1,
			row:   c,
		})
	}

	totalSheets := len(refsBySheet)
	doneSheets := 0
	i.emitProgress("toon_full_values", "", 0, totalSheets)
	for sheet, refs := range refsBySheet {
		rows := i.xl.ReadRows(sheet)
		valuesByCol := make(map[int][]string)
		rowNum := 0
		for row := range rows {
			rowNum++
			if rowNum > 1000 {
				break
			}
			trimmed := make([]string, len(row.Cells))
			for idx, cell := range row.Cells {
				trimmed[idx] = strings.TrimSpace(cell.Value)
			}
			trimmed = trimTrailingEmpty(trimmed)
			if len(trimmed) == 0 || isLikelyHeaderRow(trimmed) || isSectionMarkerRow(trimmed) {
				continue
			}
			for _, ref := range refs {
				if ref.idx >= len(trimmed) {
					continue
				}
				v := strings.TrimSpace(trimmed[ref.idx])
				if v == "" {
					continue
				}
				valuesByCol[ref.idx] = append(valuesByCol[ref.idx], v)
			}
			if rowNum%100 == 0 || rowNum == 1000 {
				i.emitProgress("toon_full_values_rows", sheet, rowNum, 1000)
			}
		}

		for _, ref := range refs {
			ref.row["samples"] = strings.Join(valuesByCol[ref.idx], "|")
		}
		doneSheets++
		i.emitProgress("toon_full_values", sheet, doneSheets, totalSheets)
	}

	payload["columns"] = cols
	return payload
}

func toSampleStrings(values []interface{}) []string {
	if len(values) == 0 {
		return nil
	}
	out := make([]string, 0, len(values))
	for _, v := range values {
		s := strings.TrimSpace(fmt.Sprintf("%v", v))
		if s != "" {
			out = append(out, s)
		}
	}
	return out
}

func mergeSampleStrings(base []string, incoming []string, maxSamples int) []string {
	if base == nil {
		base = make([]string, 0, maxSamples)
	}
	for _, s := range incoming {
		if len(base) >= maxSamples {
			break
		}
		exists := false
		for _, b := range base {
			if b == s {
				exists = true
				break
			}
		}
		if !exists {
			base = append(base, s)
		}
	}
	return base
}

func isSectionMarkerRow(row []string) bool {
	upper := strings.ToUpper(strings.Join(row, " "))
	if strings.Contains(upper, "CROSS SELLING") || strings.Contains(upper, "NON CROSS SELLING") {
		return true
	}
	if strings.Contains(upper, "LAST UPDATE") || strings.Contains(upper, "HANDOVER") {
		return true
	}
	return false
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
	allRows := make([][]string, 0, 1000)
	maxCols := 0
	for row := range rows {
		rowCount++
		values := make([]string, len(row.Cells))
		for idx, cell := range row.Cells {
			values[idx] = strings.TrimSpace(cell.Value)
		}
		if len(values) > maxCols {
			maxCols = len(values)
		}
		allRows = append(allRows, values)
		if rowCount%100 == 0 || rowCount == 1000 {
			i.emitProgress("scan_sheet_rows", sheetName, rowCount, 1000)
		}
		if rowCount >= 1000 {
			break
		}
	}

	detail.RowCount = rowCount
	detail.ColumnCount = maxCols
	detail.Sections = extractSections(allRows)

	if len(detail.Sections) > 0 {
		detail.Headers = detail.Sections[0].Headers
		detail.Columns = detail.Sections[0].Columns
		return detail
	}

	// Fallback for simple single-table sheets.
	headerRow := findFirstNonEmptyRow(allRows)
	if headerRow == 0 {
		return detail
	}
	headers := allRows[headerRow-1]
	detail.Headers = trimTrailingEmpty(headers)
	detail.ColumnCount = len(detail.Headers)
	detail.Columns = buildColumnsFromSection(allRows, headerRow, detail.Headers, 5, rowCount)
	return detail
}

func (i *Inspector) visibleSheets(sheets []string) []string {
	out := make([]string, 0, len(sheets))
	for _, sheetName := range sheets {
		visible, err := i.file.GetSheetVisible(sheetName)
		if err != nil || !visible {
			continue
		}
		out = append(out, sheetName)
	}
	return out
}

func (i *Inspector) emitProgress(phase, sheet string, current, total int) {
	if i.progressCallback == nil && i.progressChan == nil {
		return
	}
	pct := 0.0
	if total > 0 {
		pct = (float64(current) / float64(total)) * 100.0
		if pct < 0 {
			pct = 0
		}
		if pct > 100 {
			pct = 100
		}
	}
	info := ProgressInfo{
		Phase:   phase,
		Sheet:   sheet,
		Current: current,
		Total:   total,
		Percent: pct,
	}
	if i.progressCallback != nil {
		i.progressCallback(info)
	}
	if i.progressChan != nil {
		select {
		case i.progressChan <- info:
		default:
		}
	}
}

func extractSections(rows [][]string) []Section {
	reportHeaderIdx := findReportHeaderRows(rows)
	if len(reportHeaderIdx) > 0 {
		sections := extractSectionsByHeaderIndexes(rows, reportHeaderIdx)
		return mergeReportSections(sections)
	}

	return extractSectionsByHeuristic(rows)
}

func extractSectionsByHeuristic(rows [][]string) []Section {
	sections := make([]Section, 0)
	for i := 0; i < len(rows); i++ {
		current := trimTrailingEmpty(rows[i])
		if !isLikelyHeaderRow(current) {
			continue
		}

		title := sectionTitleFromRow(rows, i-1)

		start := i + 2
		end := start - 1
		for j := start; j <= len(rows); j++ {
			if j == len(rows) {
				end = j
				break
			}
			next := trimTrailingEmpty(rows[j])
			if isLikelyHeaderRow(next) {
				end = j
				break
			}
			if isEmptyRow(next) && hasConsecutiveBlankRows(rows, j, 2) {
				end = j
				break
			}
			end = j + 1
		}

		headerRow := i + 1
		headers := current
		cols := buildColumnsFromSection(rows, headerRow, headers, 5, end)
		section := Section{
			Title:       title,
			HeaderRow:   headerRow,
			StartRow:    start,
			EndRow:      end,
			Headers:     headers,
			Columns:     cols,
			RowCount:    max(0, end-start+1),
			ColumnCount: len(headers),
		}
		sections = append(sections, section)
		i = end - 1
	}
	return sections
}

func extractSectionsByHeaderIndexes(rows [][]string, headerIdx []int) []Section {
	sections := make([]Section, 0, len(headerIdx))
	for idx, rowIdx := range headerIdx {
		headers := trimTrailingEmpty(rows[rowIdx])
		headerRow := rowIdx + 1
		start := headerRow + 1
		endExclusive := len(rows)
		if idx+1 < len(headerIdx) {
			endExclusive = headerIdx[idx+1]
		}
		end := endExclusive
		for end > rowIdx+1 && isEmptyRow(trimTrailingEmpty(rows[end-1])) {
			end--
		}

		title := sectionTitleFromRow(rows, rowIdx-1)

		section := Section{
			Title:       title,
			HeaderRow:   headerRow,
			StartRow:    start,
			EndRow:      end,
			Headers:     headers,
			Columns:     buildColumnsFromSection(rows, headerRow, headers, 5, end),
			RowCount:    max(0, end-start+1),
			ColumnCount: len(headers),
		}
		sections = append(sections, section)
	}
	return sections
}

func buildColumnsFromSection(rows [][]string, headerRow int, headers []string, maxSamples int, stopAtRow int) []ColumnInfo {
	columns := make([]ColumnInfo, len(headers))
	for colIdx, header := range headers {
		columns[colIdx] = ColumnInfo{
			Name:          header,
			StartPosition: fmt.Sprintf("%s%d", columnLetter(colIdx), headerRow),
			SampleValues:  make([]interface{}, 0, maxSamples),
		}
	}

	dataStart := headerRow + 1
	if stopAtRow <= 0 || stopAtRow > len(rows) {
		stopAtRow = len(rows)
	}
	for rowIdx := dataStart; rowIdx <= stopAtRow; rowIdx++ {
		row := rows[rowIdx-1]
		for colIdx := range headers {
			if colIdx >= len(row) {
				continue
			}
			v := strings.TrimSpace(row[colIdx])
			if v == "" || len(columns[colIdx].SampleValues) >= maxSamples {
				continue
			}
			columns[colIdx].SampleValues = append(columns[colIdx].SampleValues, v)
			if columns[colIdx].DataType == "" {
				columns[colIdx].DataType = getDataType(v)
			}
		}
	}
	return columns
}

func isLikelyHeaderRow(row []string) bool {
	if len(row) < 3 {
		return false
	}
	nonEmpty := 0
	known := 0
	for _, cell := range row {
		v := strings.TrimSpace(strings.ToUpper(cell))
		if v == "" {
			continue
		}
		nonEmpty++
		if isKnownHeaderToken(v) {
			known++
		}
	}
	if nonEmpty < 3 {
		return false
	}
	return known >= 2
}

func findReportHeaderRows(rows [][]string) []int {
	idx := make([]int, 0)
	for i, row := range rows {
		normalized := normalizeRow(row)
		if len(normalized) < 3 {
			continue
		}
		if !containsToken(normalized, "MERK") || !containsToken(normalized, "TYPE") {
			continue
		}
		if !containsAnyToken(normalized, "TRANSMITION", "TRANSMISSION", "YEAR", "ODOMETER", "STNK", "PURCHASE DATE") {
			continue
		}
		idx = append(idx, i)
	}
	return idx
}

func normalizeRow(row []string) []string {
	out := make([]string, 0, len(row))
	for _, cell := range row {
		v := strings.TrimSpace(strings.ToUpper(cell))
		if v != "" {
			out = append(out, v)
		}
	}
	return out
}

func containsToken(row []string, token string) bool {
	for _, v := range row {
		if v == token {
			return true
		}
	}
	return false
}

func containsAnyToken(row []string, tokens ...string) bool {
	for _, t := range tokens {
		if containsToken(row, t) {
			return true
		}
	}
	return false
}

func isKnownHeaderToken(v string) bool {
	switch v {
	case "MERK", "TYPE", "TRANSMITION", "TRANSMISSION", "YEAR", "COLOR", "ODOMETER", "STNK",
		"PURCHASE DATE", "AGING", "CREDIT PRICE", "CASH PRICE", "SELLING PRICE", "MARKET PRICE",
		"TOTAL NILAI STOCK (EST.)", "TOTAL NILAI STOCK (ACT.)", "NOTES DOCUMENT", "MR2",
		"NO", "STATUS", "PLATE NO", "PLATE NO ", "UNIT CATEGORY":
		return true
	default:
		return false
	}
}

func trimTrailingEmpty(row []string) []string {
	last := -1
	for i, cell := range row {
		if strings.TrimSpace(cell) != "" {
			last = i
		}
	}
	if last < 0 {
		return []string{}
	}
	return row[:last+1]
}

func isEmptyRow(row []string) bool {
	for _, cell := range row {
		if strings.TrimSpace(cell) != "" {
			return false
		}
	}
	return true
}

func hasConsecutiveBlankRows(rows [][]string, idx int, needed int) bool {
	count := 0
	for i := idx; i < len(rows); i++ {
		if isEmptyRow(trimTrailingEmpty(rows[i])) {
			count++
			if count >= needed {
				return true
			}
		} else {
			return false
		}
	}
	return false
}

func firstNonEmpty(row []string) string {
	for _, cell := range row {
		if strings.TrimSpace(cell) != "" {
			return cell
		}
	}
	return ""
}

func sectionTitleFromRow(rows [][]string, rowIdx int) string {
	if rowIdx < 0 || rowIdx >= len(rows) {
		return ""
	}
	tokens := make([]string, 0)
	for _, cell := range rows[rowIdx] {
		v := strings.TrimSpace(cell)
		if v != "" {
			tokens = append(tokens, v)
		}
	}
	if len(tokens) == 0 {
		return ""
	}

	upper := make([]string, len(tokens))
	for i, t := range tokens {
		upper[i] = strings.ToUpper(t)
	}

	if idx := indexOfContains(upper, "CROSS SELLING"); idx >= 0 {
		start := max(0, idx-2)
		return strings.Join(tokens[start:idx+1], " ")
	}
	if idx := indexOfContains(upper, "NON CROSS SELLING"); idx >= 0 {
		start := max(0, idx-2)
		return strings.Join(tokens[start:idx+1], " ")
	}
	if len(tokens) >= 2 {
		return strings.Join(tokens[:2], " ")
	}
	return tokens[0]
}

func indexOfContains(values []string, pattern string) int {
	for i, v := range values {
		if strings.Contains(v, pattern) {
			return i
		}
	}
	return -1
}

func mergeReportSections(sections []Section) []Section {
	if len(sections) == 0 {
		return sections
	}

	merged := make([]Section, 0, len(sections))
	indexByKey := make(map[string]int)
	for _, sec := range sections {
		key := reportSectionKey(sec)
		if key == "" {
			merged = append(merged, sec)
			continue
		}

		existingIdx, ok := indexByKey[key]
		if !ok {
			indexByKey[key] = len(merged)
			merged = append(merged, sec)
			continue
		}

		base := &merged[existingIdx]
		// Keep StartRow/EndRow from the first observed block to avoid
		// implying a continuous range when repeated blocks are non-contiguous.
		base.RowCount += sec.RowCount
		base.Columns = mergeColumnSamples(base.Columns, sec.Columns, 5)
	}
	return merged
}

func reportSectionKey(sec Section) string {
	title := strings.ToUpper(strings.TrimSpace(sec.Title))
	if !strings.Contains(title, "HANDOVER") {
		return ""
	}
	if !(strings.Contains(title, "CROSS SELLING") || strings.Contains(title, "NON CROSS SELLING")) {
		return ""
	}
	// Keep this strict so only HQ-like report sections are merged.
	return title + "|" + strings.Join(sec.Headers, "|")
}

func mergeColumnSamples(base []ColumnInfo, incoming []ColumnInfo, maxSamples int) []ColumnInfo {
	limit := len(base)
	if len(incoming) < limit {
		limit = len(incoming)
	}
	for i := 0; i < limit; i++ {
		if base[i].DataType == "" && incoming[i].DataType != "" {
			base[i].DataType = incoming[i].DataType
		}
		if base[i].SampleValues == nil {
			base[i].SampleValues = make([]interface{}, 0, maxSamples)
		}
		for _, sample := range incoming[i].SampleValues {
			if len(base[i].SampleValues) >= maxSamples {
				break
			}
			if !hasSample(base[i].SampleValues, sample) {
				base[i].SampleValues = append(base[i].SampleValues, sample)
			}
		}
	}
	return base
}

func hasSample(values []interface{}, target interface{}) bool {
	ts := fmt.Sprintf("%v", target)
	for _, v := range values {
		if fmt.Sprintf("%v", v) == ts {
			return true
		}
	}
	return false
}

func findFirstNonEmptyRow(rows [][]string) int {
	for idx, row := range rows {
		if !isEmptyRow(trimTrailingEmpty(row)) {
			return idx + 1
		}
	}
	return 0
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
