package limentinus

import (
	"encoding/csv"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"
)

// Freeze/fix the first row of the given sheet in an excel file
func StaticHeader(excelFile string, sheetName string) (err error) {
	if err := validateInputs(excelFile, sheetName); err != nil {
		return err
	}

	f, err := excelize.OpenFile(excelFile)
	if err != nil {
		return fmt.Errorf("open excel file: %w", err)
	}
	defer func() {
		err = errors.Join(err, f.Close())
	}()

	if _, err := ensureSheetExists(f, sheetName); err != nil {
		return err
	}

	// Freeze the first row: y split 1, top-left cell is A2, active pane bottomLeft
	if err := f.SetPanes(sheetName, &excelize.Panes{
		Freeze:      true,
		YSplit:      1,
		TopLeftCell: "A2",
		ActivePane:  "bottomLeft",
	}); err != nil {
		return fmt.Errorf("set panes (freeze header): %w", err)
	}

	if err := f.Save(); err != nil {
		return fmt.Errorf("save excel file: %w", err)
	}
	return nil
}

// Enable data filtering on all the column that have values
// for first row of the given sheet in an excel file
func EnableData(excelFile string, sheetName string) (err error) {
	if err := validateInputs(excelFile, sheetName); err != nil {
		return err
	}

	f, err := excelize.OpenFile(excelFile)
	if err != nil {
		return fmt.Errorf("open excel file: %w", err)
	}
	defer func() {
		err = errors.Join(err, f.Close())
	}()

	if _, err := ensureSheetExists(f, sheetName); err != nil {
		return err
	}

	firstCol, lastCol, firstRow, lastRow, err := getUsedRangeByHeader(f, sheetName)
	if err != nil {
		return err
	}

	rangeRef := fmt.Sprintf("%s%d:%s%d", firstCol, firstRow, lastCol, lastRow)
	if err := f.AutoFilter(sheetName, rangeRef, nil); err != nil {
		return fmt.Errorf("apply auto filter to range %q: %w", rangeRef, err)
	}

	if err := f.Save(); err != nil {
		return fmt.Errorf("save excel file: %w", err)
	}
	return nil
}

// Resize to fit all the columns in the given sheet in an excel file
func AutoSize(excelFile string, sheetName string) (err error) {
	if err := validateInputs(excelFile, sheetName); err != nil {
		return err
	}

	f, err := excelize.OpenFile(excelFile)
	if err != nil {
		return fmt.Errorf("open excel file: %w", err)
	}
	defer func() {
		err = errors.Join(err, f.Close())
	}()

	if _, err := ensureSheetExists(f, sheetName); err != nil {
		return err
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("read rows: %w", err)
	}
	if len(rows) == 0 {
		return fmt.Errorf("sheet %q appears empty: nothing to auto-size", sheetName)
	}

	// Determine max column count by scanning rows
	maxCols := 0
	for _, r := range rows {
		if len(r) > maxCols {
			maxCols = len(r)
		}
	}
	if maxCols == 0 {
		return fmt.Errorf("no columns detected in sheet %q: nothing to auto-size", sheetName)
	}

	// Approximate Excel auto-fit:
	// - use the longest visual line width (runes) per cell
	// - add padding to avoid clipping and accommodate filter dropdown on header
	const (
		minWidth       = 8.0   // a reasonable minimum
		maxWidth       = 255.0 // Excel's effective max column width in characters
		contentPadding = 2.0   // general padding for readability
		filterPadding  = 2.0   // extra space to account for filter icon on header
	)

	for colIdx := 1; colIdx <= maxCols; colIdx++ {
		var maxVisual float64
		for rIdx, r := range rows {
			if colIdx-1 < len(r) {
				cell := r[colIdx-1]

				// Consider multi-line cells: pick the longest line
				var cellMax float64
				for _, line := range strings.Split(cell, "\n") {
					// Count runes as proxy for width; this is a simple, robust approximation
					w := float64(utf8.RuneCountInString(line))
					if w > cellMax {
						cellMax = w
					}
				}

				// Add additional padding on header to prevent filter icon overlapping text
				if rIdx == 0 && strings.TrimSpace(cell) != "" {
					cellMax += filterPadding
				}

				if cellMax > maxVisual {
					maxVisual = cellMax
				}
			}
		}

		// Add content padding and clamp
		width := maxVisual + contentPadding
		if width < minWidth {
			width = minWidth
		}
		if width > maxWidth {
			width = maxWidth
		}

		// Convert col index to letter and apply width
		colName, convErr := excelize.ColumnNumberToName(colIdx)
		if convErr != nil {
			return fmt.Errorf("convert column index %d: %w", colIdx, convErr)
		}
		if err := f.SetColWidth(sheetName, colName, colName, width); err != nil {
			return fmt.Errorf("set width for column %s: %w", colName, err)
		}
	}

	if err := f.Save(); err != nil {
		return fmt.Errorf("save excel file: %w", err)
	}
	return nil
}

// Applies the Accent1 color style to the first row of the given sheet in an excel file
func Accent1(excelFile string, sheetName string) (err error) {
	if err := validateInputs(excelFile, sheetName); err != nil {
		return err
	}

	f, err := excelize.OpenFile(excelFile)
	if err != nil {
		return fmt.Errorf("open excel file: %w", err)
	}
	defer func() {
		err = errors.Join(err, f.Close())
	}()

	if _, err := ensureSheetExists(f, sheetName); err != nil {
		return err
	}

	firstCol, lastCol, firstRow, _, err := getUsedRangeByHeader(f, sheetName)
	if err != nil {
		return err
	}

	// Create a style similar to Excel's Accent1 for headers: bold text and solid light accent fill
	styleID, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
			// default font color black
			Color: "#000000",
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
			WrapText:   true,
		},
		Fill: excelize.Fill{
			// Solid light blue fill (Accent1-like)
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{"#D9E1F2"},
		},
		Border: []excelize.Border{
			{Type: "left", Color: "#BFBFBF", Style: 1},
			{Type: "right", Color: "#BFBFBF", Style: 1},
			{Type: "top", Color: "#BFBFBF", Style: 1},
			{Type: "bottom", Color: "#BFBFBF", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("create style: %w", err)
	}

	start := fmt.Sprintf("%s%d", firstCol, firstRow)
	end := fmt.Sprintf("%s%d", lastCol, firstRow)

	if err := f.SetCellStyle(sheetName, start, end, styleID); err != nil {
		return fmt.Errorf("apply style to header range %s:%s: %w", start, end, err)
	}

	if err := f.Save(); err != nil {
		return fmt.Errorf("save excel file: %w", err)
	}
	return nil
}

func ApplyReportingFormat(excelFile string, sheetName string) (err error) {
	if err := validateInputs(excelFile, sheetName); err != nil {
		return err
	}
	// Freeze/fix the row
	if err = StaticHeader(excelFile, sheetName); err != nil {
		return err
	}
	// Apply color style to the first row
	if err = Accent1(excelFile, sheetName); err != nil {
		return err
	}
	// Enable data filtering on all the columns that have values in the first row
	if err = EnableData(excelFile, sheetName); err != nil {
		return err
	}
	// Resize to fit all the columns
	if err = AutoSize(excelFile, sheetName); err != nil {
		return err
	}
	return nil
}

// Convert a CSV file to an Excel file.
// The first row is treated as header. The separator must be ',' or ';'.
// The resulting workbook will contain a single sheet named sheetName.
func CSVToExcel(csvFile, excelFile, sheetName string, sep rune) (err error) {
	// Validate separator
	if sep != ',' && sep != ';' {
		return fmt.Errorf("unsupported separator %q: only ',' or ';' are allowed", string(sep))
	}
	// Validate sheet name
	if strings.TrimSpace(sheetName) == "" {
		return errors.New("sheet name must not be empty")
	}
	// Validate source CSV
	if strings.TrimSpace(csvFile) == "" {
		return errors.New("csv file path must not be empty")
	}
	info, err := os.Stat(csvFile)
	if err != nil {
		return fmt.Errorf("stat csv file: %w", err)
	}
	if info.IsDir() {
		return fmt.Errorf("csv path %q is a directory, expected a file", csvFile)
	}
	// Validate destination Excel path and ensure directory exists
	if strings.TrimSpace(excelFile) == "" {
		return errors.New("excel destination path must not be empty")
	}
	destDir := filepath.Dir(excelFile)
	if destDir != "." && destDir != "" {
		if mkErr := os.MkdirAll(destDir, 0o755); mkErr != nil {
			return fmt.Errorf("ensure destination directory: %w", mkErr)
		}
	}

	// Open CSV
	cf, err := os.Open(csvFile)
	if err != nil {
		return fmt.Errorf("open csv: %w", err)
	}
	defer func() {
		err = errors.Join(err, cf.Close())
	}()

	reader := csv.NewReader(cf)
	reader.Comma = sep
	reader.FieldsPerRecord = -1 // allow variable field counts per row

	// Prepare Excel
	f := excelize.NewFile()
	defer func() {
		err = errors.Join(err, f.Close())
	}()

	// Ensure the sheet has the requested name
	defaultSheet := f.GetSheetName(0)
	if defaultSheet != sheetName {
		if rnErr := f.SetSheetName(defaultSheet, sheetName); rnErr != nil {
			return fmt.Errorf("set sheet name: %w", rnErr)
		}
	}

	// Write cells row by row as literal strings to avoid formula execution.
	rowIdx := 1
	wroteAny := false
	for {
		record, rErr := reader.Read()
		if rErr != nil {
			if errors.Is(rErr, io.EOF) || rErr.Error() == "EOF" {
				break
			}
			return fmt.Errorf("read csv: %w", rErr)
		}
		wroteAny = true

		for colIdx, v := range record {
			colName, convErr := excelize.ColumnNumberToName(colIdx + 1)
			if convErr != nil {
				return fmt.Errorf("convert column number %d: %w", colIdx+1, convErr)
			}
			cell := fmt.Sprintf("%s%d", colName, rowIdx)
			// Force string cell to prevent formula evaluation while preserving exact text
			if err := f.SetCellStr(sheetName, cell, v); err != nil {
				return fmt.Errorf("set cell %s: %w", cell, err)
			}
		}
		rowIdx++
	}

	if !wroteAny {
		return fmt.Errorf("csv file %q is empty", csvFile)
	}

	// Save the Excel file
	if err := f.SaveAs(excelFile); err != nil {
		return fmt.Errorf("save excel file: %w", err)
	}
	return nil
}

// sanitizeExcelString neutralizes potential formula injection vectors in Excel by
// prefixing values that start with =, +, -, or @ with an apostrophe.
func sanitizeExcelString(s string) string {
	trim := s
	if trim == "" {
		return s
	}
	switch trim[0] {
	case '=', '+', '-', '@':
		return "'" + s
	default:
		return s
	}
}

// validateInputs performs basic validation for file path and sheet name to avoid obvious misuse.
func validateInputs(excelFile, sheetName string) error {
	if strings.TrimSpace(excelFile) == "" {
		return errors.New("excel file path must not be empty")
	}
	if strings.TrimSpace(sheetName) == "" {
		return errors.New("sheet name must not be empty")
	}
	info, err := os.Stat(excelFile)
	if err != nil {
		return fmt.Errorf("stat excel file: %w", err)
	}
	if info.IsDir() {
		return fmt.Errorf("path %q is a directory, expected an .xlsx file", excelFile)
	}
	return nil
}

// ensureSheetExists checks that a sheet exists and returns its index
func ensureSheetExists(f *excelize.File, sheetName string) (int, error) {
	idx, err := f.GetSheetIndex(sheetName)
	if err != nil {
		return 0, fmt.Errorf("get sheet index for %q: %w", sheetName, err)
	}
	if idx < 0 {
		return 0, fmt.Errorf("sheet %q does not exist", sheetName)
	}
	return idx, nil
}

// getUsedRangeByHeader determines the used range based on the header row (row 1).
// It returns start column "A", last used header column, first row number (1), and last row number (len(rows)).
func getUsedRangeByHeader(f *excelize.File, sheetName string) (firstCol string, lastCol string, firstRow int, lastRow int, err error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return "", "", 0, 0, fmt.Errorf("read rows: %w", err)
	}
	if len(rows) == 0 {
		return "", "", 0, 0, fmt.Errorf("sheet %q appears empty: header not found", sheetName)
	}

	header := rows[0]
	// Find last non-empty cell in header
	lastIdx := -1
	for i := len(header) - 1; i >= 0; i-- {
		if strings.TrimSpace(header[i]) != "" {
			lastIdx = i
			break
		}
	}
	if lastIdx == -1 {
		return "", "", 0, 0, fmt.Errorf("sheet %q header row is empty: nothing to operate on", sheetName)
	}

	colName, convErr := excelize.ColumnNumberToName(lastIdx + 1)
	if convErr != nil {
		return "", "", 0, 0, fmt.Errorf("convert column index %d: %w", lastIdx+1, convErr)
	}
	return "A", colName, 1, len(rows), nil
}
