package limentinus

import (
	"encoding/csv"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

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

	// Excel auto-fit approximation:
	// - compute visual width using heuristic per-rune weights (wide/narrow chars)
	// - consider longest visual line for multi-line cells
	// - ignore trailing spaces like Excel's auto-fit
	// - add padding to avoid clipping and accommodate filter dropdown on header
	const (
		minWidth         = 8.0   // a reasonable minimum
		maxWidth         = 255.0 // Excel's effective max column width in characters
		contentPadding   = 2.0   // general padding for readability
		filterPadding    = 2.0   // extra space to account for filter icon on header
		headerBoldFactor = 1.08  // conservative multiplier: bold text tends to render wider
	)

	// local helper to estimate visual width of a string in default Excel font
	visualWidth := func(s string) float64 {
		var w float64
		for _, r := range s {
			switch {
			case r == '\t':
				w += 4 // treat tab as a few spaces
			case r <= 0x007F: // basic Latin
				// Narrow punctuation and glyphs
				switch r {
				case ' ', '\'', '`', '.', ',', ':', ';', '!', '|':
					w += 0.5
				case 'i', 'l':
					w += 0.6
				case 'I':
					w += 0.7
				case '1':
					w += 0.8
				case '0', '2', '3', '4', '5', '6', '7', '8', '9':
					w += 0.9
				case 'W', 'M':
					w += 1.3
				default:
					w += 1.0
				}
			case (r >= 0x1100 && r <= 0x11FF) || // Hangul Jamo
				(r >= 0x2E80 && r <= 0x9FFF) || // CJK Radicals + Unified Ideographs
				(r >= 0xAC00 && r <= 0xD7AF) || // Hangul Syllables
				(r >= 0xF900 && r <= 0xFAFF) || // CJK Compatibility Ideographs
				(r >= 0xFE10 && r <= 0xFE6F) || // Vertical forms etc.
				(r >= 0xFF00 && r <= 0xFF60) || // Fullwidth forms
				(r >= 0x1F300 && r <= 0x1FAFF): // Emoji and symbols
				w += 2.0 // wide glyphs
			default:
				w += 1.4 // other non-ASCII assumed slightly wider
			}
		}
		return w
	}

	for colIdx := 1; colIdx <= maxCols; colIdx++ {
		var maxVisual float64
		for rIdx, r := range rows {
			if colIdx-1 < len(r) {
				cell := r[colIdx-1]

				// Consider multi-line cells: pick the longest visual line
				var cellMax float64
				for _, rawLine := range strings.Split(cell, "\n") {
					// Excel's auto-fit ignores trailing spaces when measuring
					line := strings.TrimRight(rawLine, " \t")
					vw := visualWidth(line)
					if vw > cellMax {
						cellMax = vw
					}
				}

				// Header tweaks: bold text typically renders wider and filter icon needs room
				if rIdx == 0 && strings.TrimSpace(cell) != "" {
					cellMax *= headerBoldFactor
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
	if err := validateSheetName(sheetName); err != nil {
		return err
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
		// Strip UTF-8 BOM from the first header cell if present
		if rowIdx == 1 && len(record) > 0 {
			record[0] = strings.TrimPrefix(record[0], "\uFEFF")
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
	if err := validateSheetName(sheetName); err != nil {
		return err
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

// validateSheetName ensures a sheet name conforms to Excel constraints for production safety.
// Rules: non-empty, max 31 characters, cannot contain: : \ / ? * [ ], and cannot start or end with an apostrophe.
func validateSheetName(name string) error {
	trimmed := strings.TrimSpace(name)
	if trimmed == "" {
		return errors.New("sheet name must not be empty")
	}
	// Excel prohibits leading/trailing apostrophes in sheet names
	if strings.HasPrefix(trimmed, "'") || strings.HasSuffix(trimmed, "'") {
		return fmt.Errorf("invalid sheet name %q: leading or trailing apostrophe is not allowed", name)
	}
	// Disallowed characters per Excel specification (spaces are allowed)
	for _, bad := range []rune{':', '\\', '/', '?', '*', '[', ']'} {
		if strings.ContainsRune(trimmed, bad) {
			return fmt.Errorf("invalid sheet name %q: contains one of the invalid characters : \\ / ? * [ ]", name)
		}
	}
	// Limit length to 31 characters (counting Unicode code points)
	if runeCount := len([]rune(trimmed)); runeCount > 31 {
		return fmt.Errorf("invalid sheet name %q: length %d exceeds Excel limit of 31 characters", name, runeCount)
	}
	return nil
}
