package limentinus_test

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"

	limentinus "github.com/e6tUcu7c9h/limentinus"
	"github.com/xuri/excelize/v2"
)

func createTempWorkbook(t *testing.T, sheet string, data [][]interface{}) string {
	t.Helper()

	f := excelize.NewFile()
	// Ensure we use the requested sheet name
	defaultSheet := f.GetSheetName(0)
	if defaultSheet != sheet {
		if err := f.SetSheetName(defaultSheet, sheet); err != nil {
			t.Fatalf("rename sheet: %v", err)
		}
	}

	// Write data
	for rIdx, row := range data {
		for cIdx, v := range row {
			col, err := excelize.ColumnNumberToName(cIdx + 1)
			if err != nil {
				t.Fatalf("col name: %v", err)
			}
			cell := fmt.Sprintf("%s%d", col, rIdx+1)
			if err := f.SetCellValue(sheet, cell, v); err != nil {
				t.Fatalf("set cell %s: %v", cell, err)
			}
		}
	}

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "book.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save temp workbook: %v", err)
	}
	if err := f.Close(); err != nil {
		t.Fatalf("close workbook: %v", err)
	}
	return path
}

func TestStaticHeader_Valid(t *testing.T) {
	path := createTempWorkbook(t, "Sheet1", [][]interface{}{
		{"H1", "H2", "H3"},
		{1, 2, 3},
	})

	if err := limentinus.StaticHeader(path, "Sheet1"); err != nil {
		t.Fatalf("StaticHeader error: %v", err)
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("reopen: %v", err)
	}
	defer f.Close()

	panes, err := f.GetPanes("Sheet1")
	if err != nil {
		t.Fatalf("GetPanes: %v", err)
	}
	if !panes.Freeze {
		t.Fatalf("expected Freeze=true, got false")
	}
	if panes.YSplit != 1 {
		t.Fatalf("expected YSplit=1, got %d", panes.YSplit)
	}
	if panes.TopLeftCell != "A2" {
		t.Fatalf("expected TopLeftCell=A2, got %s", panes.TopLeftCell)
	}
}

func TestStaticHeader_InvalidInputs(t *testing.T) {
	if err := limentinus.StaticHeader("", "X"); err == nil {
		t.Fatal("expected error for empty path")
	}
	path := t.TempDir() // path is a directory
	if err := limentinus.StaticHeader(path, "X"); err == nil {
		t.Fatal("expected error for directory path")
	}
	path = createTempWorkbook(t, "S", [][]interface{}{{"A"}})
	if err := limentinus.StaticHeader(path, "Missing"); err == nil {
		t.Fatal("expected error for missing sheet")
	}
}

func TestEnableData_AppliesFilter(t *testing.T) {
	path := createTempWorkbook(t, "Data", [][]interface{}{
		{"Name", "Age", "City"},
		{"Alice", 30, "NY"},
		{"Bob", 25, "LA"},
	})

	if err := limentinus.EnableData(path, "Data"); err != nil {
		t.Fatalf("EnableData error: %v", err)
	}

	// Reopen to ensure file is valid
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("reopen: %v", err)
	}
	defer f.Close()

	// No GetAutoFilter API to assert range; ensure not corrupted and that applying another filter on same range succeeds
	if err := limentinus.EnableData(path, "Data"); err != nil {
		t.Fatalf("EnableData idempotent apply should not error: %v", err)
	}
}

func TestEnableData_EmptyHeaderErrors(t *testing.T) {
	path := createTempWorkbook(t, "Empty", [][]interface{}{})
	err := limentinus.EnableData(path, "Empty")
	if err == nil {
		t.Fatal("expected error for empty sheet")
	}
}

func TestAutoSize_SetsReasonableWidths(t *testing.T) {
	path := createTempWorkbook(t, "Sheet1", [][]interface{}{
		{"Header", "Very Very Very Long Header Value"},
		{"abc", "short"},
		{"some longer content", "this is a much longer content that should affect width"},
	})

	if err := limentinus.AutoSize(path, "Sheet1"); err != nil {
		t.Fatalf("AutoSize error: %v", err)
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("reopen: %v", err)
	}
	defer f.Close()

	wA, err := f.GetColWidth("Sheet1", "A")
	if err != nil {
		t.Fatalf("GetColWidth A: %v", err)
	}
	wB, err := f.GetColWidth("Sheet1", "B")
	if err != nil {
		t.Fatalf("GetColWidth B: %v", err)
	}
	if wA < 8.0 {
		t.Fatalf("expected column A width >= 8, got %f", wA)
	}
	if wB <= wA {
		t.Fatalf("expected column B wider than A due to longer content, got A=%f, B=%f", wA, wB)
	}
}

func TestAutoSize_CJKAndMultiline(t *testing.T) {
	path := createTempWorkbook(t, "Sheet1", [][]interface{}{
		{"ASCII", "标题\n第二行"}, // CJK header with newline
		{"abc", "数据"},         // CJK content should be wider than ASCII
	})
	if err := limentinus.AutoSize(path, "Sheet1"); err != nil {
		t.Fatalf("AutoSize error: %v", err)
	}
	f, err := excelize.OpenFile(path)
	if err != nil { t.Fatalf("reopen: %v", err) }
	defer f.Close()
	wA, _ := f.GetColWidth("Sheet1", "A")
	wB, _ := f.GetColWidth("Sheet1", "B")
	if wB <= wA {
		t.Fatalf("expected CJK column B to be wider than ASCII column A, got A=%f B=%f", wA, wB)
	}
}

func TestAccent1_AppliesStyleToHeader(t *testing.T) {
	path := createTempWorkbook(t, "Styled", [][]interface{}{
		{"C1", "C2", "C3"},
		{"v1", "v2", "v3"},
	})

	if err := limentinus.Accent1(path, "Styled"); err != nil {
		t.Fatalf("Accent1 error: %v", err)
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("reopen: %v", err)
	}
	defer f.Close()

	sA, err := f.GetCellStyle("Styled", "A1")
	if err != nil { t.Fatalf("GetCellStyle A1: %v", err) }
	sB, err := f.GetCellStyle("Styled", "B1")
	if err != nil { t.Fatalf("GetCellStyle B1: %v", err) }
	sC, err := f.GetCellStyle("Styled", "C1")
	if err != nil { t.Fatalf("GetCellStyle C1: %v", err) }

	if sA == 0 || sB == 0 || sC == 0 {
		t.Fatalf("expected non-zero style on header cells, got A=%d B=%d C=%d", sA, sB, sC)
	}
	if !(sA == sB && sB == sC) {
		t.Fatalf("expected same style across header cells, got A=%d B=%d C=%d", sA, sB, sC)
	}
}

func TestAccent1_EmptyHeaderErrors(t *testing.T) {
	path := createTempWorkbook(t, "Empty", [][]interface{}{})
	if err := limentinus.Accent1(path, "Empty"); err == nil {
		t.Fatal("expected error for empty header")
	}
}

func TestApplyReportingFormat_EndToEnd(t *testing.T) {
	path := createTempWorkbook(t, "Report", [][]interface{}{
		{"First Name", "Last Name", "Amount"},
		{"Alice", "W.", 100},
		{"Bob", "C.", 200},
	})
	if err := limentinus.ApplyReportingFormat(path, "Report"); err != nil {
		t.Fatalf("ApplyReportingFormat: %v", err)
	}
	f, err := excelize.OpenFile(path)
	if err != nil { t.Fatalf("reopen: %v", err) }
	defer f.Close()
	panes, err := f.GetPanes("Report")
	if err != nil { t.Fatalf("GetPanes: %v", err) }
	if !panes.Freeze || panes.YSplit != 1 { t.Fatalf("expected frozen header row") }
	style, err := f.GetCellStyle("Report", "A1")
	if err != nil || style == 0 { t.Fatalf("expected styled header") }
	wA, _ := f.GetColWidth("Report", "A")
	if wA < 8.0 { t.Fatalf("expected autosized columns, got A=%f", wA) }
}

func TestSecureHandling_InvalidPath(t *testing.T) {
	// Ensure safe error on missing file, not panic
	tmp := filepath.Join(t.TempDir(), "missing.xlsx")
	if _, err := os.Stat(tmp); !os.IsNotExist(err) {
		t.Fatalf("expected not exist: %v", err)
	}
	for _, fn := range []func(string, string) error{limentinus.StaticHeader, limentinus.EnableData, limentinus.AutoSize, limentinus.Accent1} {
		if err := fn(tmp, "Sheet1"); err == nil {
			t.Fatalf("expected error for missing path")
		}
	}
}

func TestCSVToExcel_CommaSeparator(t *testing.T) {
	dir := t.TempDir()
	csvPath := filepath.Join(dir, "data.csv")
	excelPath := filepath.Join(dir, "data.xlsx")
	content := "Name,Value\nAlice,=2+2\nBob,hello\n"
	if err := os.WriteFile(csvPath, []byte(content), 0o644); err != nil {
		t.Fatalf("write csv: %v", err)
	}

	if err := limentinus.CSVToExcel(csvPath, excelPath, "Sheet1", ','); err != nil {
		t.Fatalf("CSVToExcel: %v", err)
	}

	f, err := excelize.OpenFile(excelPath)
	if err != nil {
		t.Fatalf("open excel: %v", err)
	}
	defer f.Close()

	name, err := f.GetCellValue("Sheet1", "A1")
	if err != nil { t.Fatalf("get A1: %v", err) }
	valHeader, err := f.GetCellValue("Sheet1", "B1")
	if err != nil { t.Fatalf("get B1: %v", err) }
	if name != "Name" || valHeader != "Value" {
		t.Fatalf("unexpected headers: A1=%q B1=%q", name, valHeader)
	}

	// Ensure formula injection is preserved as literal (not executed)
	b2, err := f.GetCellValue("Sheet1", "B2")
	if err != nil { t.Fatalf("get B2: %v", err) }
	if b2 != "=2+2" { t.Fatalf("expected literal '=2+2' in B2, got %q", b2) }
}

func TestCSVToExcel_SemicolonSeparator(t *testing.T) {
	dir := t.TempDir()
	csvPath := filepath.Join(dir, "data.csv")
	excelPath := filepath.Join(dir, "data.xlsx")
	content := "A;B;C\n1;2;3\nx;y;z\n"
	if err := os.WriteFile(csvPath, []byte(content), 0o644); err != nil {
		t.Fatalf("write csv: %v", err)
	}

	if err := limentinus.CSVToExcel(csvPath, excelPath, "Sheet1", ';'); err != nil {
		t.Fatalf("CSVToExcel: %v", err)
	}

	f, err := excelize.OpenFile(excelPath)
	if err != nil { t.Fatalf("open excel: %v", err) }
	defer f.Close()

	a1, err := f.GetCellValue("Sheet1", "A1")
	if err != nil { t.Fatalf("get A1: %v", err) }
	c2, err := f.GetCellValue("Sheet1", "C2")
	if err != nil { t.Fatalf("get C2: %v", err) }
	if a1 != "A" || c2 != "3" { t.Fatalf("unexpected values: A1=%q C2=%q", a1, c2) }
}

func TestCSVToExcel_InvalidSeparator(t *testing.T) {
	dir := t.TempDir()
	csvPath := filepath.Join(dir, "data.csv")
	excelPath := filepath.Join(dir, "data.xlsx")
	content := "H1|H2\n1|2\n"
	if err := os.WriteFile(csvPath, []byte(content), 0o644); err != nil {
		t.Fatalf("write csv: %v", err)
	}
	if err := limentinus.CSVToExcel(csvPath, excelPath, "Sheet1", '|'); err == nil {
		t.Fatal("expected error for invalid separator")
	}
}

func TestCSVToExcel_EmptyFile(t *testing.T) {
	dir := t.TempDir()
	csvPath := filepath.Join(dir, "empty.csv")
	excelPath := filepath.Join(dir, "empty.xlsx")
	if err := os.WriteFile(csvPath, nil, 0o644); err != nil {
		t.Fatalf("write csv: %v", err)
	}
	if err := limentinus.CSVToExcel(csvPath, excelPath, "Sheet1", ','); err == nil {
		t.Fatal("expected error for empty csv")
	}
}

func TestCSVToExcel_HeadersWithSpacesAndQuotes(t *testing.T) {
	dir := t.TempDir()
	csvPath := filepath.Join(dir, "table data.csv") // path with space
	excelPath := filepath.Join(dir, "report data.xlsx") // excel path with space
	content := "\"First Name\",Last Name,Age\n\"Alice, A.\",Smith,30\nBob,Jones,40\n"
	if err := os.WriteFile(csvPath, []byte(content), 0o644); err != nil { t.Fatalf("write csv: %v", err) }
	if err := limentinus.CSVToExcel(csvPath, excelPath, "Report Data 2025", ','); err != nil {
		t.Fatalf("CSVToExcel with spaces failed: %v", err)
	}
	f, err := excelize.OpenFile(excelPath)
	if err != nil { t.Fatalf("open excel: %v", err) }
	defer f.Close()
	fn, _ := f.GetCellValue("Report Data 2025", "A1")
	ln, _ := f.GetCellValue("Report Data 2025", "B1")
	if fn != "First Name" || ln != "Last Name" {
		t.Fatalf("unexpected headers with spaces: %q / %q", fn, ln)
	}
}

func TestCSVToExcel_WithBOM_Trimmed(t *testing.T) {
	dir := t.TempDir()
	csvPath := filepath.Join(dir, "bom.csv")
	excelPath := filepath.Join(dir, "bom.xlsx")
	// UTF-8 BOM + header
	bom := []byte{0xEF, 0xBB, 0xBF}
	content := append(bom, []byte("Name,Val\nA,1\n")...)
	if err := os.WriteFile(csvPath, content, 0o644); err != nil { t.Fatalf("write csv: %v", err) }
	if err := limentinus.CSVToExcel(csvPath, excelPath, "Sheet1", ','); err != nil { t.Fatalf("CSVToExcel: %v", err) }
	f, err := excelize.OpenFile(excelPath)
	if err != nil { t.Fatalf("open excel: %v", err) }
	defer f.Close()
	a1, _ := f.GetCellValue("Sheet1", "A1")
	if a1 != "Name" { t.Fatalf("expected BOM trimmed header 'Name', got %q", a1) }
}

func TestCSVToExcel_IrregularRows(t *testing.T) {
	dir := t.TempDir()
	csvPath := filepath.Join(dir, "irregular.csv")
	excelPath := filepath.Join(dir, "irregular.xlsx")
	content := "H1,H2,H3\n1,2\n3,4,5,6\n"
	if err := os.WriteFile(csvPath, []byte(content), 0o644); err != nil { t.Fatalf("write csv: %v", err) }
	if err := limentinus.CSVToExcel(csvPath, excelPath, "Sheet1", ','); err != nil { t.Fatalf("CSVToExcel: %v", err) }
	f, err := excelize.OpenFile(excelPath)
	if err != nil { t.Fatalf("open excel: %v", err) }
	defer f.Close()
	c21, _ := f.GetCellValue("Sheet1", "C2")
	c24, _ := f.GetCellValue("Sheet1", "C3")
	if c21 != "" || c24 != "5" { t.Fatalf("unexpected irregular row values C2=%q C3=%q", c21, c24) }
}


func TestAutoSize_TrailingSpacesIgnored(t *testing.T) {
	// Create two similar workbooks, differing only by trailing spaces in a cell.
	path1 := createTempWorkbook(t, "S", [][]interface{}{
		{"Name"},
		{"Alice"},
	})
	path2 := createTempWorkbook(t, "S", [][]interface{}{
		{"Name"},
		{"Alice            "}, // many trailing spaces
	})

	if err := limentinus.AutoSize(path1, "S"); err != nil {
		t.Fatalf("AutoSize path1: %v", err)
	}
	if err := limentinus.AutoSize(path2, "S"); err != nil {
		t.Fatalf("AutoSize path2: %v", err)
	}

	f1, err := excelize.OpenFile(path1)
	if err != nil { t.Fatalf("open1: %v", err) }
	defer f1.Close()
	f2, err := excelize.OpenFile(path2)
	if err != nil { t.Fatalf("open2: %v", err) }
	defer f2.Close()

	w1, err := f1.GetColWidth("S", "A")
	if err != nil { t.Fatalf("GetColWidth A (w1): %v", err) }
	w2, err := f2.GetColWidth("S", "A")
	if err != nil { t.Fatalf("GetColWidth A (w2): %v", err) }

	// Expect negligible difference because Excel's auto-fit ignores trailing spaces.
	diff := w1 - w2
	if diff < 0 { diff = -diff }
	if diff > 0.5 {
		t.Fatalf("expected trailing spaces to have minimal impact on width (|Δ|<=0.5), got w1=%f w2=%f Δ=%f", w1, w2, diff)
	}
}

func TestCSVToExcel_InvalidSheetNames(t *testing.T) {
	dir := t.TempDir()
	csvPath := filepath.Join(dir, "data.csv")
	excelPath := filepath.Join(dir, "data.xlsx")
	if err := os.WriteFile(csvPath, []byte("A,B\n1,2\n"), 0o644); err != nil {
		t.Fatalf("write csv: %v", err)
	}
	invalid := []string{
		"Bad/Name",
		"Bad:Name",
		"Bad*Name",
		"Bad[Name]",
		"'LeadingApostrophe",
		"TrailingApostrophe'",
	}
	// name with 32 characters (exceeds 31 limit)
	tooLong := strings.Repeat("a", 32)
	invalid = append(invalid, tooLong)
	for _, name := range invalid {
		if err := limentinus.CSVToExcel(csvPath, excelPath, name, ','); err == nil {
			t.Fatalf("expected error for invalid sheet name %q", name)
		}
	}
}

func TestCSVToExcel_SheetName_MaxLenOK(t *testing.T) {
	dir := t.TempDir()
	csvPath := filepath.Join(dir, "data.csv")
	excelPath := filepath.Join(dir, "data.xlsx")
	if err := os.WriteFile(csvPath, []byte("A,B\n1,2\n"), 0o644); err != nil {
		t.Fatalf("write csv: %v", err)
	}
	name := strings.Repeat("x", 31) // boundary allowed length
	if err := limentinus.CSVToExcel(csvPath, excelPath, name, ','); err != nil {
		t.Fatalf("unexpected error for 31-char sheet name: %v", err)
	}
	// sanity read
	f, err := excelize.OpenFile(excelPath)
	if err != nil {
		t.Fatalf("open excel: %v", err)
	}
	defer f.Close()
	v, err := f.GetCellValue(name, "A2")
	if err != nil || v != "1" {
		t.Fatalf("unexpected read from sheet %q: v=%q err=%v", name, v, err)
	}
}
