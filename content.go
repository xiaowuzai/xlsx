package xlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"strconv"
	"strings"

	yymmdd "github.com/extrame/goyymmdd"
)

type SheetContent struct {
	Name      string
	Content   map[int]map[int]*Cell
	BlankCols map[int]struct{}
	BlankRows map[int]struct{}
	MaxRow    int
	MaxCol    int
	Rows      int
	Cols      int
}

type indexedSheetContent struct {
	Index        int
	SheetContent *SheetContent
	Error        error
}

type xlsxWorksheetContent struct {
	Dimension  xlsxDimension   `xml:"dimension"`
	Cols       *xlsxCols       `xml:"cols,omitempty"`
	SheetData  xlsxSheetData   `xml:"sheetData"`
	MergeCells *xlsxMergeCells `xml:"mergeCells,omitempty"`
}

func ReadSheetContents(bs []byte) ([]*SheetContent, error) {
	r := bytes.NewReader(bs)
	file, err := zip.NewReader(r, int64(r.Len()))
	if err != nil {
		return nil, err
	}
	return ReadZipReaderContentsWithRowLimit(file, NoRowLimit)
}

func ReadZipReaderContentsWithRowLimit(r *zip.Reader, rowLimit int) ([]*SheetContent, error) {
	var err error
	var reftable *RefTable
	var sharedStrings *zip.File
	var sheetXMLMap map[string]string
	var sheets []*SheetContent
	var style *xlsxStyleSheet
	var v *zip.File
	var theme *theme
	var styles *zip.File
	var themeFile *zip.File
	var workbook *zip.File
	var workbookRels *zip.File
	var worksheets map[string]*zip.File

	worksheets = make(map[string]*zip.File, len(r.File))
	for _, v = range r.File {
		switch v.Name {
		case "xl/sharedStrings.xml":
			sharedStrings = v
		case "xl/workbook.xml":
			workbook = v
		case "xl/_rels/workbook.xml.rels":
			workbookRels = v
		case "xl/styles.xml":
			styles = v
			// break
		case "xl/theme/theme1.xml":
			themeFile = v
			// break
		default:
			if len(v.Name) > 14 {
				if v.Name[0:13] == "xl/worksheets" {
					worksheets[v.Name[14:len(v.Name)-4]] = v
				}
			}
		}
	}
	if workbookRels == nil {
		return nil, fmt.Errorf("xl/_rels/workbook.xml.rels not found in input xlsx.")
	}
	sheetXMLMap, err = readWorkbookRelationsFromZipFile(workbookRels)
	if err != nil {
		return nil, err
	}
	if len(worksheets) == 0 {
		return nil, fmt.Errorf("Input xlsx contains no worksheets.")
	}
	reftable, err = readSharedStringsFromZipFile(sharedStrings)
	if err != nil {
		return nil, err
	}
	if themeFile != nil {
		theme, err = readThemeFromZipFile(themeFile)
		if err != nil {
			return nil, err
		}
	}
	if styles != nil {
		style, err = readStylesFromZipFile(styles, theme)
		if err != nil {
			return nil, err
		}
	}
	sheets, err = readSheetContentsFromZipFile(workbook, worksheets, style, reftable, sheetXMLMap, rowLimit)
	if err != nil {
		return nil, err
	}
	if sheets == nil {
		readerErr := new(XLSXReaderError)
		readerErr.Err = "No sheets found in XLSX File"
		return nil, readerErr
	}
	return sheets, nil
}

func readSheetContentsFromZipFile(f *zip.File, worksheets map[string]*zip.File, style *xlsxStyleSheet, reftable *RefTable, sheetXMLMap map[string]string, rowLimit int) ([]*SheetContent, error) {
	var err error
	var rc io.ReadCloser
	var decoder *xml.Decoder
	var sheetCount int
	var workbook = new(xlsxWorkbook)
	rc, err = f.Open()
	if err != nil {
		return nil, err
	}
	decoder = xml.NewDecoder(rc)
	err = decoder.Decode(workbook)
	if err != nil {
		return nil, err
	}
	// Only try and read sheets that have corresponding files.
	// Notably this excludes chartsheets don't right now
	var workbookSheets []xlsxSheet
	for _, sheet := range workbook.Sheets.Sheet {
		if f := worksheetFileForSheet(sheet, worksheets, sheetXMLMap); f != nil {
			workbookSheets = append(workbookSheets, sheet)
		}
	}
	sheetCount = len(workbookSheets)
	sheets := make([]*SheetContent, sheetCount)
	sheetChan := make(chan *indexedSheetContent, sheetCount)

	go func() {
		defer close(sheetChan)
		err = nil
		for i, rawsheet := range workbookSheets {
			if err := readSheetContentFromFile(sheetChan, i, worksheets, style, reftable, rawsheet, sheetXMLMap, rowLimit); err != nil {
				return
			}
		}
	}()

	for j := 0; j < sheetCount; j++ {
		sheet := <-sheetChan
		if sheet.Error != nil {
			return nil, sheet.Error
		}
		sheetName := workbookSheets[sheet.Index].Name
		sheet.SheetContent.Name = sheetName
		sheets[sheet.Index] = sheet.SheetContent
	}
	return sheets, nil
}

func readSheetContentFromFile(sc chan *indexedSheetContent, index int, worksheets map[string]*zip.File, style *xlsxStyleSheet, reftable *RefTable, rsheet xlsxSheet, sheetXMLMap map[string]string, rowLimit int) (errRes error) {
	result := &indexedSheetContent{Index: index, SheetContent: nil, Error: nil}
	defer func() {
		if e := recover(); e != nil {
			switch e.(type) {
			case error:
				result.Error = e.(error)
				errRes = e.(error)
			default:
				result.Error = errors.New("unexpected error")
			}
			// The only thing here, is if one close the channel. but its not the case
			sc <- result
		}
	}()

	worksheet, err := getWorksheetContentFromSheet(rsheet, worksheets, sheetXMLMap, rowLimit)
	if err != nil {
		result.Error = err
		sc <- result
		return err
	}
	result.SheetContent = readContentFromSheet(worksheet, style, reftable)
	sc <- result
	return nil
}

func readContentFromSheet(worksheet *xlsxWorksheetContent, style *xlsxStyleSheet, reftable *RefTable) *SheetContent {
	var maxCol, maxRow, cols, rows int

	colis := map[int]struct{}{}
	contents := make(map[int]map[int]*Cell)
	sharedFormulas := map[int]sharedFormula{}
	valuesRows := make(map[int]struct{})

	for rowIndex := 0; rowIndex < len(worksheet.SheetData.Row); rowIndex++ {
		rawrow := worksheet.SheetData.Row[rowIndex]
		var row = map[int]*Cell{}
		for _, rawcell := range rawrow.C {
			h, v, err := worksheet.MergeCells.getExtent(rawcell.R)
			if err != nil {
				panic(err.Error())
			}
			x, _, _ := GetCoordsFromCellIDString(rawcell.R)
			cell := new(Cell)
			cell.HMerge = h
			cell.VMerge = v
			if style != nil {
				cell.style = style.getStyle(rawcell.S)
				cell.NumFmt, cell.parsedNumFmt = style.getNumberFormat(rawcell.S)
			}
			fillCellString(rawcell, reftable, sharedFormulas, cell)
			row[x] = cell
			if x > maxCol {
				maxCol = x
			}
			colis[x] = struct{}{}
			if h > 0 {
				for mx := 0; mx <= h; mx++ {
					colis[x+mx] = struct{}{}
				}
			}
			if cell.Value != "" {
				valuesRows[rowIndex] = struct{}{}
			}
		}
		contents[rowIndex] = row
		if maxRow < rowIndex {
			maxRow = rowIndex
		}
	}

	for rowIdx, row := range contents {
		for _, cell := range row {
			if cell.VMerge > 0 {
				valueRelated := false
				for midx := 0; midx <= cell.VMerge; midx++ {
					if _, has := valuesRows[rowIdx+midx]; has {
						valueRelated = true
						break
					}
				}
				if valueRelated {
					for midx := 0; midx <= cell.VMerge; midx++ {
						valuesRows[rowIdx+midx] = struct{}{}
					}
				}
			}
		}
	}
	rows = len(contents)
	cols = len(colis)
	blankCols := map[int]struct{}{}
	for i := 0; i <= maxCol; i++ {
		if _, has := colis[i]; !has {
			blankCols[i] = struct{}{}
		}
	}
	blankRows := map[int]struct{}{}
	for rowIdx := 0; rowIdx < len(contents); rowIdx++ {
		if _, has := valuesRows[rowIdx]; !has {
			blankRows[rowIdx] = struct{}{}
		}
	}

	sheet := new(SheetContent)
	sheet.Content = contents
	sheet.BlankCols = blankCols
	sheet.MaxCol = maxCol
	sheet.MaxRow = maxRow
	sheet.Cols = cols
	sheet.Rows = rows
	sheet.BlankRows = blankRows

	return sheet
}

func fillCellString(rawCell xlsxC, refTable *RefTable, sharedFormulas map[int]sharedFormula, cell *Cell) {
	val := strings.Trim(rawCell.V, " \t\n\r")
	cell.formula = formulaForCell(rawCell, sharedFormulas)
	switch rawCell.T {
	case "s": // Shared String
		cell.cellType = CellTypeString
		if val != "" {
			ref, err := strconv.Atoi(val)
			if err != nil {
				panic(err)
			}
			cell.Value = refTable.ResolveSharedString(ref)
		}
	case "inlineStr":
		cell.cellType = CellTypeInline
		fillCellDataFromInlineString(rawCell, cell)
	case "b": // Boolean
		cell.Value = val
		cell.cellType = CellTypeBool
	case "e": // Error
		cell.Value = val
		cell.cellType = CellTypeError
	case "str":
		cell.Value = val
		cell.cellType = CellTypeStringFormula
	case "d": // Date: Cell contains a date in the ISO 8601 format.
		cell.Value = val
		cell.cellType = CellTypeDate
	case "": // Numeric is the default
		if cell.parsedNumFmt.isTimeFormat {
			f, err := strconv.ParseFloat(val, 64)
			if err == nil {
				val := TimeFromExcelTime(f, false)
				cell.Value = yymmdd.Format(val, "2006-01-02")
				cell.cellType = CellTypeDate
				return
			}
		}
		fallthrough
	case "n": // Numeric
		cell.Value = val
		cell.cellType = CellTypeNumeric
	default:
		cell.Value = val
	}
}

func getWorksheetContentFromSheet(sheet xlsxSheet, worksheets map[string]*zip.File, sheetXMLMap map[string]string, rowLimit int) (*xlsxWorksheetContent, error) {
	var r io.Reader
	var decoder *xml.Decoder
	var worksheet *xlsxWorksheetContent
	var err error
	worksheet = new(xlsxWorksheetContent)

	f := worksheetFileForSheet(sheet, worksheets, sheetXMLMap)
	if f == nil {
		return nil, fmt.Errorf("Unable to find sheet '%s'", sheet)
	}
	if rc, err := f.Open(); err != nil {
		return nil, err
	} else {
		defer rc.Close()
		r = rc
	}

	if rowLimit != NoRowLimit {
		r, err = truncateSheetXML(r, rowLimit)
		if err != nil {
			return nil, err
		}
	}

	decoder = xml.NewDecoder(r)
	err = decoder.Decode(worksheet)
	if err != nil {
		return nil, err
	}
	return worksheet, nil
}
