package xlsx

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

type Hyperlink struct {
	CellValue string
	Formula   string
	ShtName   string
	Row       int
	Col       int
	StartRow  int
	StartCol  int
	EndRow    int
	EndCol    int
	Valid     bool
}

func GetHyperLink(formula string, rdx, cdx int) (*Hyperlink, bool) {
	match := regexp.MustCompile(`^HYPERLINK\(.*\)$`)
	if match.MatchString(formula) {
		return NewHyperlink(formula, rdx, cdx), true
	}
	return nil, false
}

func NewHyperlink(formula string, rdx, cdx int) *Hyperlink {
	h := &Hyperlink{Formula: formula, Row: rdx, Col: cdx}
	h.Valid = h.Decode()
	return h
}

func (h *Hyperlink) Decode() bool {
	coors := h.Formula[10 : len(h.Formula)-1]
	array := strings.Split(coors, ",")
	if len(array) != 2 {
		return false
	}
	h.CellValue = strings.ReplaceAll(array[1], "\"", "")
	array[0] = strings.ReplaceAll(array[0], "\"", "")
	array[0] = strings.ReplaceAll(array[0], "！", "!")
	array = strings.Split(array[0], "!")
	if len(array) != 2 {
		return false
	}
	if len(array[0]) <= 2 || array[0][0] != '#' {
		return false
	}
	h.ShtName = array[0][1:len(array[0])]
	array[1] = strings.ReplaceAll(array[1], "：", ":")
	areas := strings.Split(array[1], ":")
	if len(areas) < 1 {
		return false
	}
	var left, right string
	left = areas[0]
	if len(areas) == 1 {
		right = left
	} else if len(areas) == 2 {
		right = areas[1]
	} else {
		fmt.Println("coordinate error", areas)
		return false
	}
	sr, sc, err := ParseCoordinate(left)
	if err != nil {
		return false
	}
	h.StartRow = sr - 1
	h.StartCol = sc - 1
	er, ec, err := ParseCoordinate(right)
	if err != nil {
		return false
	}
	h.EndRow = er - 1
	h.EndCol = ec - 1
	return true
}

func ParseCoordinate(coordinate string) (int, int, error) {
	var row, col string
	for idx, c := range coordinate {
		if c >= '0' && c <= '9' {
			row = coordinate[:idx]
			col = coordinate[idx:]
			break
		}
	}
	if row == "" || col == "" {
		return 0, 0, fmt.Errorf("coordinate format error")
	}
	rowIdx, err := ParseRowNumber(row)
	if err != nil {
		return 0, 0, err
	}
	colIdx, err := strconv.Atoi(col)
	if err != nil {
		return 0, 0, err
	}
	return colIdx, rowIdx, nil
}
