package xlsx

import (
	"fmt"
	"strconv"
	"strings"
)

type DataValidation struct {
	Type             string
	AllowBlank       string
	ShowInputMessage string
	ShowErrorMessage string
	Sqref            string
	Formula1         string
	Formula2         string
	StartRow         int
	StartCol         int
	EndRow           int
	EndCol           int
	ShtName          string
	DependRow        int
	DependCol        int
}

func (d *DataValidation) DecryptFormula() {
	var err error
	if strings.Compare(d.Type, "list") == 0 {
		dict := strings.Split(d.Formula1, ":")
		if len(dict) > 0 {
			d1 := strings.Split(dict[0], "$")
			if len(d1) != 3 {
				d.StartCol = -1
				d.StartRow = -1
				d.EndCol = -1
				d.EndRow = -1
				return
			} else {
				d.StartCol = ColLettersToIndex(d1[1])
				d.StartRow, err = strconv.Atoi(d1[2])
				if err != nil {
					d.StartCol = -1
					d.StartRow = -1
					d.EndCol = -1
					d.EndRow = -1
					return
				}
				d.StartRow--
			}
		}
		if len(dict) == 2 {
			d2 := strings.Split(dict[1], "$")
			if len(d2) != 3 {
				d.StartCol = -1
				d.StartRow = -1
				d.EndCol = -1
				d.EndRow = -1
				return
			} else {
				d.EndCol = ColLettersToIndex(d2[1])
				d.EndRow, err = strconv.Atoi(d2[2])
				if err != nil {
					d.StartCol = -1
					d.StartRow = -1
					d.EndCol = -1
					d.EndRow = -1
					return
				}
				d.EndRow--
			}
		} else {
			d.EndCol = d.StartCol
			d.EndRow = d.StartRow
		}
		if len(dict) != 1 && len(dict) != 2 {
			d.StartCol = -1
			d.StartRow = -1
			d.EndCol = -1
			d.EndRow = -1
			return
		}
	}
	return
}

// After user modifying values, EncryptFormula() must be called to reform Formula1.
func (d *DataValidation) EncryptFormula() {
	if strings.Compare(d.Type, "list") == 0 && d.StartRow != -1 {
		shtPrefix := ""
		if d.ShtName != "" {
			shtPrefix = d.ShtName + "!"
		}
		if d.StartRow == d.EndRow && d.StartCol == d.EndCol {
			d.Formula1 = shtPrefix + "$" + ColIndexToLetters(d.StartCol) + "$" + strconv.Itoa(d.StartRow+1)
		} else {
			d.Formula1 = shtPrefix + "$" + ColIndexToLetters(d.StartCol) + "$" + strconv.Itoa(d.StartRow+1) + ":$" + ColIndexToLetters(d.EndCol) + "$" + strconv.Itoa(d.EndRow+1)
		}
	}
}

func (d *DataValidation) EncryptLayeredFormula() {
	if strings.Compare(d.Type, "list") == 0 && d.StartRow != -1 {
		d.Formula1 = "INDIRECT(" + MakeColStr(d.DependCol) + strconv.Itoa(d.DependRow) + "&" + fmt.Sprintf("\"%v\"", MakeColStr(d.StartCol)) + ")"
	}
}
