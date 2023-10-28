package e2j

import (
	"errors"
	"fmt"
	"io"
	"os"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type RowType int

const (
	RowTypeHeader RowType = iota + 1
	RowTypeEmpty
	RowTypeData
	RowTypeComment
)

type Cell struct {
	Row     int
	Col     int
	Name    string
	Value   string
	Formula string
	Type    excelize.CellType
	StyleId int
}

type Row struct {
	Index int
	Type  RowType
	Cells []*Cell
}

type Column struct {
	Index int
	Name  string
	Width float64
}

type Sheet struct {
	Id        int
	Name      string
	Dimension string
	Rows      []*Row
	Cols      []*Column
	Formats   map[string][]excelize.ConditionalFormatOptions
}

func FillSheet(file *excelize.File, shname string, save_style func(style_id int), save_cond_style func(style_id int)) (*Sheet, error) {
	id, err := file.GetSheetIndex(shname)
	if err != nil {
		return nil, err
	}
	sh := Sheet{
		Id:   id,
		Name: shname,
	}
	sh_range, err := file.GetSheetDimension(shname)
	sh.Dimension = sh_range
	points := strings.Split(sh_range, ":")
	last_col_name, last_row, err := excelize.SplitCellName(points[1])
	if err != nil {
		return &sh, err
	}
	last_col, err := excelize.ColumnNameToNumber(last_col_name)
	if err != nil {
		return &sh, err
	}
	sh.Rows = make([]*Row, 0, last_row)
	for r := 1; r <= last_row; r++ {
		row := Row{
			Index: r,
			Type:  RowTypeEmpty,
			Cells: make([]*Cell, 0, last_col),
		}

		for c := 1; c <= last_col; c++ {
			cell_name, _ := excelize.CoordinatesToCellName(c, r)
			cell := Cell{}
			cell.Col = c
			cell.Row = r
			cell.Name = cell_name
			cell.Type, err = file.GetCellType(shname, cell_name)
			cell.Value, err = file.GetCellValue(shname, cell_name)
			cell.Formula, err = file.GetCellFormula(shname, cell_name)
			cell.StyleId, err = file.GetCellStyle(shname, cell_name)
			save_style(cell.StyleId)
			if cell.Value != "" || cell.Formula != "" {
				row.Type = RowTypeData
			}
			row.Cells = append(row.Cells, &cell)
		}
		if row.Index == 1 {
			row.Type = RowTypeHeader
		}
		if strings.HasPrefix(row.Cells[0].Value, "#") {
			row.Type = RowTypeComment
		}
		sh.Rows = append(sh.Rows, &row)
	}
	sh.Cols = make([]*Column, 0, last_col)
	for c := 1; c <= last_col; c++ {
		col := Column{
			Index: c,
		}
		col.Name, err = excelize.ColumnNumberToName(c)
		col.Width, err = file.GetColWidth(shname, col.Name)
		sh.Cols = append(sh.Cols, &col)
	}
	sh.Formats, err = file.GetConditionalFormats(shname)
	for _, options := range sh.Formats {
		for _, opt := range options {
			save_cond_style(opt.Format)
		}
	}
	return &sh, nil
}

type WorkBook struct {
	Sheets     []*Sheet
	Styles     map[int]*excelize.Style
	CondStyles map[int]*excelize.Style
	// JsonStyles map[int]string
}

func FillWorkBook(file *excelize.File) (*WorkBook, error) {
	wb := WorkBook{
		Sheets:     make([]*Sheet, 0, 10),
		Styles:     make(map[int]*excelize.Style),
		CondStyles: make(map[int]*excelize.Style),
		// JsonStyles: make(map[int]string),
	}
	err := wb.FillSheets(file, func(style_id int) {
		wb.FillStyles(file, style_id)
	}, func(style_id int) {
		wb.FillCondStyles(file, style_id)
	})
	return &wb, err
}

func (wb *WorkBook) FillSheets(file *excelize.File, save_style func(style_id int), save_cond_style func(style_id int)) error {
	var res error
	sh_map := file.GetSheetMap()
	keys := make([]int, 0, len(sh_map))
	for k := range sh_map {
		keys = append(keys, k)
	}
	sort.Ints(keys)
	for _, six := range keys {
		shname := sh_map[six]
		sh, err := FillSheet(file, shname, save_style, save_cond_style)
		if err == nil {
			wb.Sheets = append(wb.Sheets, sh)
		} else {
			res = err
			fmt.Println(err)
		}
	}
	return res
}

func (wb *WorkBook) FillStyles(file *excelize.File, style_id int) {
	if _, ok := wb.Styles[style_id]; !ok && style_id != 0 {
		if style, err := file.GetStyle(style_id); err == nil {
			wb.Styles[style_id] = style
		}
	}
}

func (wb *WorkBook) FillCondStyles(file *excelize.File, style_id int) {
	if _, ok := wb.CondStyles[style_id]; !ok {
		if style, err := file.GetConditionalStyle(style_id); err == nil {
			wb.CondStyles[style_id] = style
			// if dat, err := json.Marshal(style); err == nil {
			// 	wb.JsonStyles[style_id] = string(dat)
			// }
		}
	}
}

func (wb *WorkBook) StoreStyles(file *excelize.File) (map[int]int, error) {
	var res error
	styles_map := map[int]int{0: 0}
	keys := make([]int, 0, len(wb.Styles))
	for k := range wb.Styles {
		keys = append(keys, k)
	}
	sort.Ints(keys)
	for _, old_id := range keys {
		if v, ok := wb.Styles[old_id]; ok {
			if new_id, err := file.NewStyle(v); err == nil {
				styles_map[old_id] = new_id
			} else {
				res = err
				fmt.Println(err)
			}
		}
	}
	return styles_map, res
}

func (wb *WorkBook) StoreConditionalStyles(file *excelize.File) error {
	var res error
	keys := make([]int, 0, len(wb.CondStyles))
	for k := range wb.CondStyles {
		keys = append(keys, k)
	}
	sort.Ints(keys)
	for _, six := range keys {
		id, err := file.NewConditionalStyle(wb.CondStyles[six])
		if err != nil {
			res = err
			fmt.Println(err)
		} else if id != six {
			fmt.Println("Wrong index for created conditional style!")
			res = errors.New("Wrong index for created conditional style!")
		}
	}
	return res
}

func StoreSheet(file *excelize.File, six int, sheet *Sheet, styles_map map[int]int) error {
	var res error

	shname := sheet.Name
	if six == 0 {
		if err := file.SetSheetName("Sheet1", shname); err != nil {
			return err
		}
	} else {
		if _, err := file.NewSheet(shname); err != nil {
			return err
		}
	}
	if err := file.SetSheetDimension(shname, sheet.Dimension); err != nil {
		return err
	}
	for _, row := range sheet.Rows {
		for _, cell := range row.Cells {
			cell_name, err := excelize.CoordinatesToCellName(cell.Col, row.Index)
			if err == nil {
				if cell.Formula == "" {
					if cell.Type == excelize.CellTypeUnset {
						file.SetCellDefault(shname, cell_name, cell.Value)
					} else if cell.Type != excelize.CellTypeNumber {
						file.SetCellValue(shname, cell_name, cell.Value)
					} else {
						if f, err := strconv.ParseFloat(cell.Value, 64); err == nil {
							file.SetCellFloat(shname, cell_name, f, -1, 64)
						}
					}
				} else {
					file.SetCellFormula(shname, cell_name, cell.Formula)
				}
				file.SetCellStyle(shname, cell_name, cell_name, styles_map[cell.StyleId])
			}
		}
	}
	for _, col := range sheet.Cols {
		file.SetColWidth(shname, col.Name, col.Name, col.Width)
	}
	for cell_range, options := range sheet.Formats {
		if err := file.SetConditionalFormat(shname, cell_range, options); err != nil {
			fmt.Println(err)
			res = err
		}
	}
	file.UpdateLinkedValue()
	return res
}

func (wb *WorkBook) StoreSheets(file *excelize.File, styles_map map[int]int) error {
	var res error
	for six, sheet := range wb.Sheets {
		err := StoreSheet(file, six, sheet, styles_map)
		if err != nil {
			res = err
			fmt.Println(err)
		}
	}
	return res
}

func (wb *WorkBook) Store(file *excelize.File) error {
	var err error
	if styles_map, err := wb.StoreStyles(file); err == nil {
		if err = wb.StoreConditionalStyles(file); err == nil {
			err = wb.StoreSheets(file, styles_map)
		}
	}
	return err
}

func (wb *WorkBook) ToCsv(filename string, sep string) error {
	f, err := os.OpenFile(filename, os.O_RDWR|os.O_CREATE, 0755)
	if err != nil {
		fmt.Println(err)
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	keys := make([]int, 0, len(wb.Sheets))
	for k := range wb.Sheets {
		keys = append(keys, k)
	}
	sort.Ints(keys)
	for _, six := range keys {
		err := SheetToCsv(f, six, wb.Sheets[six])
		if err != nil {
			fmt.Println(err)
		}
	}

	return err
}

func SheetToCsv(w io.Writer, six int, sheet *Sheet) error {
	var res error

	shname := sheet.Name
	fmt.Fprintln(w, "--- Sheet: ", shname, " ---")
	for _, row := range sheet.Rows {
		values := make([]string, 0, len(row.Cells))
		for _, cell := range row.Cells {
			values = append(values, cell.Value)
		}
		r := strings.Join(values, ",")
		fmt.Fprintln(w, r)
	}
	fmt.Fprintln(w)
	return res
}
