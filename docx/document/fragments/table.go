package fragments

import (
	"encoding/xml"
	"github.com/AlexsRyzhkov/freeoffice/docx/document/enums"
)

type ITable interface {
	AddRow()
	AddCol()

	GetRow(int) ITableRow
	GetCols(int) []ITableCol
	GetCol(int, int) ITableCol

	JoinCol(int, int, int)
}

type FTable struct {
	XMLName   xml.Name `xml:"w:tbl"`
	Property  FTableProperty
	TableGrid FTableGrid

	TableRows []*FTableRow
}

func (t *FTable) AddRow() {
	colNumber := len(t.TableGrid.GridCol)

	cols := make([]*FTableColumn, colNumber)

	for i := 0; i < colNumber; i++ {
		cols[i] = &FTableColumn{
			TableColumnProperty: TableColumnProperty{
				TcW: TcW{
					Width: enums.TableColumnWidthDefault / colNumber,
				},
				GridSpan: &GridSpan{
					Val: 1,
				},
			},
			TableColumnText: CreateFTextParagraph("", nil),
		}
	}

	t.TableRows = append(t.TableRows, &FTableRow{TableColumns: cols})
}

func (t *FTable) AddCol() {
	colNumber := len(t.TableGrid.GridCol) + 1

	t.TableGrid.GridCol = append(t.TableGrid.GridCol, new(FTableGridCol))

	for i := 0; i < len(t.TableGrid.GridCol); i++ {
		t.TableGrid.GridCol[i].Width = enums.TableColumnWidthDefault / colNumber
	}

	for i := 0; i < len(t.TableRows); i++ {
		t.TableRows[i].TableColumns = append(t.TableRows[i].TableColumns, &FTableColumn{
			TableColumnProperty: TableColumnProperty{
				TcW: TcW{
					Width: enums.TableColumnWidthDefault / colNumber,
				},
				GridSpan: &GridSpan{
					Val: 1,
				},
			},
			TableColumnText: CreateFTextParagraph("", nil),
		})
		for j := 0; j < len(t.TableRows[i].TableColumns); j++ {
			spanVal := t.TableRows[i].TableColumns[j].TableColumnProperty.GridSpan.Val
			colWidth := &t.TableRows[i].TableColumns[j].TableColumnProperty.TcW

			colWidth.Width = (enums.TableColumnWidthDefault / colNumber) * spanVal
		}
	}
}

func (t *FTable) JoinCol(row int, startCol int, endCol int) {
	row--
	startCol--
	endCol--

	if row < 0 {
		panic("row less than 0")
	}

	if row >= len(t.TableRows) {
		panic("row greater than len of table rows")
	}

	cols := t.TableRows[row].TableColumns

	if startCol < 0 {
		panic("start col less than 0")
	}

	if endCol >= len(cols) {
		panic("end col greater than len of cols")
	}

	if startCol > endCol {
		panic("startCol more than endCol")
	}

	colsNumber := len(t.TableGrid.GridCol)
	spanVal := 0

	for i := startCol; i <= endCol; i++ {
		spanVal += t.TableRows[row].TableColumns[i].TableColumnProperty.GridSpan.Val
	}

	newCols := make([]*FTableColumn, 0)
	for i := 0; i < len(cols); i++ {
		if i < startCol || i > endCol {
			newCols = append(newCols, cols[i])
		}

		if i >= startCol && i <= endCol {
			newCols = append(newCols, &FTableColumn{
				TableColumnProperty: TableColumnProperty{
					TcW: TcW{
						Width: (enums.TableColumnWidthDefault / colsNumber) * spanVal,
						Type:  enums.TableColumnType,
					},
					GridSpan: &GridSpan{
						Val: spanVal,
					},
				},
				TableColumnText: CreateFTextParagraph("", nil),
			})

			i = endCol
		}
	}
	t.TableRows[row].TableColumns = newCols
}

func (t *FTable) GetRow(row int) ITableRow {
	row--

	if row > len(t.TableRows) {
		return nil
	}

	return t.TableRows[row]
}

func (t *FTable) GetCols(row int) []ITableCol {
	row--

	columns := make([]ITableCol, 0, len(t.TableRows[row].TableColumns))

	for _, col := range t.TableRows[row].TableColumns {
		columns = append(columns, col)
	}

	return columns
}

func (t *FTable) GetCol(row int, col int) ITableCol {
	row--
	col--

	if row < 0 || row > len(t.TableRows) {
		return nil
	}

	if col < 0 || col > len(t.TableRows[row].TableColumns) {
		return nil
	}

	return t.TableRows[row].TableColumns[col]
}

type FTableProperty struct {
	XMLName  xml.Name `xml:"w:tblPr"`
	TblStyle FTableStyle
	TblW     FTableW
}

type FTableStyle struct {
	XMLName xml.Name `xml:"w:tblStyle"`
	Val     string   `xml:"w:val,attr"`
}

type FTableW struct {
	XMLName xml.Name `xml:"w:tblW"`
	Type    string   `xml:"w:type,attr"`
}

type FTableGrid struct {
	XMLName xml.Name `xml:"w:tblGrid"`
	GridCol []*FTableGridCol
}

type FTableGridCol struct {
	XMLName xml.Name `xml:"w:tblCol"`
	Width   int      `xml:"w:w,attr"`
}

type ITableRow interface {
	GetCol(int) ITableCol
}

type FTableRow struct {
	XMLName      xml.Name `xml:"w:tr"`
	TableColumns []*FTableColumn
}

func (tr *FTableRow) GetCol(col int) ITableCol {
	col--

	if col < 0 || col > len(tr.TableColumns) {
		return nil
	}

	return tr.TableColumns[col]
}

type ITableCol interface {
	GetText() ITextParagraph
	SetText(string) ITextParagraph
}

type FTableColumn struct {
	XMLName             xml.Name `xml:"w:tc"`
	TableColumnProperty TableColumnProperty
	TableColumnText     ITextParagraph
}

func (tc *FTableColumn) GetText() ITextParagraph {
	return tc.TableColumnText
}

func (tc *FTableColumn) SetText(text string) ITextParagraph {
	tc.TableColumnText.SetText(text)

	return tc.GetText()
}

type TableColumnProperty struct {
	XMLName  xml.Name `xml:"w:tcPr"`
	TcW      TcW
	GridSpan *GridSpan
}

type TcW struct {
	XMLName xml.Name `xml:"w:tcW"`
	Width   int      `xml:"w:w,attr"`
	Type    string   `xml:"w:type,attr"`
}

type GridSpan struct {
	XMLName xml.Name `xml:"w:gridSpan"`
	Val     int      `xml:"w:val,attr"`
}

func CreateFTable(row int, col int) ITable {
	if row < 1 {
		row = 1
	}

	if col < 1 {
		col = 1
	}

	gridCols := make([]*FTableGridCol, 0, col)
	for i := 0; i < col; i++ {
		gridCols = append(gridCols, &FTableGridCol{
			Width: enums.TableColumnWidthDefault / col,
		})
	}

	rows := make([]*FTableRow, row)
	for i := 0; i < row; i++ {
		rows[i] = &FTableRow{TableColumns: make([]*FTableColumn, col)}

		for j := 0; j < len(rows[i].TableColumns); j++ {
			rows[i].TableColumns[j] = &FTableColumn{
				TableColumnProperty: TableColumnProperty{
					TcW: TcW{
						Width: enums.TableColumnWidthDefault / col,
						Type:  enums.TableColumnType,
					},
					GridSpan: &GridSpan{
						Val: 1,
					},
				},
				TableColumnText: CreateFTextParagraph("", nil),
			}
		}
	}

	return &FTable{
		Property: FTableProperty{
			TblStyle: FTableStyle{
				Val: enums.TableStyleValDefault,
			},
			TblW: FTableW{
				Type: enums.TableWidthType,
			},
		},
		TableGrid: FTableGrid{
			GridCol: gridCols,
		},
		TableRows: rows,
	}
}
