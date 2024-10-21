package fragments

import (
	"encoding/xml"
	"github.com/AlexsRyzhkov/freeoffice/docx/document/enums"
)

const (
	orientationPortrait  = "portrait"
	orientationLandscape = "landscape"
)

const (
	defaultMarginTop    = "1134"
	defaultMarginRight  = "850"
	defaultMarginBottom = "1134"
	defaultMarginLeft   = "1701"

	defaultMarginHeader = "708"
	defaultMarginFooter = "708"
	defaultMarginGutter = "0"
)

const (
	defaultColsSpace = "708"
)

type Paragraph interface{}

type IBody interface {
	AddParagraph(Paragraph)
}

type FBody struct {
	XMLName         xml.Name `xml:"w:body"`
	Paragraphs      []Paragraph
	SectionProperty *SectionProperty
}

func (fb *FBody) AddParagraph(p Paragraph) {
	fb.Paragraphs = append(fb.Paragraphs, p)
}

type SectionProperty struct {
	XMLName    xml.Name `xml:"w:sectPr"`
	FootnotePr string   `xml:"w:footnotePr"`
	EndnotePR  string   `xml:"w:endnotePr"`
	Type       *SectionPropertyType
	PageSize   *PageSize
	PageMargin *PageMargin
	Cols       *Cols
	DocGrid    *DocGrid
}

// Настройка свойств документа
type SectionPropertyType struct {
	XMlName xml.Name `xml:"w:type"`
	Val     string   `xml:"w:val"`
}

type PageSize struct {
	XMLName xml.Name `xml:"w:pgSz"`
	Height  int      `xml:"w:h,attr"`
	Width   int      `xml:"w:w,attr"`
	Orient  string   `xml:"w:orient,attr"`
}

type PageMargin struct {
	XMLName xml.Name `xml:"w:pgMar"`
	Top     string   `xml:"w:top,attr"`
	Right   string   `xml:"w:right,attr"`
	Bottom  string   `xml:"w:bottom,attr"`
	Left    string   `xml:"w:left,attr"`
	Header  string   `xml:"w:header,attr"`
	Footer  string   `xml:"w:footer,attr"`
	Gutter  string   `xml:"w:gutter,attr"`
}

type Cols struct {
	XMLName xml.Name `xml:"w:cols"`
	Space   string   `xml:"w:space,attr"`
}

type DocGrid struct {
	XMLName   xml.Name `xml:"w:docGrid"`
	LinePitch string   `xml:"w:linePitch,attr"`
}

func CreateFBody() *FBody {
	return &FBody{
		Paragraphs: nil,
		SectionProperty: &SectionProperty{
			PageSize: &PageSize{
				Width:  enums.DefaultPageWidth,
				Height: enums.DefaultPageHeight,
				Orient: orientationPortrait,
			},
			PageMargin: &PageMargin{
				Top:    defaultMarginTop,
				Right:  defaultMarginRight,
				Bottom: defaultMarginBottom,
				Left:   defaultMarginLeft,
				Header: defaultMarginHeader,
				Footer: defaultMarginFooter,
				Gutter: defaultMarginGutter,
			},
			Cols: &Cols{
				Space: defaultColsSpace,
			},
		},
	}
}
