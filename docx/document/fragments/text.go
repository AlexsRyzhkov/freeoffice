package fragments

import (
	"encoding/xml"
	"github.com/AlexsRyzhkov/freeoffice/docx/document/enums"
)

var (
	emptyString = ""
)

var (
	lineRule       = "auto"
	spaceBefore    = 240
	spaceAfterZero = 0
)

type ITextParagraph interface {
	SetText(string) ITextParagraph

	SetBold() ITextParagraph
	SetItalic() ITextParagraph
	SetUnderline() ITextParagraph
	SetStrike() ITextParagraph

	UnSetBold() ITextParagraph
	UnSetItalic() ITextParagraph
	UnSetUnderline() ITextParagraph
	UnSetStrike() ITextParagraph

	SetFontFamily(string) ITextParagraph
	SetFontSize(uint) ITextParagraph

	SetTextColor(string) ITextParagraph
	SetTextHighlightColor(string) ITextParagraph

	SetJustify(string) ITextParagraph
	SetLeftOffSet(uint) ITextParagraph
	SetRightOffSet(uint) ITextParagraph

	SetSpaceBefore() ITextParagraph
	SetSpaceAfter() ITextParagraph
	SetLineSpace(uint) ITextParagraph

	UnSetSpaceBefore() ITextParagraph
	UnSetSpaceAfter() ITextParagraph
}

type FTextParagraph struct {
	XMLName  xml.Name `xml:"w:p"`
	Property *FTextParagraphProperty
	FText    *FText `xml:"w:r"`
}

func (ftp *FTextParagraph) SetText(text string) ITextParagraph {
	ftp.FText.Text.Val = text
	return ftp
}

func (ftp *FTextParagraph) SetBold() ITextParagraph {
	ftp.FText.Property.Bold = &emptyString
	return ftp
}

func (ftp *FTextParagraph) SetItalic() ITextParagraph {
	ftp.FText.Property.Italic = &emptyString
	return ftp
}

func (ftp *FTextParagraph) SetUnderline() ITextParagraph {
	ftp.FText.Property.Underline = &emptyString
	return ftp
}

func (ftp *FTextParagraph) SetStrike() ITextParagraph {
	ftp.FText.Property.Strike = &emptyString
	return ftp
}

func (ftp *FTextParagraph) UnSetBold() ITextParagraph {
	ftp.FText.Property.Bold = nil
	return ftp
}

func (ftp *FTextParagraph) UnSetItalic() ITextParagraph {
	ftp.FText.Property.Italic = nil
	return ftp
}

func (ftp *FTextParagraph) UnSetUnderline() ITextParagraph {
	ftp.FText.Property.Underline = nil
	return ftp
}

func (ftp *FTextParagraph) UnSetStrike() ITextParagraph {
	ftp.FText.Property.Strike = nil
	return ftp
}

func (ftp *FTextParagraph) SetFontFamily(fontFamily string) ITextParagraph {
	if fontFamily == "" {
		ftp.FText.Property.Font = &FFonts{Ascii: enums.TimesNewRoman, HAnsi: enums.TimesNewRoman}
	} else {
		ftp.FText.Property.Font = &FFonts{
			Ascii: fontFamily,
			HAnsi: fontFamily,
		}
	}

	return ftp
}

func (ftp *FTextParagraph) SetFontSize(fontSize uint) ITextParagraph {
	if fontSize == 0 {
		ftp.FText.Property.FTextSize = nil
		ftp.FText.Property.FTextSizeComplex = nil
	} else {
		ftp.FText.Property.FTextSize = &FTextSize{Val: fontSize * 2}
		ftp.FText.Property.FTextSizeComplex = &FTextSizeComplex{Val: fontSize * 2}
	}

	return ftp
}

func (ftp *FTextParagraph) SetTextColor(color string) ITextParagraph {
	if color == "" {
		ftp.FText.Property.TextColor = nil
	} else {
		ftp.FText.Property.TextColor = &FTextColor{Val: color}
	}

	return ftp
}

func (ftp *FTextParagraph) SetTextHighlightColor(color string) ITextParagraph {
	if color == "" {
		ftp.FText.Property.Highlight = nil
	} else {
		ftp.FText.Property.Highlight = &FHighlight{Val: color}
	}

	return ftp
}

func (ftp *FTextParagraph) SetJustify(justify string) ITextParagraph {
	if justify == "" {
		ftp.Property.Justify = nil
	} else {
		ftp.Property.Justify = &FTextParagraphJustify{Val: justify}
	}

	return ftp
}

func (ftp *FTextParagraph) SetLeftOffSet(leftOffSet uint) ITextParagraph {
	if ftp.Property.OffSet != nil {
		ftp.Property.OffSet.LeftOffSet = &leftOffSet
	} else {
		ftp.Property.OffSet = &FTextParagraphOffSet{LeftOffSet: &leftOffSet}
	}

	return ftp
}

func (ftp *FTextParagraph) SetRightOffSet(rightOffSet uint) ITextParagraph {
	if ftp.Property.OffSet != nil {
		ftp.Property.OffSet.RightOffSet = &rightOffSet
	} else {
		ftp.Property.OffSet = &FTextParagraphOffSet{RightOffSet: &rightOffSet}
	}

	return ftp
}

func (ftp *FTextParagraph) SetSpaceBefore() ITextParagraph {
	if ftp.Property.Space == nil {
		ftp.Property.Space = &FTextParagraphSpacing{
			Before:   &spaceBefore,
			LineRule: lineRule,
		}
	}

	ftp.Property.Space.Before = &spaceBefore

	return ftp
}

func (ftp *FTextParagraph) SetSpaceAfter() ITextParagraph {
	if ftp.Property.Space == nil {
		ftp.Property.Space = &FTextParagraphSpacing{
			After:    nil,
			LineRule: lineRule,
		}
	}

	ftp.Property.Space.After = nil

	return ftp
}

func (ftp *FTextParagraph) SetLineSpace(lineSpace uint) ITextParagraph {
	if ftp.Property.Space == nil {
		ftp.Property.Space = &FTextParagraphSpacing{
			Line:     &lineSpace,
			LineRule: lineRule,
		}
	}

	ftp.Property.Space.Line = &lineSpace

	return ftp
}

func (ftp *FTextParagraph) UnSetSpaceBefore() ITextParagraph {
	if ftp.Property.Space != nil {
		ftp.Property.Space.Before = nil
	}

	return ftp
}

func (ftp *FTextParagraph) UnSetSpaceAfter() ITextParagraph {
	if ftp.Property.Space == nil {
		ftp.Property.Space = &FTextParagraphSpacing{
			After:    &spaceAfterZero,
			LineRule: lineRule,
		}
	}

	ftp.Property.Space.After = &spaceAfterZero

	return ftp
}

type FTextParagraphProperty struct {
	XMLName xml.Name `xml:"w:pPr"`
	Space   *FTextParagraphSpacing
	Justify *FTextParagraphJustify
	OffSet  *FTextParagraphOffSet
}

type FTextParagraphSpacing struct {
	XMLName  xml.Name `xml:"w:spacing"`
	After    *int     `xml:"w:after,attr"`
	Before   *int     `xml:"w:before,attr"`
	Line     *uint    `xml:"w:line,attr"`
	LineRule string   `xml:"w:lineRule,attr"`
}

type FTextParagraphJustify struct {
	XMLName xml.Name `xml:"w:js"`
	Val     string   `xml:"w:val,attr"`
}

type FTextParagraphOffSet struct {
	XMLName     xml.Name `xml:"w:ind,omitempty"`
	LeftOffSet  *uint    `xml:"w:firstLine,attr,omitempty"`
	RightOffSet *uint    `xml:"w:right,attr,omitempty"`
}

type FText struct {
	Property *FTextProperty
	Text     *Text
}

type FTextProperty struct {
	XMLName          xml.Name `xml:"w:rPr,omitempty"`
	Bold             *string  `xml:"w:b,omitempty"`
	Italic           *string  `xml:"w:i,omitempty"`
	Underline        *string  `xml:"w:underline,omitempty"`
	Strike           *string  `xml:"w:strike,omitempty"`
	Font             *FFonts
	TextColor        *FTextColor
	Highlight        *FHighlight
	FTextSize        *FTextSize
	FTextSizeComplex *FTextSizeComplex
}

type FFonts struct {
	XMLName xml.Name `xml:"w:rFonts,omitempty"`
	Ascii   string   `xml:"w:ascii,attr"`
	HAnsi   string   `xml:"w:hAnsi,attr"`
}

type FTextColor struct {
	XMLName xml.Name `xml:"w:color,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

type FTextSize struct {
	XMLName xml.Name `xml:"w:sz,omitempty"`
	Val     uint     `xml:"w:val,attr"`
}

type FTextSizeComplex struct {
	XMLName xml.Name `xml:"w:szCs,omitempty"`
	Val     uint     `xml:"w:val,attr"`
}

type FHighlight struct {
	XMLName xml.Name `xml:"w:highlight,omitempty"`
	Val     string   `xml:"w:val,attr"`
}

type Text struct {
	XMLName xml.Name `xml:"w:t"`
	Val     string   `xml:",chardata"`
}

type TextProperty struct {
	Bold      bool
	Italic    bool
	Underline bool
	Strike    bool

	FontFamily string
	FontSize   uint

	TextColor          string
	TextHighlightColor string

	Justify     string
	LeftOffSet  uint
	RightOffSet uint
}

func CreateFTextParagraph(text string, prop *TextProperty) ITextParagraph {
	fTextProperty := &FTextProperty{}
	fTextParagraphProperty := &FTextParagraphProperty{}

	if prop == nil {
		prop = &TextProperty{}
	}

	if prop.Bold {
		fTextProperty.Bold = &emptyString
	}

	if prop.Italic {
		fTextProperty.Italic = &emptyString
	}

	if prop.Underline {
		fTextProperty.Underline = &emptyString
	}

	if prop.Strike {
		fTextProperty.Strike = &emptyString
	}

	if prop.FontFamily != "" {
		fTextProperty.Font = &FFonts{
			Ascii: prop.FontFamily,
			HAnsi: prop.FontFamily,
		}
	} else {
		fTextProperty.Font = &FFonts{
			Ascii: enums.TimesNewRoman,
			HAnsi: enums.TimesNewRoman,
		}
	}

	if prop.FontSize != 0 {
		fTextProperty.FTextSize = &FTextSize{Val: prop.FontSize}
		fTextProperty.FTextSizeComplex = &FTextSizeComplex{Val: prop.FontSize}
	}

	if prop.TextColor != "" {
		fTextProperty.TextColor = &FTextColor{Val: prop.TextColor}
	}

	if prop.TextHighlightColor != "" {
		fTextProperty.Highlight = &FHighlight{Val: prop.TextHighlightColor}
	}

	if prop.Justify != "" {
		fTextParagraphProperty.Justify = &FTextParagraphJustify{Val: prop.Justify}
	}

	if prop.LeftOffSet > 0 {
		fTextParagraphProperty.OffSet = &FTextParagraphOffSet{LeftOffSet: &prop.LeftOffSet}
	}

	if prop.RightOffSet > 0 {
		if fTextParagraphProperty.OffSet != nil {
			fTextParagraphProperty.OffSet.RightOffSet = &prop.RightOffSet
		} else {
			fTextParagraphProperty.OffSet = &FTextParagraphOffSet{RightOffSet: &prop.RightOffSet}
		}
	}

	return &FTextParagraph{
		FText: &FText{
			Text:     &Text{Val: text},
			Property: fTextProperty,
		},
		Property: fTextParagraphProperty,
	}
}
