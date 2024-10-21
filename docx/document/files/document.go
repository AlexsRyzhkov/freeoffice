package files

import (
	"encoding/xml"
	"github.com/AlexsRyzhkov/freeoffice/docx/document/entity"
	"github.com/AlexsRyzhkov/freeoffice/docx/document/enums"
	"github.com/AlexsRyzhkov/freeoffice/docx/document/fragments"
	"image"
	_ "image/jpeg"
	_ "image/png"
	"log"
	"os"
	"path/filepath"
)

const (
	wpc = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"

	cx  = "http://schemas.microsoft.com/office/drawing/2014/chartex"
	cx1 = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
	cx2 = "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
	cx3 = "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
	cx4 = "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
	cx5 = "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
	cx6 = "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
	cx7 = "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
	cx8 = "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"

	mc        = "http://schemas.openxmlformats.org/markup-compatibility/2006"
	aink      = "http://schemas.microsoft.com/office/drawing/2016/ink"
	am3d      = "http://schemas.microsoft.com/office/drawing/2017/model3d"
	o         = "urn:schemas-microsoft-com:office:office"
	oel       = "http://schemas.microsoft.com/office/2019/extlst"
	r         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	m         = "http://schemas.openxmlformats.org/officeDocument/2006/math"
	v         = "urn:schemas-microsoft-com:vml"
	wp14      = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
	wp        = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
	w10       = "urn:schemas-microsoft-com:office:word"
	w         = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	w14       = "http://schemas.microsoft.com/office/word/2010/wordml"
	w15       = "http://schemas.microsoft.com/office/word/2012/wordml"
	w16cex    = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
	w16cid    = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
	w16       = "http://schemas.microsoft.com/office/word/2018/wordml"
	w16DU     = "http://schemas.microsoft.com/office/word/2023/wordml/word16du"
	w16sdtdh  = "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
	w16se     = "http://schemas.microsoft.com/office/word/2015/wordml/symex"
	wpg       = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
	wpi       = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
	wne       = "http://schemas.microsoft.com/office/word/2006/wordml"
	wps       = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
	ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16du wp14"
)

const (
	XMLSchema = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
)

type IDocumentFile interface {
	AddParagraph(string, *fragments.TextProperty) fragments.ITextParagraph
	AddImage(string, *fragments.ImageProperty) fragments.IImageParagraph
	AddTable(int, int) fragments.ITable
}

type DocumentFile struct {
	XMLSchema  string          `xml:"-"`
	ImagesData []*entity.Image `xml:"-"`

	XMLName xml.Name `xml:"w:document"`
	Body    *fragments.FBody

	*DocumentSchemas
}

func (d *DocumentFile) AddParagraph(text string, property *fragments.TextProperty) fragments.ITextParagraph {
	textParagraph := fragments.CreateFTextParagraph(text, property)

	d.Body.AddParagraph(textParagraph)

	return textParagraph
}

func (d *DocumentFile) AddImage(url string, property *fragments.ImageProperty) fragments.IImageParagraph {
	file, err := os.Open(url)
	if err != nil {
		panic(err)
	}
	defer file.Close()

	img, _, err := image.DecodeConfig(file)
	if err != nil {
		log.Fatal(err)
	}

	cx := img.Width * entity.EmuPerPixel
	cy := img.Height * entity.EmuPerPixel

	rationHeightToWidth := float64(cy) / float64(cx)
	maxWidth := (enums.DefaultPageWidth - enums.DefaultMarginLeft - enums.DefaultMarginRight) * entity.TwipsToEmu

	if cx > maxWidth {
		cx = maxWidth
	}

	imageEntity := &entity.Image{
		NvPrID:              entity.GenNvPrID(),
		Cx:                  uint(cx),
		RationWidthToHeight: rationHeightToWidth,
		Embed:               GetRelationID(),
		Name:                entity.GenImageName(filepath.Ext(url)),
		Url:                 url,
	}

	d.ImagesData = append(d.ImagesData, imageEntity)

	imageParagraph := fragments.CreateFImageParagraph(imageEntity, property)

	d.Body.AddParagraph(imageParagraph)

	return imageParagraph
}

func (d *DocumentFile) AddTable(rows int, cols int) fragments.ITable {
	table := fragments.CreateFTable(rows, cols)

	d.Body.AddParagraph(table)

	return table
}

type DocumentSchemas struct {
	Wpc string `xml:"xmlns:wpc,attr"`

	Cx  string `xml:"xmlns:cx,attr"`
	Cx1 string `xml:"xmlns:cx1,attr"`
	Cx2 string `xml:"xmlns:cx2,attr"`
	Cx3 string `xml:"xmlns:cx3,attr"`
	Cx4 string `xml:"xmlns:cx4,attr"`
	Cx5 string `xml:"xmlns:cx5,attr"`
	Cx6 string `xml:"xmlns:cx6,attr"`
	Cx7 string `xml:"xmlns:cx7,attr"`
	Cx8 string `xml:"xmlns:cx8,attr"`

	Mc   string `xml:"xmlns:mc,attr"`
	Aink string `xml:"xmlns:aink,attr"`
	Am3d string `xml:"xmlns:am3d,attr"`
	O    string `xml:"xmlns:o,attr"`
	Oel  string `xml:"xmlns:oel,attr"`
	R    string `xml:"xmlns:r,attr"`
	M    string `xml:"xmlns:m,attr"`
	V    string `xml:"xmlns:v,attr"`
	Wp14 string `xml:"xmlns:wp14,attr"`

	Wp     string `xml:"xmlns:wp,attr"`
	W10    string `xml:"xmlns:w10,attr"`
	W      string `xml:"xmlns:w,attr"`
	W14    string `xml:"xmlns:w14,attr"`
	W15    string `xml:"xmlns:w15,attr"`
	W16cex string `xml:"xmlns:w16cex,attr"`
	W16cid string `xml:"xmlns:w16cid,attr"`
	W16    string `xml:"xmlns:w16,attr"`
	W16du  string `xml:"xmlns:w16du,attr"`

	W16sdtdh string `xml:"xmlns:w16sdtdh,attr"`
	W16se    string `xml:"xmlns:w16se,attr"`
	Wpg      string `xml:"xmlns:wpg,attr"`
	Wpi      string `xml:"xmlns:wpi,attr"`
	Wne      string `xml:"xmlns:wne,attr"`
	Wps      string `xml:"xmlns:wps,attr"`

	Ignorable string `xml:"mc:Ignorable,attr"`
}

func newDocumentSchema() *DocumentSchemas {
	return &DocumentSchemas{
		Wpc: wpc,
		Cx:  cx,
		Cx1: cx1,
		Cx2: cx2,
		Cx3: cx3,
		Cx4: cx4,
		Cx5: cx5,
		Cx6: cx6,
		Cx7: cx7,
		Cx8: cx8,

		Mc:   mc,
		Aink: aink,
		Am3d: am3d,
		O:    o,
		Oel:  oel,
		R:    r,
		M:    m,
		V:    v,
		Wp14: wp14,

		Wp:     wp,
		W10:    w10,
		W:      w,
		W14:    w14,
		W15:    w15,
		W16cex: w16cex,
		W16cid: w16cid,
		W16:    w16,
		W16du:  w16DU,

		W16sdtdh:  w16sdtdh,
		W16se:     w16se,
		Wpg:       wpg,
		Wpi:       wpi,
		Wne:       wne,
		Wps:       wps,
		Ignorable: ignorable,
	}
}

func CreateDocumentFile() *DocumentFile {
	return &DocumentFile{
		XMLSchema:       XMLSchema,
		DocumentSchemas: newDocumentSchema(),
		Body:            fragments.CreateFBody(),
	}
}
