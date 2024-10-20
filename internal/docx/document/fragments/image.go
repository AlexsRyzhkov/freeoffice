package fragments

import (
	"encoding/xml"
	"github.com/freeoffice/internal/docx/document/entity"
	"math/rand/v2"
	"strconv"
)

const (
	emuPerPixel = 9525
)

const (
	xMLNSA   = "http://schemas.openxmlformats.org/drawingml/2006/main"
	xMLNSPIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
	uri      = "http://schemas.openxmlformats.org/drawingml/2006/picture"
)

const (
	effectExtentDefaultL = "0"
	effectExtentDefaultT = "0"
	effectExtentDefaultR = "5080"
	effectExtentDefaultB = "0"
)

const (
	noChangeAspectDefault     = "1"
	noChangeArrowheadsDefault = "1"
)

const (
	bwModeDefault = "auto"
	offXDefault   = "0"
	offYDefault   = "0"
	prst          = "rect"
)

type IImageParagraph interface {
	SetSizeByWidth(int)
	SetSizeByHeight(int)

	SetWidth(int)
	SetHeight(int)

	SetJustify(string)
	SetLeftOffSet(string)
	SetRightOffSet(string)
}

type FImageParagraph struct {
	XMLName  xml.Name `xml:"w:p"`
	FImage   *FImage  `xml:"w:r"`
	Property *FImageParagraphProperty
}

func (fip *FImageParagraph) SetSizeByWidth(width int) {
	if width > 0 {
		oldWidth := fip.FImage.Drawing.Inline.Extent.Cx
		oldHeight := fip.FImage.Drawing.Inline.Extent.Cy

		heightToWidthRation := float64(oldHeight) / float64(oldWidth)
		heightByRation := heightToWidthRation * float64(width)

		fip.FImage.Drawing.Inline.Extent.Cx = width * emuPerPixel
		fip.FImage.Drawing.Inline.Graphic.GraphicData.Pic.SpPr.Xfrm.Ext.CX = width * emuPerPixel

		fip.FImage.Drawing.Inline.Extent.Cy = int(heightByRation)
		fip.FImage.Drawing.Inline.Graphic.GraphicData.Pic.SpPr.Xfrm.Ext.CY = int(heightByRation)
	}
}

func (fip *FImageParagraph) SetSizeByHeight(height int) {
	if height > 0 {
		oldWidth := fip.FImage.Drawing.Inline.Extent.Cx
		oldHeight := fip.FImage.Drawing.Inline.Extent.Cy

		widthToHeightRation := float64(oldWidth) / float64(oldHeight)
		widthByRation := widthToHeightRation * float64(height)

		fip.FImage.Drawing.Inline.Extent.Cy = height * emuPerPixel
		fip.FImage.Drawing.Inline.Graphic.GraphicData.Pic.SpPr.Xfrm.Ext.CY = height * emuPerPixel

		fip.FImage.Drawing.Inline.Extent.Cx = int(widthByRation)
		fip.FImage.Drawing.Inline.Graphic.GraphicData.Pic.SpPr.Xfrm.Ext.CX = int(widthByRation)

	}
}

func (fip *FImageParagraph) SetWidth(width int) {
	if width > 0 {
		fip.FImage.Drawing.Inline.Extent.Cx = width * emuPerPixel
		fip.FImage.Drawing.Inline.Graphic.GraphicData.Pic.SpPr.Xfrm.Ext.CX = width * emuPerPixel
	}
}

func (fip *FImageParagraph) SetHeight(height int) {
	if height > 0 {
		fip.FImage.Drawing.Inline.Extent.Cx = height * emuPerPixel
		fip.FImage.Drawing.Inline.Graphic.GraphicData.Pic.SpPr.Xfrm.Ext.CX = height * emuPerPixel
	}
}

func (fip *FImageParagraph) SetJustify(justify string) {
	if justify == "" {
		fip.Property.Justify = nil
	} else {
		fip.Property.Justify = &FImageParagraphJustify{Val: justify}
	}
}

func (fip *FImageParagraph) SetLeftOffSet(leftOffSet string) {
	if leftOffSet == "" && fip.Property.OffSet != nil {
		fip.Property.OffSet.LeftOffSet = nil
	}

	if leftOffSet != "" {
		if fip.Property.OffSet != nil {
			fip.Property.OffSet.LeftOffSet = &leftOffSet
		} else {
			fip.Property.OffSet = &FImageParagraphOffSet{LeftOffSet: &leftOffSet}
		}
	}
}

func (fip *FImageParagraph) SetRightOffSet(rightOffSet string) {
	if rightOffSet == "" && fip.Property.OffSet != nil {
		fip.Property.OffSet.RightOffSet = nil
	}

	if rightOffSet != "" {
		if fip.Property.OffSet != nil {
			fip.Property.OffSet.RightOffSet = &rightOffSet
		} else {
			fip.Property.OffSet = &FImageParagraphOffSet{RightOffSet: &rightOffSet}
		}
	}
}

type FImageParagraphProperty struct {
	XMLName xml.Name `xml:"w:pPr"`
	Justify *FImageParagraphJustify
	OffSet  *FImageParagraphOffSet
}

type FImageParagraphJustify struct {
	XMLName xml.Name `xml:"w:js"`
	Val     string   `xml:"w:val,attr"`
}

type FImageParagraphOffSet struct {
	XMLName     xml.Name `xml:"w:ind,omitempty"`
	LeftOffSet  *string  `xml:"w:firstLine,attr,omitempty"`
	RightOffSet *string  `xml:"w:right,attr,omitempty"`
}

type FImage struct {
	Drawing Drawing
}

type Drawing struct {
	XMLName xml.Name `xml:"w:drawing"`
	Inline  Inline
}

type Inline struct {
	XMLName           xml.Name `xml:"wp:inline"`
	Extent            Extent
	EffectExtent      EffectExtent
	DocPr             DocPr
	CNvGraphicFramePr CNvGraphicFramePr
	Graphic           Graphic
}

type Extent struct {
	XMLName xml.Name `xml:"wp:extent"`
	Cx      int      `xml:"cx,attr"`
	Cy      int      `xml:"cy,attr"`
}

type EffectExtent struct {
	XMLName xml.Name `xml:"wp:effectExtent"`
	L       string   `xml:"l,attr"`
	T       string   `xml:"t,attr"`
	R       string   `xml:"r,attr"`
	B       string   `xml:"b,attr"`
}

type DocPr struct {
	XMLName xml.Name `xml:"wp:docPr"`
	ID      string   `xml:"id,attr"`
	Name    string   `xml:"name,attr"`
}

type CNvGraphicFramePr struct {
	XMLName           xml.Name `xml:"wp:cNvGraphicFramePr"`
	GraphicFrameLocks GraphicFrameLocks
}

type GraphicFrameLocks struct {
	XMLName        xml.Name `xml:"a:graphicFrameLocks"`
	XmlnsA         string   `xml:"xmlns:a,attr"`
	NoChangeAspect string   `xml:"noChangeAspect,attr"`
}

type Graphic struct {
	XMLName     xml.Name `xml:"a:graphic"`
	XmlnsA      string   `xml:"xmlns:a,attr"`
	GraphicData GraphicData
}

type GraphicData struct {
	XMLName xml.Name `xml:"a:graphicData"`
	Uri     string   `xml:"uri,attr"`
	Pic     Pic
}

type Pic struct {
	XMLName  xml.Name `xml:"pic:pic"`
	XmlnsPic string   `xml:"xmlns:pic,attr"`
	NvPicPr  NvPicPr
	BlipFill BlipFill
	SpPr     SpPr
}

type NvPicPr struct {
	XMLName  xml.Name `xml:"pic:nvPicPr"`
	CNvPr    CNvPr
	CNvPicPr CNvPicPr
}

type CNvPr struct {
	XMLName xml.Name `xml:"pic:cNvPr"`
	ID      string   `xml:"id,attr"`
	Name    string   `xml:"name,attr"`
}

type CNvPicPr struct {
	XMLName   xml.Name `xml:"pic:cNvPicPr"`
	PickLocks PickLocks
}

type PickLocks struct {
	XMLName            xml.Name `xml:"a:picLocks"`
	NoChangeAspect     string   `xml:"noChangeAspect,attr"`
	NoChangeArrowheads string   `xml:"noChangeArrowheads,attr"`
}

type BlipFill struct {
	XMLName         xml.Name `xml:"pic:blipFill"`
	Blip            Blip
	SrcRect         string `xml:"a:srcRect"`
	StretchFillRect string `xml:"a:stretch>a:fillRect"`
}

type Blip struct {
	XMLName xml.Name `xml:"a:blip"`
	Embed   string   `xml:"r:embed,attr"`
	ExtLst  string   `xml:"a:extLst"`
}

type SpPr struct {
	XMLName  xml.Name `xml:"pic:spPr"`
	BwMode   string   `xml:"bwMode,attr"`
	Xfrm     Xfrm
	PrstGeom PrstGeom
}

type Xfrm struct {
	XMLName xml.Name `xml:"a:xfrm"`
	Off     Off
	Ext     Ext
}

type Off struct {
	XMLName xml.Name `xml:"a:off"`
	X       string   `xml:"x,attr"`
	Y       string   `xml:"y,attr"`
}

type Ext struct {
	XMLName xml.Name `xml:"a:ext"`
	CX      int      `xml:"cx,attr"`
	CY      int      `xml:"cy,attr"`
}

type PrstGeom struct {
	XMLName xml.Name `xml:"a:prstGeom"`
	Prst    string   `xml:"prst,attr"`
	AvLst   string   `xml:"a:avLst"`
}

type ImageProperty struct {
	Width int

	Justify     string
	LeftOffSet  string
	RightOffSet string
}

func CreateFImageParagraph(image *entity.Image, prop *ImageProperty) IImageParagraph {
	fImageParagraphProperty := &FImageParagraphProperty{}

	if prop == nil {
		prop = &ImageProperty{}
	}

	if prop.Width > 0 {
		image.Cx = prop.Width * emuPerPixel
	}

	if prop.Justify != "" {
		fImageParagraphProperty.Justify = &FImageParagraphJustify{Val: prop.Justify}
	}

	if prop.LeftOffSet != "" {
		fImageParagraphProperty.OffSet = &FImageParagraphOffSet{LeftOffSet: &prop.LeftOffSet}
	}

	if prop.RightOffSet != "" {
		if fImageParagraphProperty.OffSet == nil {
			fImageParagraphProperty.OffSet = &FImageParagraphOffSet{RightOffSet: &prop.RightOffSet}
		} else {
			fImageParagraphProperty.OffSet.RightOffSet = &prop.RightOffSet
		}
	}

	return &FImageParagraph{
		Property: fImageParagraphProperty,
		FImage:   createFImage(image),
	}
}

func createFImage(image *entity.Image) *FImage {
	return &FImage{Drawing: Drawing{
		Inline: Inline{
			Extent: Extent{
				Cx: image.Cx,
				Cy: int(float64(image.Cx) * image.RationWidthToHeight),
			},
			EffectExtent: EffectExtent{
				L: effectExtentDefaultL,
				T: effectExtentDefaultT,
				R: effectExtentDefaultR,
				B: effectExtentDefaultB,
			},
			DocPr: DocPr{
				ID:   strconv.Itoa(rand.IntN(1000000000)),
				Name: "Рисунок " + image.NvPrID,
			},
			CNvGraphicFramePr: CNvGraphicFramePr{
				GraphicFrameLocks: GraphicFrameLocks{
					XmlnsA:         xMLNSA,
					NoChangeAspect: noChangeAspectDefault,
				},
			},
			Graphic: Graphic{
				XmlnsA: xMLNSA,
				GraphicData: GraphicData{
					Uri: uri, Pic: Pic{
						XmlnsPic: xMLNSPIC,
						NvPicPr: NvPicPr{
							CNvPr: CNvPr{
								ID:   image.NvPrID,
								Name: "Picture " + image.NvPrID,
							},
							CNvPicPr: CNvPicPr{
								PickLocks: PickLocks{
									NoChangeAspect:     noChangeAspectDefault,
									NoChangeArrowheads: noChangeArrowheadsDefault,
								},
							},
						},
						BlipFill: BlipFill{
							Blip: Blip{
								Embed:  image.Embed,
								ExtLst: "",
							},
							SrcRect:         "",
							StretchFillRect: "",
						},
						SpPr: SpPr{
							BwMode: bwModeDefault,
							Xfrm: Xfrm{
								Off: Off{
									X: offXDefault,
									Y: offYDefault,
								},
								Ext: Ext{
									CX: image.Cx,
									CY: int(float64(image.Cx) * image.RationWidthToHeight),
								},
							},
							PrstGeom: PrstGeom{
								Prst:  prst,
								AvLst: "",
							},
						},
					},
				},
			},
		},
	}}
}
