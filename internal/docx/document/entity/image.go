package entity

import (
	"strconv"
)

const (
	EmuPerPixel = 9525
	TwipsToEmu  = 635
)

type Image struct {
	Cx                  int
	RationWidthToHeight float64
	NvPrID              string
	Embed               string
	Name                string
	Url                 string
}

var (
	nvPrID  = -1
	imageID = -1
)

func GenNvPrID() string {
	nvPrID++
	return strconv.Itoa(nvPrID)
}

func GenImageName(imageType string) string {
	imageID++
	return "image" + strconv.Itoa(imageID) + imageType
}
