package files

import (
	"encoding/xml"
	"strconv"
)

const (
	xmlSchema = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>`
)

const (
	relationshipXmlns       = "http://schemas.openxmlformats.org/package/2006/relationships"
	relationShipWebSettings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
	relationShipSettings    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
	relationShipStyles      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
	relationShipThem1       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
	relationShipFontTable   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
	RelationShipImage       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)

const (
	targetWebSetting = "webSettings.xml"
	targetSettings   = "settings.xml"
	targetStyles     = "styles.xml"
	targetTheme1     = "theme/theme1.xml"
	targetFontTable  = "fontTable.xml"
	TargetMedia      = "media/"
)

const (
	ImageName = "image"
)

var (
	genID = 5
)

type IRelationFile interface {
	AddRelationship(string, string, string) *Relationship
}

type RelationshipFile struct {
	XmlSchema     string `xml:"-"`
	Relationships Relationships
}

type Relationships struct {
	XMLName       xml.Name `xml:"Relationships"`
	Xmlns         string   `xml:"xmlns,attr"`
	Relationships []*Relationship
}

type Relationship struct {
	XMLName xml.Name `xml:"Relationship"`
	ID      string   `xml:"Id,attr"`
	Type    string   `xml:"Type,attr"`
	Target  string   `xml:"Target,attr"`
}

func CreateRelationshipsFile() *RelationshipFile {
	return &RelationshipFile{
		XmlSchema: xmlSchema,
		Relationships: Relationships{
			Xmlns: relationshipXmlns,
			Relationships: []*Relationship{
				&Relationship{
					ID:     "rId1",
					Type:   relationShipWebSettings,
					Target: targetWebSetting,
				},
				&Relationship{
					ID:     "rId2",
					Type:   relationShipSettings,
					Target: targetSettings,
				},
				&Relationship{
					ID:     "rId3",
					Type:   relationShipStyles,
					Target: targetStyles,
				},
				&Relationship{
					ID:     "rId4",
					Type:   relationShipThem1,
					Target: targetTheme1,
				},
				&Relationship{
					ID:     "rId5",
					Type:   relationShipFontTable,
					Target: targetFontTable,
				},
			},
		},
	}
}

func (r *RelationshipFile) AddRelationship(id string, relType string, target string) *Relationship {
	relationShip := &Relationship{
		ID:     id,
		Type:   relType,
		Target: target,
	}
	r.Relationships.Relationships = append(r.Relationships.Relationships, relationShip)

	return relationShip
}

func GetRelationID() string {
	genID++

	return "rId" + strconv.Itoa(genID)
}
