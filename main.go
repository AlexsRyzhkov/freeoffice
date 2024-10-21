package main

import (
	"github.com/AlexsRyzhkov/freeoffice/docx"
	"github.com/AlexsRyzhkov/freeoffice/docx/document/enums"
)

func main() {
	d := docx.New()

	doc := d.GetDocument()

	doc.AddParagraph("Text", nil).SetLineSpace(enums.LineSpace200)

	d.Save(".", "test34")
}
