package main

import (
	"github.com/AlexsRyzhkov/freeoffice/docx"
)

func main() {
	d := docx.New()

	doc := d.GetDocument()

	doc.AddParagraph("Text", nil)

	d.Save(".", "test34")
}
