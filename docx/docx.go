package docx

import (
	"archive/zip"
	"bytes"
	files2 "github.com/AlexsRyzhkov/freeoffice/docx/document/files"
	"github.com/AlexsRyzhkov/freeoffice/docx/helper"
	"github.com/AlexsRyzhkov/freeoffice/docx/templates"
	templrels "github.com/AlexsRyzhkov/freeoffice/docx/templates/_rels"
	templword "github.com/AlexsRyzhkov/freeoffice/docx/templates/word"
	tempwordthem "github.com/AlexsRyzhkov/freeoffice/docx/templates/word/theme"
	"io"
	"os"
	"path/filepath"
)

type Docx struct {
	relations *files2.RelationshipFile
	document  *files2.DocumentFile
}

func New() *Docx {
	return &Docx{
		relations: files2.CreateRelationshipsFile(),
		document:  files2.CreateDocumentFile(),
	}
}

func (d *Docx) GetDocument() files2.IDocumentFile {
	return d.document
}

func (d *Docx) GetRelations() files2.IRelationFile {
	return d.relations
}

func (d *Docx) GetBytes() bytes.Buffer {
	var buffer bytes.Buffer

	d.addToZip(&buffer)

	return buffer
}

func (d *Docx) Save(folder string, name string) {
	if name == "" {
		name = "gen"
	}

	zipDocx, err := os.Create(filepath.Join(folder, name+".docx"))
	if err != nil {
		panic(err)
	}

	defer zipDocx.Close()
	d.addToZip(zipDocx)
}

func (d *Docx) addToZip(zipDocx io.Writer) {
	zipDocxWriter := zip.NewWriter(zipDocx)
	defer zipDocxWriter.Close()

	copyContentToFile := [][]string{
		{templrels.Rels, "_rels/.rels"},
		{tempwordthem.Theme1, "word/theme/theme1.xml"},
		{templword.FontTable, "word/fontTable.xml"},
		{templword.Settings, "word/settings.xml"},
		{templword.Styles, "word/styles.xml"},
		{templword.WebSettings, "word/webSettings.xml"},
		{templates.ContentType, "[Content_Types].xml"},
	}

	for _, contentFile := range copyContentToFile {
		helper.WriteFileToZip(zipDocxWriter, contentFile[0], contentFile[1])
	}

	helper.WriteImageRelationToZip(zipDocxWriter, d.document, d.relations)
	helper.WriteDocumentToZip(zipDocxWriter, d.document)
	helper.WriteRelationWordToZip(zipDocxWriter, d.relations)
}
