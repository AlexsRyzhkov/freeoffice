package docx

import (
	"archive/zip"
	"bytes"
	"github.com/AlexsRyzhkov/freeoffice/docx/document/files"
	"github.com/AlexsRyzhkov/freeoffice/docx/helper"
	"github.com/AlexsRyzhkov/freeoffice/docx/templates"
	templrels "github.com/AlexsRyzhkov/freeoffice/docx/templates/_rels"
	templword "github.com/AlexsRyzhkov/freeoffice/docx/templates/word"
	tempwordthem "github.com/AlexsRyzhkov/freeoffice/docx/templates/word/theme"
	"io"
	"os"
	"path/filepath"
)

type IDocx interface {
	GetDocument() files.IDocumentFile
	Save(string, string) error
	GetBytes() (*bytes.Buffer, error)
}

type Docx struct {
	relations *files.RelationshipFile
	document  *files.DocumentFile
}

func New() IDocx {
	return &Docx{
		relations: files.CreateRelationshipsFile(),
		document:  files.CreateDocumentFile(),
	}
}

func (d *Docx) GetDocument() files.IDocumentFile {
	return d.document
}

func (d *Docx) GetRelations() files.IRelationFile {
	return d.relations
}

func (d *Docx) GetBytes() (*bytes.Buffer, error) {
	var buffer bytes.Buffer

	err := d.addToZip(&buffer)
	if err != nil {
		return nil, err
	}

	return &buffer, nil
}

func (d *Docx) Save(folder string, name string) error {
	if name == "" {
		name = "gen"
	}

	zipDocx, err := os.Create(filepath.Join(folder, name+".docx"))
	if err != nil {
		return err
	}

	defer zipDocx.Close()
	return d.addToZip(zipDocx)
}

func (d *Docx) addToZip(zipDocx io.Writer) error {
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
		err := helper.WriteFileToZip(zipDocxWriter, contentFile[0], contentFile[1])
		if err != nil {
			return err
		}
	}

	err := helper.WriteImageRelationToZip(zipDocxWriter, d.document, d.relations)
	if err != nil {
		return err
	}

	err = helper.WriteDocumentToZip(zipDocxWriter, d.document)
	if err != nil {
		return err
	}

	err = helper.WriteRelationWordToZip(zipDocxWriter, d.relations)
	if err != nil {
		return err
	}

	return nil
}
