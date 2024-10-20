package docx

import (
	"archive/zip"
	"bytes"
	files2 "github.com/AlexsRyzhkov/freeoffice/docx/document/files"
	"github.com/AlexsRyzhkov/freeoffice/docx/helper"
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

	copyFiles := [][]string{
		[]string{"internal/docx/default-struct/_rels/.rels", "_rels/.rels"},
		[]string{"internal/docx/default-struct/word/theme/theme1.xml", "word/theme/theme1.xml"},
		[]string{"internal/docx/default-struct/word/fontTable.xml", "word/fontTable.xml"},
		[]string{"internal/docx/default-struct/word/settings.xml", "word/settings.xml"},
		[]string{"internal/docx/default-struct/word/styles.xml", "word/styles.xml"},
		[]string{"internal/docx/default-struct/word/webSettings.xml", "word/webSettings.xml"},
		[]string{"internal/docx/default-struct/[Content_Types].xml", "[Content_Types].xml"},
	}

	for _, files := range copyFiles {
		helper.WriteFileToZip(zipDocxWriter, files[0], files[1])
	}

	helper.WriteImageRelationToZip(zipDocxWriter, d.document, d.relations)
	helper.WriteDocumentToZip(zipDocxWriter, d.document)
	helper.WriteRelationWordToZip(zipDocxWriter, d.relations)
}
