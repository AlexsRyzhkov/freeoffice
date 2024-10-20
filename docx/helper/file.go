package helper

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	files2 "github.com/AlexsRyzhkov/freeoffice/docx/document/files"
	"io"
	"os"
)

func WriteFileToZip(zipDocxWriter *zip.Writer, from string, to string) {
	zipFile, _ := zipDocxWriter.Create(to)

	relsFile, err := os.Open(from)
	if err != nil {
		panic(err)
	}
	defer relsFile.Close()

	io.Copy(zipFile, relsFile)
}

func WriteDocumentToZip(zipDocxWriter *zip.Writer, doc *files2.DocumentFile) {
	zipWordDocumentFile, _ := zipDocxWriter.Create("word/document.xml")

	var buf bytes.Buffer

	docData, err := xml.MarshalIndent(doc, "", "  ")
	if err != nil {
		panic(err)
	}

	buf.WriteString(doc.XMLSchema)
	buf.Write(docData)

	_, err = buf.WriteTo(zipWordDocumentFile)
	if err != nil {
		return
	}
}

func WriteRelationWordToZip(zipDocxWriter *zip.Writer, rel *files2.RelationshipFile) {
	zipWordDocumentFile, _ := zipDocxWriter.Create("word/_rels/document.xml.rels")

	var buf bytes.Buffer

	docData, err := xml.MarshalIndent(rel.Relationships, "", "  ")
	if err != nil {
		panic(err)
	}

	buf.WriteString(files2.XMLSchema)
	buf.Write(docData)

	_, err = buf.WriteTo(zipWordDocumentFile)
	if err != nil {
		return
	}
}

func WriteImageRelationToZip(zipDocxWriter *zip.Writer, d *files2.DocumentFile, rel *files2.RelationshipFile) {
	images := d.ImagesData

	for _, image := range images {
		rel.AddRelationship(
			image.Embed,
			files2.RelationShipImage,
			files2.TargetMedia+image.Name,
		)

		zipImageWriter, _ := zipDocxWriter.Create("word/media/" + image.Name)

		relsFile, err := os.Open(image.Url)
		if err != nil {
			panic(err)
		}

		io.Copy(zipImageWriter, relsFile)

		relsFile.Close()
	}
}
