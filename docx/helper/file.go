package helper

import (
	"archive/zip"
	"encoding/xml"
	files2 "github.com/AlexsRyzhkov/freeoffice/docx/document/files"
	"io"
	"os"
)

func WriteFileToZip(zipDocxWriter *zip.Writer, content string, to string) {
	zipFile, _ := zipDocxWriter.Create(to)
	zipFile.Write([]byte(content))
}

func WriteDocumentToZip(zipDocxWriter *zip.Writer, doc *files2.DocumentFile) {
	zipWordDocumentFile, _ := zipDocxWriter.Create("word/document.xml")

	docData, err := xml.MarshalIndent(doc, "", "  ")
	if err != nil {
		panic(err)
	}

	zipWordDocumentFile.Write([]byte(doc.XMLSchema))
	zipWordDocumentFile.Write(docData)
}

func WriteRelationWordToZip(zipDocxWriter *zip.Writer, rel *files2.RelationshipFile) {
	zipWordDocumentFile, _ := zipDocxWriter.Create("word/_rels/document.xml.rels")

	docData, err := xml.MarshalIndent(rel.Relationships, "", "  ")
	if err != nil {
		panic(err)
	}

	zipWordDocumentFile.Write([]byte(files2.XMLSchema))
	zipWordDocumentFile.Write(docData)
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
