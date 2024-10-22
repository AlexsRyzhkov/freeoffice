package helper

import (
	"archive/zip"
	"encoding/xml"
	files2 "github.com/AlexsRyzhkov/freeoffice/docx/document/files"
	"io"
	"os"
)

func WriteFileToZip(zipDocxWriter *zip.Writer, content string, to string) error {
	zipFile, err := zipDocxWriter.Create(to)
	if err != nil {
		return err
	}

	_, err = zipFile.Write([]byte(content))
	if err != nil {
		return err
	}

	return nil
}

func WriteDocumentToZip(zipDocxWriter *zip.Writer, doc *files2.DocumentFile) error {
	zipWordDocumentFile, err := zipDocxWriter.Create("word/document.xml")
	if err != nil {
		return err
	}

	docData, err := xml.MarshalIndent(doc, "", "  ")
	if err != nil {
		return err
	}

	_, err = zipWordDocumentFile.Write([]byte(doc.XMLSchema))
	if err != nil {
		return err
	}
	_, err = zipWordDocumentFile.Write(docData)
	if err != nil {
		return err
	}

	return nil
}

func WriteRelationWordToZip(zipDocxWriter *zip.Writer, rel *files2.RelationshipFile) error {
	zipWordDocumentFile, err := zipDocxWriter.Create("word/_rels/document.xml.rels")
	if err != nil {
		return err
	}

	docData, err := xml.MarshalIndent(rel.Relationships, "", "  ")
	if err != nil {
		return err
	}

	_, err = zipWordDocumentFile.Write([]byte(files2.XMLSchema))
	if err != nil {
		return err
	}

	_, err = zipWordDocumentFile.Write(docData)
	if err != nil {
		return err
	}

	return nil
}

func WriteImageRelationToZip(zipDocxWriter *zip.Writer, d *files2.DocumentFile, rel *files2.RelationshipFile) error {
	images := d.ImagesData

	for _, image := range images {
		rel.AddRelationship(
			image.Embed,
			files2.RelationShipImage,
			files2.TargetMedia+image.Name,
		)

		zipImageWriter, err := zipDocxWriter.Create("word/media/" + image.Name)
		if err != nil {
			return err
		}

		relsFile, err := os.Open(image.Url)
		if err != nil {
			return err
		}

		_, err = io.Copy(zipImageWriter, relsFile)
		if err != nil {
			return err
		}

		err = relsFile.Close()
		if err != nil {
			return err
		}
	}

	return nil
}
