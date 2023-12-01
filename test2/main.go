package main

import (
	"log"
	"os"

	"github.com/fumiama/go-docx"
)

func main() {
	readFile, err := os.OpenFile("测试表格.docx", os.O_RDWR|os.O_CREATE, 0755)

	if err != nil {
		log.Println(err)
	}
	fileinfo, err := readFile.Stat()
	if err != nil {
		log.Println(err)

	}
	size := fileinfo.Size()
	doc, err := docx.Parse(readFile, size)

	if err != nil {
		log.Println(err)
	}
	defer readFile.Close()

	for _, it := range doc.Document.Body.Items {

		switch o := it.(type) {

		case *docx.Paragraph: // printable

			o.Properties = nil
		case *docx.Table: // printable

			o.TableRows[0].TableCells[1].Paragraphs[0].AddText("ss").Shade("clear", "auto", "E7E6E6")

		}

	}
	// save to file
	if err != nil {
		panic(err)
	}

	_, err = doc.WriteTo(readFile)
	if err != nil {
		log.Println("1", err)
	}
	err = readFile.Close()
	if err != nil {
		log.Println("2", err)

	}
}
