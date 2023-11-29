package main

import (
	"fmt"
	"io"
	"log"
	"os"

	"github.com/fumiama/go-docx"
)

func DocAdd() {
	w := docx.NewA4()
	// add new paragraph
	para1 := w.AddParagraph()
	// add text
	para1.AddText("test").AddTab()
	para1.AddText("size").Size("44").AddTab()
	tbl2 := w.AddTableTwips([]int64{2333, 2333, 2333}, []int64{2333, 2333}).Justification("center")
	tbl2.TableRows[0].TableCells[0].Shade("clear", "auto", "E7E6E6")

	f, err := os.Create("generated.docx")
	// save to file
	if err != nil {
		panic(err)
	}
	_, err = w.WriteTo(f)
	if err != nil {
		panic(err)
	}
	err = f.Close()
	if err != nil {
		panic(err)
	}
	log.Println(w.Document.Body.Items)
}
func main() {
	readFile, err := os.OpenFile("测试表格.docx", os.O_RDWR, 0644)

	if err != nil {
		panic(err)
	}
	fileinfo, err := readFile.Stat()
	if err != nil {
		panic(err)
	}
	size := fileinfo.Size()
	doc, err := docx.Parse(readFile, size)
	if err != nil {
		log.Println(err)
	}
	defer readFile.Close()

	for _, it := range doc.Document.Body.Items {

		fmt.Println(it)

	}
	f, err := os.Create("generated.docx")
	newFile := docx.NewA4()

	para1 := newFile.AddParagraph()
	// add text
	para1.AddText("test").AddTab()

	para1.AddText("size").Size("44").AddTab()
	tbl2 := doc.AddTableTwips([]int64{2333, 2333, 2333}, []int64{2333, 2333}).Justification("end")
	tbl2.TableRows[0].TableCells[0].Shade("clear", "auto", "E7E6E6")

	tbl2.TableRows[0].TableCells[0].AddParagraph().AddText("test11").AddTab()
	tbl2.TableRows[0].TableCells[1].AddParagraph().AddText("test22").AddTab()
	tbl2.TableRows[1].TableCells[1].AddParagraph().AddText("test22").AddTab()
	tbl2.TableRows[2].TableCells[0].AddParagraph().Justification("center").AddText("test22").AddTab().Font("宋体", "宋体", "eastAsia").Bold()

	newFile.AppendFile(doc)

	_, err = io.Copy(f, newFile)
	if err != nil {
		log.Println(err)
	}
}
