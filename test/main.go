package main

import (
	"io"
	"log"
	"os"

	docx "github.com/fumiama/go-docx"
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

			// fmt.Println("==Paragraph", o)
			// doc.Document.Body.Items[i] = nil
			// o.Properties = nil

			// fmt.Println("--", o)
			c := make([]interface{}, 0, 64)

			c = append(c, &docx.Text{
				Text: "test==\t",
			})

			run := &docx.Run{
				RunProperties: &docx.RunProperties{},
				Children:      c,
			}

			for i, _ := range o.Children {

				o.Children[i] = run

			}

			// o.Justification("center")

			// doc.Document.Body.Items[i] = o
			// o.Properties = nil
		case *docx.Table: // printable

			o.TableRows[0].TableCells[1].Paragraphs[0].AddText("ss").Shade("clear", "auto", "E7E6E6")

			// for _, tr := range o.TableRows {
			// 	for _, tc := range tr.TableCells {

			// 		for _, p := range tc.Paragraphs {

			// 			if p.String() == "活动名称" {
			// 				fmt.Println("==add", o.TableRows[1].TableCells[1])
			// 				p.AddText("add")
			// 			}
			// 			p.Properties = nil
			// 		}
			// 	}
			// }
		}

	}

	f, err := os.Create("generated.docx")
	newFile := docx.NewA4()

	// para1 := newFile.AddParagraph()
	// // add text
	// para1.AddText("test").AddTab()

	// para1.AddText("size").Size("44").AddTab()
	// tbl2 := doc.AddTableTwips([]int64{2333, 2333, 2333}, []int64{2333, 2333}).Justification("end")
	// tbl2.TableRows[0].TableCells[0].Shade("clear", "auto", "E7E6E6")

	// tbl2.TableRows[0].TableCells[0].AddParagraph().AddText("test11").AddTab()
	// tbl2.TableRows[0].TableCells[1].AddParagraph().AddText("test22").AddTab()
	// tbl2.TableRows[1].TableCells[1].AddParagraph().AddText("test22").AddTab()
	// tbl2.TableRows[2].TableCells[0].AddParagraph().Justification("center").AddText("test22").AddTab().Font("宋体", "宋体", "eastAsia").Bold()

	newFile.AppendFile(doc)

	_, err = io.Copy(f, newFile)
	if err != nil {
		log.Println(err)
	}
	defer f.Close()

}
