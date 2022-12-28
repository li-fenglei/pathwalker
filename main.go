package main

import (
	"baliance.com/gooxml/document"
	"bufio"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
)

var dir = flag.String("dir", "", "")
var ext1 = flag.String("ext1", ".docx", "")
var ext2 = flag.String("ext2", ".doc", "")
var files []string

func WriteFile(files []string) {
	file, err := os.OpenFile("path.txt", os.O_WRONLY|os.O_TRUNC|os.O_CREATE, 0666)
	defer file.Close()
	if err != nil {
		log.Fatal("文件打开失败", err)
		return
	}
	write := bufio.NewWriter(file)
	for _, path := range files {
		write.WriteString(path + "\n")
	}
	write.Flush()
}

func visit() filepath.WalkFunc {
	return func(path string, info os.FileInfo, err error) error {
		if err != nil {
			log.Fatal(err)
		}
		if filepath.Ext(path) == *ext1 || filepath.Ext(path) == *ext2 {
			ReadWord(path, &files)
		}
		return nil
	}
}

func ReadWord(filename string, files *[]string) {
	doc, err := document.Open(filename)
	if err != nil {
		log.Fatalf("error opening document: %s", err)
	}
	/**
	  遍历word文档中的表格
	*/
	for _, table := range doc.Tables() {
		//fmt.Println("-----------第", i, "个表格-------------")
		for _, row := range table.Rows() {
			//fmt.Println("-----------第", j, "表格行-------------")
			for _, cell := range row.Cells() {
				//fmt.Println("-----------第", k, "列-------------")
				for _, para := range cell.Paragraphs() {
					//fmt.Println("-----------第", n, "段-------------")
					for _, run := range para.Runs() {
						//fmt.Println("-----------第", m, "格式片段-------------")
						t := run.Text()
						if len(t) > 0 {
							if strings.HasSuffix(t, "png") || strings.HasSuffix(t, "svg") || strings.HasSuffix(t, "jpg") {
								blankPos := strings.Index(t, "：")
								if blankPos > 0 {
									ft := t[blankPos+3:]
									ft = strings.TrimSpace(ft)
									*files = append(*files, filename+"\\"+ft)
									//fmt.Println(ft)
								} else {
									blankPos = strings.Index(t, ":")
									if blankPos > 0 {
										fft := t[blankPos+1:]
										fft = strings.TrimSpace(fft)
										*files = append(*files, filename+"\\"+fft)
										//fmt.Println(fft)
									} else {
										*files = append(*files, filename+"\\"+t)
										//fmt.Println(t)
									}
								}
							}
						}
					}
				}
			}
		}
	}

	// doc.Paragraphs()得到包含文档所有的段落的切片
	/**
	遍历word中的段落
	*/
	for _, para := range doc.Paragraphs() {
		//run为每个段落相同格式的文字组成的片段
		//fmt.Println("-----------第", i, "段-------------")
		for _, run := range para.Runs() {
			//fmt.Print("\t-----------第", j, "格式片段-------------")
			//fmt.Print(run.Text())
			t := run.Text()
			if len(t) > 0 {
				if strings.HasSuffix(t, "png") || strings.HasSuffix(t, "svg") || strings.HasSuffix(t, "jpg") {
					blankPos := strings.Index(t, "：")
					if blankPos > 0 {
						ft := t[blankPos+3:]
						ft = strings.TrimSpace(ft)
						*files = append(*files, filename+"\\"+ft)
						//fmt.Println(ft)
					} else {
						blankPos = strings.Index(t, ":")
						if blankPos > 0 {
							fft := t[blankPos+1:]
							fft = strings.TrimSpace(fft)
							*files = append(*files, filename+"\\"+fft)
							//fmt.Println(fft)
						} else {
							//fmt.Println(t)
							*files = append(*files, filename+"\\"+t)
						}
					}
				}
			}
		}
		//fmt.Println()
	}
}

func main() {

	flag.Parse()
	root := *dir

	err := filepath.Walk(root, visit())
	if err != nil {
		panic(err)
	}
	for _, file := range files {
		fmt.Println(file)
	}
	WriteFile(files)
}
