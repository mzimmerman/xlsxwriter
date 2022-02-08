# xlsxwriter
Cheap/fast/simple XLSX file writer for textual data -- no fancy formatting or graphs

go get github.com/mzimmerman/xlsxwriter
```
	data := [][]string{
    {"hi","there"},
    {"you","fast"},
  }
	fo, err := os.Create("excel.xlsx")
	if err != nil {
		log.Fatalf("error - %v", err)
		return
	}
	defer fo.Close()
	xw, err := xlsxwriter.New(fo)
	if err != nil {
		log.Fatalf("error - %v", err)
	}
	defer xw.Close()
	err = xw.WriteLines(data)
	if err != nil {
		log.Fatalf("error - %v", err)
	}
}```
