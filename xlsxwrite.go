package xlsxwriter

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"sync"

	"github.com/klauspost/compress/flate"
)

type XLSXWriter struct {
	w       io.Writer
	zw      *zip.Writer
	sheet   io.Writer
	lineNum int
}

func New(w io.Writer) (*XLSXWriter, error) {
	zw := zip.NewWriter(w)
	zw.RegisterCompressor(8, func(w io.Writer) (io.WriteCloser, error) {
		return flate.NewWriter(w, flate.DefaultCompression)
	})
	for name := range files {
		sheet, err := zw.Create(name)
		if err != nil {
			return nil, err
		}
		_, err = sheet.Write(files[name])
		if err != nil {
			return nil, err
		}
	}
	sheet, err := zw.Create("xl/worksheets/sheet1.xml")
	if err != nil {
		return nil, err
	}
	_, err = sheet.Write([]byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="A1:B2"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/><sheetData>`))
	return &XLSXWriter{
		w:       w,
		zw:      zw,
		sheet:   sheet,
		lineNum: 1,
	}, err
}

func (xw *XLSXWriter) Close() error {
	_, err := xw.sheet.Write([]byte(`<autoFilter/></sheetData></worksheet>`))
	if err != nil {
		return err
	}
	return xw.zw.Close()
}

var bufPool = sync.Pool{
	New: func() interface{} {
		return &bytes.Buffer{}
	},
}

func getBuf() *bytes.Buffer {
	buf := bufPool.Get().(*bytes.Buffer)
	buf.Reset()
	return buf
}

func putBuf(buf *bytes.Buffer) {
	bufPool.Put(buf)
}

type line struct {
	idx  int
	data []string
	buf  *bytes.Buffer
}

func (xw *XLSXWriter) WriteLine(data []string) error {
	err := writeLine(xw.sheet, xw.lineNum, data)
	if err == nil {
		xw.lineNum++
	}
	return err
}

func writeLine(w io.Writer, lineNum int, data []string) error {
	if len(data) > 16383 {
		return fmt.Errorf("Excel only supports 16383 columns, but this data has %d", len(data))
	}
	_, err := fmt.Fprintf(w, `<row r="%d">`, lineNum)
	if err != nil {
		return err
	}
	for c, v := range data {
		_, err = fmt.Fprintf(w, `<c r="%s%d" t="inlineStr"><is><t>`, columnNumberToName(c), lineNum)
		if err != nil {
			return err
		}
		err = xml.EscapeText(w, []byte(v))
		if err != nil {
			return err
		}
		_, err = w.Write([]byte(`</t></is></c>`))
		if err != nil {
			return err
		}
	}
	_, err = w.Write([]byte(`</row>`))
	return err
}

func columnNumberToName(c int) string {
	return columns[c]
}
