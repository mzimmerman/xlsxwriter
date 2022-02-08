package xlsxwriter

import (
	"bytes"
	"fmt"
	"os"
	"strconv"
	"testing"
)

func TestExcelChan(t *testing.T) {
	data := make([][]string, 50000)
	count := 0
	for x := range data {
		data[x] = make([]string, 200)
		for y := range data[x] {
			data[x][y] = strconv.Itoa(count)
			count++
		}
	}
	fo, err := os.Create(fmt.Sprintf("tmp-%s-chan.xlsx", "test"))
	if err != nil {
		t.Fatalf("error - %v", err)
		return
	}
	defer fo.Close()
	xw, err := New(fo)
	if err != nil {
		t.Fatalf("error - %v", err)
	}
	defer xw.Close()
	err = xw.WriteLines(data)
	if err != nil {
		t.Fatalf("error - %v", err)
	}
}

func TestExcel(t *testing.T) {
	data := [][]string{{"hi", "1234567890123456"}, {"12345678901234567890", "no"}}
	fo, err := os.Create(fmt.Sprintf("tmp-%s.xlsx", "test"))
	if err != nil {
		t.Fatalf("error - %v", err)
		return
	}
	defer fo.Close()
	xw, err := New(fo)
	if err != nil {
		t.Fatalf("error - %v", err)
	}
	defer xw.Close()
	for x := range data {
		err = xw.WriteLine(data[x])
		if err != nil {
			t.Fatalf("error - %v", err)
		}
	}
}

func TestColumnNumberToName(t *testing.T) {
	exp := map[int]string{
		0:  "A",
		1:  "B",
		25: "Z",
		26: "AA",
		27: "AB",
		51: "AZ",
		52: "BA",
		53: "BB",
	}
	for k, expected := range exp {
		got := columnNumberToName(k)
		if expected != got {
			t.Errorf("got %v, expected %v, for %d", got, expected, k)
		}
	}
}

func BenchmarkExcelize10x10(b *testing.B) {
	benchmarkExcelize(10, 10, b)
}

func BenchmarkExcelize100x100(b *testing.B) {
	benchmarkExcelize(100, 100, b)
}

func BenchmarkExcelize1000x1000(b *testing.B) {
	benchmarkExcelize(1000, 1000, b)
}

// func BenchmarkExcelize10000x10000(b *testing.B) {
// 	benchmarkExcelize(10000, 10000, b)
// }

func BenchmarkExcelize1000x10(b *testing.B) {
	benchmarkExcelize(1000, 10, b)
}

func BenchmarkExcelize10000x10(b *testing.B) {
	benchmarkExcelize(10000, 10, b)
}

func BenchmarkExcelize100000x10(b *testing.B) {
	benchmarkExcelize(100000, 10, b)
}

func BenchmarkExcelize100000x100(b *testing.B) {
	benchmarkExcelize(100000, 100, b)
}

func BenchmarkExcelize10000x1000(b *testing.B) {
	benchmarkExcelize(10000, 1000, b)
}

func benchmarkExcelize(rows, cols int, b *testing.B) {
	buf := bytes.Buffer{}
	for n := 0; n < b.N; n++ {
		b.StopTimer()
		buf.Reset()
		count := 0
		data := make([][]string, rows)
		for x := range data {
			data[x] = make([]string, cols)
			for y := range data[x] {
				data[x][y] = strconv.Itoa(count)
				count++
			}
		}
		b.StartTimer()
		xw, err := New(&buf)
		if err != nil {
			b.Fatalf("error writing excel - %v", err)
		}
		err = xw.WriteLines(data)
		if err != nil {
			b.Fatalf("error writing excel - %v", err)
		}
	}
}
