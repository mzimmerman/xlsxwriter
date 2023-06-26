package xlsxwriter

import (
	"bytes"
	"fmt"
	"os"
	"strconv"
	"testing"
)

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

func TestLongData(t *testing.T) {
	data := [][]string{make([]string, 100), make([]string, 1000), make([]string, 10000), make([]string, 100000)}
	for _, slice := range data {
		for x := 0; x < len(slice); x++ {
			slice[x] = strconv.Itoa(x)
		}
	}
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
		if x == 3 {
			if err == nil || err.Error() != "Excel only supports 16383 columns, but this data has 100000" {
				t.Fatalf("should have detected an error that Excel does not support more than 16383 columns - got %v", err)
			}
		} else if err != nil {
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
		for _, line := range data {
			err = xw.WriteLine(line)
			if err != nil {
				b.Fatalf("error writing excel - %v", err)
			}
		}
	}
}
