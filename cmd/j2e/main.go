package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/franissirkovic/excel2json/e2j"

	"github.com/xuri/excelize/v2"
)

var (
	inputFileName  = ""
	outputFileName = ""
	noOutput       bool
	csvFile        bool
	help           bool
)

func init() {
	const (
		SHORTHAND      = " (shorthand)"
		HELP_HELP      = "show usage help and exit"
		HELP_OUTPUT    = "output file name to be used"
		HELP_NO_OUTPUT = "no output file should be created"
		HELP_CSV       = "csv file should be created"
	)
	flag.BoolVar(&help, "help", false, HELP_HELP)
	flag.BoolVar(&help, "h", false, HELP_HELP+SHORTHAND)
	flag.BoolVar(&noOutput, "no_output", false, HELP_NO_OUTPUT)
	flag.BoolVar(&noOutput, "n", false, HELP_NO_OUTPUT+SHORTHAND)
	flag.BoolVar(&csvFile, "csv", false, HELP_CSV)
	flag.BoolVar(&csvFile, "c", false, HELP_CSV+SHORTHAND)
	flag.StringVar(&outputFileName, "output", "", HELP_OUTPUT)
	flag.StringVar(&outputFileName, "o", "", HELP_OUTPUT+SHORTHAND)
}

func usage() {
	fmt.Fprintf(flag.CommandLine.Output(), "Usage: %s [flag...] input file name\n", os.Args[0])
	flag.PrintDefaults()
}

func main() {
	fmt.Println("Start converter")
	flag.Parse()
	if help {
		usage()
		return
	}
	inputFileName = flag.Arg(0)
	if inputFileName == "" {
		usage()
		return
	}
	wb := e2j.WorkBook{}
	dat, err := os.ReadFile(inputFileName)
	if nil != err {
		fmt.Println(err)
		return
	}
	err = json.Unmarshal(dat, &wb)
	if nil != err {
		fmt.Println(err)
		return
	}
	// fmt.Println(wb)
	f := excelize.NewFile()
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	wb.Store(f)
	if outputFileName == "" {
		ext := filepath.Ext(inputFileName)
		outputFileName = strings.TrimSuffix(inputFileName, ext)
		outputFileName = strings.Join([]string{outputFileName, "xlsx"}, ".")
	}
	if !noOutput {
		err = f.SaveAs(outputFileName)
		if nil != err {
			fmt.Println(err)
			return
		}
	}
	if csvFile {
		//ext := filepath.Ext(outputFileName)
		//csvFileName := strings.TrimSuffix(outputFileName, ext)
		csvFileName := strings.Join([]string{outputFileName, "csv"}, ".")
		err = wb.ToCsv(csvFileName, ",")
		if nil != err {
			fmt.Println(err)
			return
		}
	}
	fmt.Println("Conversion OK")
}
