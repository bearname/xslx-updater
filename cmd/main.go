package main

import (
	"fmt"
	"github.com/jlaffaye/ftp"
	"github.com/robfig/cron/v3"
	"github.com/tealeg/xlsx"
	"log"
	"net/http"
	"os"
	"strconv"
	"time"
	"xlsx-parser/internal/domain"
)

func main() {
	port := os.Getenv("PORT")

	http.HandleFunc("/hello", func(w http.ResponseWriter, _ *http.Request) {
		fmt.Fprintf(w, "World")
	})

	http.ListenAndServe(":"+port, nil)

	parseFile()
	c := cron.New()

	c.AddFunc("0 7 * * *", func() {
		fmt.Println(time.Now().String() + " updated")
		parseFile()
	})
	c.AddFunc("0 12 * * *", func() {
		fmt.Println(time.Now().String() + " updated")
		parseFile()
	})
	c.AddFunc("0 17 * * *", func() {
		fmt.Println(time.Now().String() + " updated")
		parseFile()
	})

	c.Start()

	for {
		time.Sleep(time.Second)
	}
}

func parseFile() {
	fmt.Println(time.Now())
	prepareFile()
	uploadToTimeWeb("vh68.timeweb.ru", "cg77613", "1gqSlS0bWUqT")
}

func uploadToTimeWeb(host string, user string, password string) {
	c, err := ftp.Dial(host+":21", ftp.DialWithTimeout(5*time.Second))
	if err != nil {
		log.Println(err)
		return
	}

	err = c.Login(user, password)
	if err != nil {
		log.Println(err)
		return
	}

	f, err := os.Open("../MyXLSXFile.xlsx")
	if err != nil {
		log.Println(err)
		return
	}

	err = c.Stor("../MyXLSXFile.xlsx", f)
	if err != nil {
		log.Println(err)
		return
	}

	if err = c.Quit(); err != nil {
		log.Println(err)
		return
	}
}

func prepareFile() {
	wb, err := xlsx.OpenFile("C:\\Users\\mikha\\go\\src\\xlsx-parser\\ExcelPriceAllDep.xlsx")
	if err != nil {
		log.Println(err)
		return
	}

	results, ok := parseTiresSheet(wb)
	if !ok {
		log.Println(err)
		fmt.Println("failed parse")
		return
	}

	resultsDisks, ok := parseDisksSheet(wb)
	if !ok {
		log.Println(err)
		fmt.Println("failed parse")
		return
	}

	var file *xlsx.File

	file = xlsx.NewFile()
	err = fillSheet("Шины", file, results)
	err = fillSheet("Диски", file, resultsDisks)

	err = file.Save("MyXLSXFile.xlsx")
	if err != nil {
		log.Println(err)
	}
}

func parseTiresSheet(wb *xlsx.File) ([]domain.Result, bool) {
	sh, ok := wb.Sheet["Шины"]
	if !ok {
		fmt.Println("Шины does not exist")
		return nil, false
	}

	var results []domain.Result

	for i := 1; i < sh.MaxRow; i++ {
		cell := sh.Cell(i, 2)
		if cell.Value == "Легковые летние" || cell.Value == "Легкогрузовые летние" ||
			cell.Value == "Нешипуемые легкогрузовые" || cell.Value == "Нешипуемые легковые" ||
			cell.Value == "Легковые зимние" {
			sku := sh.Cell(i, 5)
			brand := sh.Cell(i, 3)
			stock := sh.Cell(i, 13)
			stock2 := sh.Cell(i, 12)
			price := sh.Cell(i, 17)
			atoi, _ := price.Int()
			results = append(results, domain.Result{
				Sku:    sku.Value,
				Brand:  brand.Value,
				Stock:  stock.Value,
				Stock2: stock2.Value,
				Price:  atoi,
			})
		}
	}
	return results, true
}

func parseDisksSheet(wb *xlsx.File) ([]domain.Result, bool) {
	sh, ok := wb.Sheet["Диски"]
	if !ok {
		fmt.Println("Диски does not exist")
		return nil, false
	}

	var results []domain.Result

	for i := 0; i < sh.MaxRow; i++ {
		sku := sh.Cell(i, 23)
		brand := sh.Cell(i, 1)
		stock := sh.Cell(i, 14)
		stock2 := sh.Cell(i, 13)
		price := sh.Cell(i, 18)
		atoi, _ := price.Int()
		results = append(results, domain.Result{
			Sku:    sku.Value,
			Brand:  brand.Value,
			Stock:  stock.Value,
			Stock2: stock2.Value,
			Price:  atoi,
		})
	}
	return results, true
}

func fillSheet(name string, file *xlsx.File, results []domain.Result) error {
	var sheet *xlsx.Sheet

	sheet, err := file.AddSheet(name)
	if err != nil {
		log.Println(err)
		return err
	}

	fillResult(sheet, results)
	return nil
}

func fillResult(sheet *xlsx.Sheet, results []domain.Result) {
	var row *xlsx.Row
	var cell *xlsx.Cell

	row = sheet.AddRow()
	cell = row.AddCell()
	cell.Value = "sku"
	cell = row.AddCell()
	cell.Value = "brand"
	cell = row.AddCell()
	cell.Value = "price"
	cell = row.AddCell()
	cell.Value = "stock"
	cell = row.AddCell()
	cell.Value = "stock2"

	for _, result := range results {
		fillRow(sheet, row, cell, result)
	}
}

func fillRow(sheet *xlsx.Sheet, row *xlsx.Row, cell *xlsx.Cell, result domain.Result) {
	row = sheet.AddRow()
	cell = row.AddCell()
	cell.Value = result.Sku
	cell = row.AddCell()
	cell.Value = result.Brand
	cell = row.AddCell()
	cell.Value = strconv.Itoa(result.Price)
	cell = row.AddCell()
	cell.Value = result.Stock
	cell = row.AddCell()
	cell.Value = result.Stock2
}
