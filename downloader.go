package main

import (
	"fmt"
	"io"
	"net/http"
	"os"
)

func main() {
	url := "https://drive.google.com/u/2/uc?id=1ExaVtTI1QQZP0CeyTQW9slzyFJj9yuGw&export=download"
	fileName := "file.xlsx"
	fmt.Println("Downloading file...")

	output, err := os.Create(fileName)
	defer output.Close()

	response, err := http.Get(url)
	if err != nil {
		fmt.Println("Error while downloading", url, "-", err)
		return
	}
	defer response.Body.Close()

	n, err := io.Copy(output, response.Body)

	fmt.Println(n, "bytes downloaded")
}
