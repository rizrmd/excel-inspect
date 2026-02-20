package main

import (
	"fmt"
	"log"
	"time"

	excelinspect "excel-inspect"
)

func main() {
	start := time.Now()

	ins, err := excelinspect.New("/Users/riz/Downloads/Inventory HQ - 19 Februari 2026.xlsx", excelinspect.WithTimeout(30))
	if err != nil {
		log.Fatalf("Failed to create inspector: %v", err)
	}
	defer ins.Close()

	fmt.Printf("Open file: %v\n", time.Since(start))

	start = time.Now()
	info, err := ins.Inspect()
	if err != nil {
		log.Fatalf("Failed to inspect: %v", err)
	}
	fmt.Printf("Inspect(): %v\n", time.Since(start))

	fmt.Printf("Sheets found: %d\n", len(info.Sheets))
	for _, s := range info.Sheets {
		fmt.Printf("  - %s: %d cols\n", s.Name, s.ColumnCount)
	}
}
