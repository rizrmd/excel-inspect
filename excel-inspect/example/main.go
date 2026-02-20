package main

import (
	"encoding/json"
	"fmt"
	"log"
	"time"

	excelinspect "excel-inspect"
)

func main() {
	start := time.Now()

	ins, err := excelinspect.New("/Users/riz/Downloads/Inventory HQ - 19 Februari 2026.xlsx")
	if err != nil {
		log.Fatalf("Failed to create inspector: %v", err)
	}
	defer ins.Close()

	fmt.Printf("Open file: %v\n", time.Since(start))

	start = time.Now()
	info, err := ins.InspectWithDetails()
	if err != nil {
		log.Fatalf("Failed to inspect: %v", err)
	}
	fmt.Printf("InspectWithDetails(): %v\n", time.Since(start))

	fmt.Printf("Sheets found: %d\n", len(info.Sheets))
	for _, s := range info.Sheets {
		fmt.Printf("  - %s: %d rows, %d cols\n", s.Name, s.RowCount, s.ColumnCount)
	}

	if len(info.SheetDetails) > 0 {
		sd := info.SheetDetails[0]
		fmt.Printf("\nFirst sheet (%s) columns:\n", sd.Name)
		for _, col := range sd.Columns {
			b, _ := json.Marshal(col)
			fmt.Printf("  %s\n", string(b))
		}
	}
}
