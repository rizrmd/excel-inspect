package main

import (
	"fmt"
	"log"
	"os"
	"time"

	excelinspect "excel-inspect"
)

func main() {
	start := time.Now()

	outPath := "out.txt"
	if err := os.Remove(outPath); err != nil && !os.IsNotExist(err) {
		log.Fatalf("Failed to delete existing %s: %v", outPath, err)
	}

	ins, err := excelinspect.New(
		"/Users/riz/Downloads/Inventory HQ - 19 Februari 2026.xlsx",
		excelinspect.WithProgressCallback(func(p excelinspect.ProgressInfo) {
			if p.Sheet != "" {
				fmt.Printf("[progress] %s | %s | %.1f%% (%d/%d)\n", p.Phase, p.Sheet, p.Percent, p.Current, p.Total)
				return
			}
			fmt.Printf("[progress] %s | %.1f%% (%d/%d)\n", p.Phase, p.Percent, p.Current, p.Total)
		}),
	)
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

	toonOut, err := ins.InspectWithDetailsTOON()
	if err != nil {
		log.Fatalf("Failed to encode TOON: %v", err)
	}
	if err := os.WriteFile(outPath, []byte(toonOut), 0o644); err != nil {
		log.Fatalf("Failed to write %s: %v", outPath, err)
	}
	fmt.Printf("\nTOON output written to %s\n", outPath)
}
