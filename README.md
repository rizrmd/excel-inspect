# excel-inspect

`excel-inspect` is a Go library for inspecting `.xlsx` files and exporting results as JSON-compatible structs or TOON.

## What Is In This Codebase

- `inspect.go`: library implementation (`package excelinspect`)
- `example/main.go`: runnable example that inspects one file and writes TOON output to `out.txt`
- `go.mod` / `go.sum`: module and dependencies
- `out.txt`: generated output file used by the example

## Core Capabilities

- Open an Excel workbook and skip hidden sheets
- Inspect sheet metadata (`name`, `row_count`, `column_count`)
- Inspect detailed sheet data:
  - detected headers
  - column metadata (`name`, `start_position`, `data_type`)
  - sample values
  - section breakdown when multiple header regions are detected
- Export as:
  - Go structs (`*FileInfo`)
  - TOON text output (`InspectTOON`, `InspectWithDetailsTOON`, `InspectWithDetailsTOONSample`)
- Emit progress updates via callback or channel

## Public API

Constructor and lifecycle:

- `New(filePath string, opts ...InspectorOption) (*Inspector, error)`
- `(*Inspector).Close() error`

Inspection methods:

- `(*Inspector).Inspect() (*FileInfo, error)`
- `(*Inspector).InspectWithDetails() (*FileInfo, error)`
- `(*Inspector).InspectTOON() (string, error)`
- `(*Inspector).InspectWithDetailsTOON() (string, error)`
- `(*Inspector).InspectWithDetailsTOONSample() (string, error)`

Options currently wired in:

- `WithProgressCallback(func(ProgressInfo))`
- `WithProgressChannel(chan<- ProgressInfo)`

Defined but currently no-op in `inspect.go`:

- `WithTimeout(int)`
- `WithHeaderRow(int)`
- `WithMaxSampleRows(int)`
- `WithIncludeRowCount(bool)`

## Usage

Because `go.mod` currently declares `module excel-inspect`, import it as:

```go
import excelinspect "excel-inspect"
```

Minimal example:

```go
package main

import (
	"fmt"
	"log"

	excelinspect "excel-inspect"
)

func main() {
	ins, err := excelinspect.New("file.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	defer ins.Close()

	info, err := ins.InspectWithDetails()
	if err != nil {
		log.Fatal(err)
	}

	fmt.Printf("sheets: %d\n", len(info.Sheets))
}
```

## Example Program

Run:

```bash
go run ./example
```

What it does:

- removes existing `out.txt` if present
- opens a hardcoded workbook path in `example/main.go`
- prints progress to stdout
- runs `InspectWithDetails()`
- writes full TOON output to `out.txt`

If you run the example locally, update the workbook path in `example/main.go` first.
