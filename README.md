# excel-inspect

Go library to inspect Excel file structure: sheets, headers, columns, row counts, and sample values.

## Installation

```bash
go get github.com/yourusername/excel-inspect
```

## Usage

```go
package main

import (
    "encoding/json"
    "fmt"
    "log"

    excelinspect "github.com/yourusername/excel-inspect"
)

func main() {
    ins, err := excelinspect.New("file.xlsx")
    if err != nil {
        log.Fatalf("Failed to create inspector: %v", err)
    }
    defer ins.Close()

    // Quick inspection (sheets only)
    info, err := ins.Inspect()
    if err != nil {
        log.Fatalf("Failed to inspect: %v", err)
    }
    fmt.Printf("Sheets: %d\n", len(info.Sheets))

    // Full inspection with details
    info, err = ins.InspectWithDetails()
    if err != nil {
        log.Fatalf("Failed to inspect: %v", err)
    }

    // Print as JSON
    b, _ := json.MarshalIndent(info, "", "  ")
    fmt.Println(string(b))
}
```

## Options

### Timeout

For large files, set a timeout to prevent hanging:

```go
ins, err := excelinspect.New("large_file.xlsx", excelinspect.WithTimeout(30))
```

## Output Structure

```json
{
  "sheets": [
    {"name": "Sheet1", "row_count": 1000, "column_count": 10}
  ],
  "sheet_details": [
    {
      "name": "Sheet1",
      "row_count": 1000,
      "column_count": 10,
      "headers": ["ID", "Name", "Email"],
      "columns": [
        {
          "name": "ID",
          "start_position": "A1",
          "sample_values": ["1", "2", "3", "4", "5"],
          "data_type": "number"
        }
      ]
    }
  ]
}
```

## Features

- **Hidden sheets**: Automatically skipped
- **Large files**: Optimized to read only headers and sample values (not entire file)
- **Timeout**: Configurable timeout for large files
- **Thread-safe**: Can be used concurrently with proper locking
