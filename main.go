package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	// Conversion factors
	metersToFeet  = 3.28084     // Meters to Feet
	metersToMiles = 0.000621371 // Meters to Miles
	mpsToMph      = 2.23694     // Meters per Second to Miles per Hour
)

// Simple converter for HH:MM format (24-hour)
func convertTimeStringToExcel(timeStr string) float64 {
	// Split the time string by colon
	parts := strings.Split(timeStr, ":")
	if len(parts) != 2 {
		fmt.Println("Invalid time format, expected HH:MM")
		return 0
	}

	// Parse hours and minutes
	hours, err := strconv.Atoi(parts[0])
	if err != nil {
		fmt.Println("Error parsing hours:", err)
		return 0
	}

	minutes, err := strconv.Atoi(parts[1])
	if err != nil {
		fmt.Println("Error parsing minutes:", err)
		return 0
	}

	// Convert to Excel time value (fraction of a day)
	return (float64(hours) + float64(minutes)/60.0) / 24.0
}

func convertDateStringToExcel(dateStr string) float64 {
	// Parse the date string to a time.Time object using YY/MM/DD format
	layout := "06/01/02" // For YY/MM/DD format (Go's reference date is 2006-01-02)
	t, err := time.Parse(layout, dateStr)
	if err != nil {
		fmt.Println("Error parsing date:", err)
		return 0
	}

	// Handle potential year ambiguity (assuming 20xx for simplicity)
	// This is only needed if you want specific century handling

	// Convert to Excel date value
	// Excel uses January 1, 1900 as day 1 (serial number 1)
	baseDate := time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC)

	// Calculate days since baseDate
	duration := t.Sub(baseDate)
	days := duration.Hours() / 24

	// Add 1 because Excel counts January 1, 1900, as day 1 (not 0)
	excelDate := days + 1

	// Add 1 more if the date is after February 28, 1900, to account for Excel's leap year bug
	//goland:noinspection GoSnakeCaseUsage
	feb28_1900 := time.Date(1900, 2, 28, 0, 0, 0, 0, time.UTC)
	if t.After(feb28_1900) {
		excelDate += 1
	}

	return excelDate
}

func main() {
	if len(os.Args) < 3 {
		fmt.Printf("Usage: %v input.csv output.xlsx\n", filepath.Base(os.Args[0]))
		return
	}

	inputFile := os.Args[1]
	outputFile := os.Args[2]

	// Open the CSV file
	file, err := os.Open(inputFile)
	if err != nil {
		fmt.Printf("Error opening file: %v\n", err)
		return
	}
	//goland:noinspection GoUnhandledErrorResult
	defer file.Close()

	// Create a new CSV reader
	reader := csv.NewReader(file)

	// Create a new Excel file
	xlsx := excelize.NewFile()
	sheetName := "Sheet1"

	// Read the header row
	headers, err := reader.Read()
	if err != nil {
		fmt.Printf("Error reading header: %v\n", err)
		return
	}

	// Write headers to Excel
	for i, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		err := xlsx.SetCellValue(sheetName, cell, header)
		if err != nil {
			fmt.Printf("Error writing header in cell %v: %v\n", cell, err)
		}
	}

	// Define the columns that need conversion
	feetColumns := []string{"exitAlt", "openAlt"}
	milesColumns := []string{"exitDist", "openDist", "cpDist", "ffDist"}
	mphColumns := []string{"ffAvgVSpd", "ffMaxVSpd", "cpAvgVSpd", "cpMaxVSpd"}

	// Process the data rows
	rowIndex := 2 // Start from row 2 (after header)
	for {
		record, err := reader.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			fmt.Printf("Error reading record: %v\n", err)
			continue
		}

		// Convert columns from meters to feet or miles
		for i, value := range record {
			colName := headers[i]

			// Skip if the value is empty
			if value == "" {
				continue
			}

			// Convert to feet
			for _, feetCol := range feetColumns {
				if colName == feetCol {
					if valueInMeters, err := strconv.ParseFloat(value, 64); err == nil {
						valueInFeet := valueInMeters * metersToFeet
						record[i] = fmt.Sprintf("%.2f", valueInFeet)
					}
					break
				}
			}

			// Convert to miles
			for _, milesCol := range milesColumns {
				if colName == milesCol {
					if valueInMeters, err := strconv.ParseFloat(value, 64); err == nil {
						valueInMiles := valueInMeters * metersToMiles
						record[i] = fmt.Sprintf("%.3f", valueInMiles)
					}
					break
				}
			}

			// Convert to miles per hour
			for _, mphCol := range mphColumns {
				if colName == mphCol {
					if valueMps, err := strconv.ParseFloat(value, 64); err == nil {
						valueMph := valueMps * mpsToMph
						record[i] = fmt.Sprintf("%.1f", valueMph)
					}
					break
				}
			}
		}

		// Write the processed row to Excel
		for i, value := range record {
			cell, _ := excelize.CoordinatesToCellName(i+1, rowIndex)

			// Convert numeric values with appropriate formatting
			colName := headers[i]

			// Columns that should be whole numbers
			integerColumns := map[string]bool{
				"num":    true,
				"ffSecs": true, "cpSecs": true,
				"cpAvgSpd": true, "cpMaxVSpd": true,
				"aircraftSecs": true,
			}

			// Try to parse as a number for numeric columns
			if colName == "time" {
				err := xlsx.SetCellValue(sheetName, cell, convertTimeStringToExcel(value))
				if err != nil {
					fmt.Printf("Error setting time value for cell %v: %v\n", cell, err)
				}
				continue
			}
			if colName == "date" {
				err := xlsx.SetCellValue(sheetName, cell, convertDateStringToExcel(value))
				if err != nil {
					fmt.Printf("Error setting date value for cell %v: %v\n", cell, err)
				}
				continue
			}

			if floatVal, err := strconv.ParseFloat(value, 64); err == nil {
				// Format as integer for specific columns
				if integerColumns[colName] {
					intVal := int(floatVal)
					err := xlsx.SetCellValue(sheetName, cell, intVal)
					if err != nil {
						fmt.Printf("Error setting integer value for cell %v: %v\n", cell, err)
					}
				} else {
					err := xlsx.SetCellValue(sheetName, cell, floatVal)
					if err != nil {
						fmt.Printf("Error setting numeric value for cell %v: %v\n", cell, err)
					}
				}
				continue
			}

			// Set as string if not numeric or if in date/time columns
			err := xlsx.SetCellValue(sheetName, cell, value)
			if err != nil {
				fmt.Printf("Error setting value for cell %v: %v\n", cell, err)
			}
		}

		rowIndex++
	}

	// Create a style for the header row
	style, err := xlsx.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#E0E0E0"},
			Pattern: 1,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: []excelize.Border{
			{Type: "bottom", Color: "#000000", Style: 1},
		},
	})
	if err != nil {
		fmt.Printf("Error creating header style: %v\n", err)
	} else {
		// Apply the style to the header row
		for i := range headers {
			cell, _ := excelize.CoordinatesToCellName(i+1, 1)
			err := xlsx.SetCellStyle(sheetName, cell, cell, style)
			if err != nil {
				fmt.Printf("Error setting header style for cell %v: %v\n", cell, err)
			}
		}
	}

	// Define a custom formatter for the time column
	timeFormatCode := "h:mm AM/PM"
	timeStyle, err := xlsx.NewStyle(&excelize.Style{
		CustomNumFmt: &timeFormatCode,
	})
	if err != nil {
		fmt.Printf("Error creating time style: %v\n", err)
	}

	// Define a custom formatter for the date column
	dateFormatCode := "ddd, mmm d, yyyy"
	dateStyle, err := xlsx.NewStyle(&excelize.Style{
		CustomNumFmt: &dateFormatCode,
	})
	if err != nil {
		fmt.Printf("Error creating date style: %v\n", err)
	}

	// Define a custom number format for altitude columns
	// Format with condition: >=1000 shows as "13.5K ft", <1000 shows as "850 ft"
	altitudeFormatCode := `[>=1000]#,##0.0,"K ft";#,##0" ft"`
	altitudeStyle, err := xlsx.NewStyle(&excelize.Style{
		CustomNumFmt: &altitudeFormatCode,
	})
	if err != nil {
		fmt.Printf("Error creating altitude style: %v\n", err)
	}

	// Define a custom formatter for miles
	milesFormatCode := `0.00" mi"`
	milesStyle, err := xlsx.NewStyle(&excelize.Style{
		CustomNumFmt: &milesFormatCode,
	})
	if err != nil {
		fmt.Printf("Error creating miles style: %v\n", err)
	}

	// And one for Miles per Hour
	mphFormatCode := `0.0" mph"`
	mphStyle, err := xlsx.NewStyle(&excelize.Style{
		CustomNumFmt: &mphFormatCode,
	})
	if err != nil {
		fmt.Printf("Error creating mph style: %v\n", err)
	}

	// Freeze the header row
	err = xlsx.SetPanes(sheetName, &excelize.Panes{
		Freeze:      true,
		Split:       false,
		XSplit:      3,    // Freeze columns A, B, C (3 columns)
		YSplit:      1,    // Freeze row 1
		TopLeftCell: "D2", // Active cell starts at D2
		ActivePane:  "bottomRight",
	})
	if err != nil {
		fmt.Printf("Error freezing header row: %v\n", err)
		return
	}

	// Apply custom formats to data columns
	for rowNum := 2; rowNum < rowIndex; rowNum++ {
		for i, header := range headers {
			if header == "time" {
				cell, _ := excelize.CoordinatesToCellName(i+1, rowNum)
				err := xlsx.SetCellStyle(sheetName, cell, cell, timeStyle)
				if err != nil {
					fmt.Printf("Error setting time style for cell %v: %v\n", cell, err)
				}
			}
			if header == "date" {
				cell, _ := excelize.CoordinatesToCellName(i+1, rowNum)
				err := xlsx.SetCellStyle(sheetName, cell, cell, dateStyle)
				if err != nil {
					fmt.Printf("Error setting date style for cell %v: %v\n", cell, err)
				}
			}

			// Apply altitude format to altitude columns
			if header == "exitAlt" || header == "openAlt" {
				cell, _ := excelize.CoordinatesToCellName(i+1, rowNum)
				err := xlsx.SetCellStyle(sheetName, cell, cell, altitudeStyle)
				if err != nil {
					fmt.Printf("Error setting altitude style for cell %v: %v\n", cell, err)
				}
			}

			if header == "exitDist" || header == "openDist" || header == "cpDist" || header == "ffDist" {
				cell, _ := excelize.CoordinatesToCellName(i+1, rowNum)
				err := xlsx.SetCellStyle(sheetName, cell, cell, milesStyle)
				if err != nil {
					fmt.Printf("Error setting miles style for cell %v: %v\n", cell, err)
				}
			}

			if header == "ffAvgVSpd" || header == "ffMaxVSpd" || header == "cpAvgVSpd" || header == "cpMaxVSpd" {
				cell, _ := excelize.CoordinatesToCellName(i+1, rowNum)
				err := xlsx.SetCellStyle(sheetName, cell, cell, mphStyle)
				if err != nil {
					fmt.Printf("Error setting mph style for cell %v: %v\n", cell, err)
				}
			}
		}
	}

	// Save the Excel file
	// First, set column widths for better readability
	for i, header := range headers {
		colName, _ := excelize.ColumnNumberToName(i + 1)

		// Set a decent default width for columns
		width := 12.0

		// Adjust width based on the header type
		switch header {
		case "num":
			width = 8.0
		case "date":
			width = 18.0
		case "time":
			width = 10.0
		case "exitAlt", "openAlt":
			width = 12.0
		case "ffSecs", "cpSecs", "aircraftSecs":
			width = 10.0
		case "ffAvgVSpd", "ffMaxVSpd", "cpAvgVSpd", "cpMaxVSpd":
			width = 14.0
		case "ffAvgGlide", "cpAvgGlide":
			width = 12.0
		default:
			// Wider columns for distance columns (in miles)
			if strings.HasSuffix(header, "Dist") {
				width = 14.0
			} else {
				width = 12.0
			}
		}

		err := xlsx.SetColWidth(sheetName, colName, colName, width)
		if err != nil {
			fmt.Printf("Error setting column width for column %v: %v\n", colName, err)
		}
	}

	if err := xlsx.SaveAs(outputFile); err != nil {
		fmt.Printf("Error saving Excel file: %v\n", err)
		return
	}

	fmt.Printf("Successfully converted %s to %s\n", inputFile, outputFile)
	fmt.Println("Conversions applied:")
	fmt.Println("- Meters to Feet columns:", feetColumns)
	fmt.Println("- Meters to Miles columns:", milesColumns)
	fmt.Println("- Meters/sec to Miles/hour columns:", mphColumns)
}
