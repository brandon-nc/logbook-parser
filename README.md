# logbook-parser
This is just a little Go program skydiovers can use to convert units from metric to imperial within Deep and Steep logbook.csv files.

# Compiling
Compile the program with:
```sh
> go build -o logbook-parser main.go
```

# Running
Pass in the path to you logbook.csv as the first argument, and supply an output xlsx filename as the second argument:
```sh
> ./logbook-parser ~/Desktop/logbook.csv ~/Desktop/logbook.xlsx
Successfully converted /Users/brandon/Desktop/logbook.csv to /Users/brandon/Desktop/logbook.xlsx
Conversions applied:
- Meters to Feet columns: [exitAlt openAlt]
- Meters to Miles columns: [exitDist openDist cpDist ffDist]
- Meters/sec to Miles/hour columns: [ffAvgVSpd ffMaxVSpd cpAvgVSpd cpMaxVSpd]
> 
```

# Samples
## Before:
![alt text](https://github.com/brandon-nc/logbook-parser/raw/main/logbook.csv.png "Original CSV")
## After:
![alt text](https://github.com/brandon-nc/logbook-parser/raw/main/logbook.xlsx.png "Converted XLSX")
