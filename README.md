# xlssheetmerger

## Usage:

```
$ go run merger.go agent.xlsx out.xlsx 2019-03-19 1 7 Sheet1 Sheet2 Sheet3
```

## Build exe

```
GOOS=windows GOARCH=amd64 go build -v .
```