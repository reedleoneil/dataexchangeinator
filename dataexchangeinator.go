package main

import (
  "fmt"
  "io/ioutil"
  "strconv"
  "strings"
  "time"
  "os"

  "github.com/360EntSecGroup-Skylar/excelize"
)

const (
  FILE_NAME_COL = "A"
  FILE_IS_ACTIVE_COL = "K"
  FILE_TRANSFER_METHOD_COL = "L"
  STAGING_DIR_COL = "X"
  ARCHIVE_DIR_COL = "AA"
  FILE_AVAILABILITY = "G"
  DAY_UNAVAILABLE = "H"
  TIMESTAMP = "M"
  SHEET = "FILE META"
)

type dataExchange struct {
  spreadSheet *excelize.File
  fileName string
}

type DataExchange interface {
  GetRows() ([][]string, error)
  GetFileTransferMethod(row int) (string, error)
  GetFileStageDirectory(row int) (string, error)
  GetFileArchiveDirectory(row int) (string, error)
  IsFileActive(row int) (bool, error)
  SetFileTransferMethod(row int, value string)
  Save()
}

func (d *dataExchange) GetRows() ([][]string, error) {
  rows, err := d.spreadSheet.GetRows(SHEET)
  if err != nil {
    fmt.Println(err)
    return nil, err
  }
  return rows, nil
}

func (d *dataExchange) GetFileName(row int) (string, error) {
  cell, err := d.spreadSheet.GetCellValue(SHEET, FILE_NAME_COL + strconv.Itoa(row))
  if err != nil {
    fmt.Println(err)
    return "", err
  }
  return cell, nil
}

func (d *dataExchange) GetFileTransferMethod(row int) (string, error) {
  cell, err := d.spreadSheet.GetCellValue(SHEET, FILE_TRANSFER_METHOD_COL + strconv.Itoa(row))
  if err != nil {
    fmt.Println(err)
    return "", err
  }
  return cell, nil
}

func (d *dataExchange) GetFileStageDirectory(row int) (string, error) {
  cell, err := d.spreadSheet.GetCellValue(SHEET, STAGING_DIR_COL + strconv.Itoa(row))
  if err != nil {
    fmt.Println(err)
    return "", err
  }
  return cell, nil
}

func (d *dataExchange) GetFileArchiveDirectory(row int) (string, error) {
  cell, err := d.spreadSheet.GetCellValue(SHEET, ARCHIVE_DIR_COL + strconv.Itoa(row))
  if err != nil {
    fmt.Println(err)
    return "", err
  }
  return cell, nil
}

func (d *dataExchange) GetFileAvailability(row int) (string, error) {
  cell, err := d.spreadSheet.GetCellValue(SHEET, FILE_AVAILABILITY + strconv.Itoa(row))
  if err != nil {
    fmt.Println(err)
    return "", err
  }
  return cell, nil
}

func (d *dataExchange) GetDayUnavailable(row int) (string, error) {
  cell, err := d.spreadSheet.GetCellValue(SHEET, DAY_UNAVAILABLE + strconv.Itoa(row))
  if err != nil {
    fmt.Println(err)
    return "", err
  }
  return cell, nil
}

func (d *dataExchange) IsFileActive(row int) (bool, error) {
  cell, err := d.spreadSheet.GetCellValue(SHEET, FILE_IS_ACTIVE_COL + strconv.Itoa(row))
  if err != nil {
    fmt.Println(err)
    return false, err
  }
  isFileActive, _ := strconv.ParseBool(cell)
  return isFileActive, nil
}

func (d *dataExchange) SetFileTransferMethod(row int, value string) {
  d.spreadSheet.SetCellValue(SHEET, FILE_TRANSFER_METHOD_COL + strconv.Itoa(row), value)
}

func (d *dataExchange) SetTimestamp(row int) {
  d.spreadSheet.SetCellValue(SHEET, FILE_TRANSFER_METHOD_COL + strconv.Itoa(row), time.Now())
}

func (d *dataExchange) Save() {
  if err := d.spreadSheet.SaveAs(d.fileName); err != nil {
        fmt.Println(err)
  }
}

func NewDataExchange(filename string) (*dataExchange, error) {
  d := &dataExchange {}
  f, err := excelize.OpenFile(filename)
  if err != nil {
    return nil, err
  }
  d.spreadSheet = f
  d.fileName = filename
  return d, nil
}

func CompareTwoStrings(stringOne, stringTwo string) float32 {
	removeSpaces(&stringOne, &stringTwo)

	if value := returnEarlyIfPossible(stringOne, stringTwo); value >= 0 {
		return value
	}

	firstBigrams := make(map[string]int)
	for i := 0; i < len(stringOne)-1; i++ {
		a := fmt.Sprintf("%c", stringOne[i])
		b := fmt.Sprintf("%c", stringOne[i+1])

		bigram := a + b

		var count int

		if value, ok := firstBigrams[bigram]; ok {
			count = value + 1
		} else {
			count = 1
		}

		firstBigrams[bigram] = count
	}

	var intersectionSize float32
	intersectionSize = 0

	for i := 0; i < len(stringTwo)-1; i++ {
		a := fmt.Sprintf("%c", stringTwo[i])
		b := fmt.Sprintf("%c", stringTwo[i+1])

		bigram := a + b

		var count int

		if value, ok := firstBigrams[bigram]; ok {
			count = value
		} else {
			count = 0
		}

		if count > 0 {
			firstBigrams[bigram] = count - 1
			intersectionSize = intersectionSize + 1
		}
	}

	return (2.0 * intersectionSize) / (float32(len(stringOne)) + float32(len(stringTwo)) - 2)
}

func IsRecievedToday(t time.Time) bool {
	dayRecieved := t.Day()
  monthRecieved := t.Month()
  yearRecieved := t.Year()
	now := time.Now()
	dayNow := now.Day()
  monthNow := now.Month()
  yearNow := now.Year()

  if dayNow == dayRecieved && monthNow == monthRecieved && yearNow == yearRecieved {
    return true
  } else {
    return false
  }
}

func Step1() {
  filename := os.Args[2]
  day := os.Args[3]

  dataexchange, _ := NewDataExchange(filename)

  rows, _ := dataexchange.GetRows()
  for i, _ := range rows {
    isFileActive, _ := dataexchange.IsFileActive(i + 1)
    fileTransferMethod, _ := dataexchange.GetFileTransferMethod(i + 1)
	fileAvailability, _ := dataexchange.GetFileAvailability(i + 1)
	dayUnavailable, _ := dataexchange.GetDayUnavailable(i + 1)

	if isFileActive {
	  if fileTransferMethod == "Automatic" {
	    dataexchange.SetFileTransferMethod(i + 1, "A")
      } else if fileTransferMethod == "Manual" {
	    dataexchange.SetFileTransferMethod(i + 1, "M")
	  } else if isFileActive && fileTransferMethod == "Not Available" {
	    dataexchange.SetFileTransferMethod(i + 1, "N")
	  }

	  fa := strings.ToLower(fileAvailability)
	  du := strings.ToLower(dayUnavailable)
	  if fa != "daily" || strings.Contains(du, day) {
	  	dataexchange.SetFileTransferMethod(i + 1, "Not Available")
	  }
	}
  }

  dataexchange.Save()
}

func Step2() {
  filename := os.Args[2]
  fileMatchPercentage, _ := strconv.ParseFloat(os.Args[3], 32)

  dataexchange, _ := NewDataExchange(filename)

  rows, _ := dataexchange.GetRows()
  for i, _ := range rows {
    isFileActive, _ := dataexchange.IsFileActive(i + 1)
    fileName, _ := dataexchange.GetFileName(i + 1)
    fileTransferMethod, _ := dataexchange.GetFileTransferMethod(i + 1)
    fileStageDirectory, _ := dataexchange.GetFileStageDirectory(i + 1)
    fileArchiveDirectory, _ := dataexchange.GetFileArchiveDirectory(i + 1)
    if isFileActive && fileTransferMethod == "A" {
      var directory string
      if fileArchiveDirectory != "" {
        directory = fileArchiveDirectory
      } else {
        directory = fileStageDirectory
      }

      if directory != "" {
        files, err := ioutil.ReadDir(directory)
        if err != nil {
            fmt.Println(err)
        }

        for _, f := range files {
          if CompareTwoStrings(f.Name(), fileName) >= float32(fileMatchPercentage) && IsRecievedToday(f.ModTime()) {
            dataexchange.SetFileTransferMethod(i + 1, "Automatic")
            dataexchange.SetTimestamp(i + 1)
            fmt.Println("Automatic:" + fileName + " " + directory)
          }
        }
      }
    }
  }

  dataexchange.Save()
}

func main() {
  mode := os.Args[1]
  

  if mode == "1" {
	Step1()
  } else if mode == "2" {
	Step2()
  }

  
}

func removeSpaces(stringOne, stringTwo *string) {
	*stringOne = strings.Replace(*stringOne, " ", "", -1)
	*stringTwo = strings.Replace(*stringTwo, " ", "", -1)
}

func returnEarlyIfPossible(stringOne, stringTwo string) float32 {
	// if both are empty strings
	if len(stringOne) == 0 && len(stringTwo) == 0 {
		return 1
	}

	// if only one is empty string
	if len(stringOne) == 0 || len(stringTwo) == 0 {
		return 0
	}

	// identical
	if stringOne == stringTwo {
		return 1
	}

	// both are 1-letter strings
	if len(stringOne) == 1 && len(stringTwo) == 1 {
		return 0
	}

	// if either is a 1-letter string
	if len(stringOne) < 2 || len(stringTwo) < 2 {
		return 0
	}

	return -1
}
