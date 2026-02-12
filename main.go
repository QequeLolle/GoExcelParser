package main

import (
	"encoding/json"
	"errors"
	"fmt"
	"log"
	"math"
	"os"
	"time"

	"github.com/xuri/excelize/v2"
)

type PhoneCall struct {
	CallID    int    `json:"call_id"`   // call id
	From      string `json:"from"`      // user number 1
	To        string `json:"to"`        // user number 2
	Talktime  int    `json:"talktime"`  // seconds
	Timestamp int64  `json:"timestamp"` // UNIX seconds
}

type ExcelTime struct {
	Seconds int
	Minutes int
	Hours   int
}

// ===================

const (
	ReportName string = "Отчет по звонкам"
	SheetName  string = "Sheet1"
)

var (
	ExcelTemplateFilepath string = os.Args[1]
	JSONFilepath          string = os.Args[2]
	ExcelOutputFilepath   string = os.Args[3]
)

// ===================

// read call data file
func readCallsDataFile() ([]PhoneCall, error) {
	file, err := os.Open(JSONFilepath)
	if err != nil {
		fmt.Println(err)
		return nil, err
	}

	var calls []PhoneCall

	decoder := json.NewDecoder(file)
	if err = decoder.Decode(&calls); err != nil {
		fmt.Println(err)
		return nil, err
	}

	if err = file.Close(); err != nil {
		fmt.Println(err)
		return nil, err
	}

	return calls, nil
}

// convert seconds into hh:mm:ss format
func convertSeconds(seconds int, hours_enable bool) ExcelTime {
	var result ExcelTime

	if hours_enable {

		if seconds < 3600 {
			result.Seconds = seconds % 60
			result.Minutes = seconds / 60
		} else {
			result.Seconds = seconds % 3600 % 60
			result.Minutes = seconds % 3600 / 60
			result.Hours = seconds / 3600
		}

	} else {

		result.Seconds = seconds % 60
		result.Minutes = seconds / 60
		result.Hours = -1

	}

	return result

}

// convert Unix time to dd.mm.yyyy date format
func convertUnixTimestampToDateStr(timestamp int64) string {
	return time.Unix(timestamp, 0).Format("02.01.2006")
}

// set cell value by tag in excel file.
// tag is a formated string:
//
//	"#<tag>"
//
// supported value types:
//
//	int
//	int8
//	int16
//	int32
//	int64
//	uint
//	uint8
//	uint16
//	uint32
//	uint64
//	float32
//	float64
//	string
//	[]byte
//	time.Duration
//	time.Time
//	bool
//	nil
func setCellByTag(file *excelize.File, tag string, value any) error {

	result, err := file.SearchSheet(SheetName, tag)
	if err != nil {
		fmt.Println(err)
		return err
	}

	if len(result) == 0 {
		fmt.Printf("Tag %s wasn't found in %s", tag, ExcelTemplateFilepath)
		err = errors.New("Tag " + tag + " wasn't found in " + ExcelTemplateFilepath)
		return err
	}

	err = file.SetCellValue(SheetName, result[0], value)
	if err != nil {
		fmt.Println(err)
		return err
	}

	return nil
}

// set report name in excel file
func setReportName(file *excelize.File, reportName string) error {

	err := setCellByTag(file, "#reportName", reportName)
	if err != nil {
		fmt.Println(err)
		return err
	}

	return nil
}

// set dates of the period in excel file
func setPeriod(file *excelize.File, from_timestamp int64, to_timestamp int64) error {

	err := setCellByTag(file, "#periodFrom", convertUnixTimestampToDateStr(from_timestamp))
	if err != nil {
		fmt.Println(err)
		return err
	}

	err = setCellByTag(file, "#periodTo", convertUnixTimestampToDateStr(to_timestamp))
	if err != nil {
		fmt.Println(err)
		return err
	}

	return nil
}

// set generation date in excel file.
// generationDate is an option, only one date is supported.
// By default:
//
//	generationDate = time.Now()
func setGenerationDate(file *excelize.File, generationDate ...time.Time) error {

	if len(generationDate) > 0 {
		err := setCellByTag(file, "#generationDate", generationDate[0].Format("02.01.2006"))
		if err != nil {
			fmt.Println(err)
			return err
		}

		return nil
	}

	err := setCellByTag(file, "#generationDate", time.Now().Format("02.01.2006"))
	if err != nil {
		fmt.Println(err)
		return err
	}

	return nil
}

// set total number of phone calls in excel file
func setTotalCalls(file *excelize.File, total int) error {

	err := setCellByTag(file, "#totalCalls", total)
	if err != nil {
		fmt.Println(err)
		return err
	}

	return nil
}

// calculate total talk time in seconds
func calcTotalTalkTime(callsData []PhoneCall) int {
	totalTalkTime := 0
	for i := range callsData {
		totalTalkTime += callsData[i].Talktime
	}

	return totalTalkTime
}

// calculate average talk time in seconds
func calcAvgTalkTime(callsData []PhoneCall) int {

	return int(math.Ceil(float64(calcTotalTalkTime(callsData) / len(callsData))))
}

// set total talk time in "hh ч mm мин" format in excel file
func setTotalTalkTime(file *excelize.File, seconds int) error {

	totalTalkTime := time.Duration(seconds) * time.Second
	totalTalkTimeStr := ""
	if totalTalkTime.Hours() < 1.0 {
		totalTalkTimeStr = "0 ч " + time.Time{}.Add(totalTalkTime).Format("4 мин")

	} else if totalTalkTime.Hours() < 12.0 {
		totalTalkTimeStr = time.Time{}.Add(totalTalkTime).Format("3 ч 4 мин")

	} else {
		totalTalkTimeStr = time.Time{}.Add(totalTalkTime).Format("15 ч 4 мин")
	}

	err := setCellByTag(file, "#totalTalkTime", totalTalkTimeStr)
	if err != nil {
		fmt.Println(err)
		return err
	}

	return nil
}

// set average talk time in "mm мин ss сек" format in excel file
func setAvgTalkTime(file *excelize.File, seconds int) error {

	avgTalkTime := time.Duration(seconds) * time.Second
	avgTalkTimeStr := time.Time{}.Add(avgTalkTime).Format("4 мин 05 сек")

	err := setCellByTag(file, "#avgTalkTime", avgTalkTimeStr)
	if err != nil {
		fmt.Println(err)
		return err
	}

	return nil
}

func formatTalkTimeToPrint(seconds int) string {

	// zeroTime := time.Time{}
	talkTime := time.Duration(seconds) * time.Second
	if talkTime.Hours() > 0 {
		return time.Time{}.Add(talkTime).Format("15:04:05")
	}
	return time.Time{}.Add(talkTime).Format("04:05")

	// if talkTime.Hours > 0 {
	// 	return fmt.Sprintf("%d:%02d:%02d", talkTime.Hours, talkTime.Minutes, talkTime.Seconds)
	// }
	// return fmt.Sprintf("%d:%02d", talkTime.Minutes, talkTime.Seconds)

}

// returns a cell in the next column in the same row
func nextCell(cell string) (string, error) {

	col, row, err := excelize.CellNameToCoordinates(cell)
	if err != nil {
		fmt.Println(err)
		return "", err
	}

	col++

	nextCell, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		fmt.Println(err)
		return "", err
	}

	return nextCell, nil
}

// returns a cell in the next row in the specified start column
func nextRow(cell string, startCol int) (string, error) {

	col, row, err := excelize.CellNameToCoordinates(cell)
	if err != nil {
		fmt.Println(err)
		return "", err
	}

	row++
	col = startCol

	nextRow, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		fmt.Println(err)
		return "", err
	}

	return nextRow, nil
}

// print cell with format (style) in excel file
func printCell(file *excelize.File, cell string, style int, value any) error {

	err := file.SetCellStyle(SheetName, cell, cell, style)
	if err != nil {
		fmt.Println(err)
		return err
	}

	err = file.SetCellValue(SheetName, cell, value)
	if err != nil {
		fmt.Println(err)
		return err
	}

	return nil
}

// print calls data in excel file
func printCallsData(file *excelize.File, callsData []PhoneCall) error {

	tag := "#callsTableStart"

	result, err := file.SearchSheet(SheetName, tag)
	if err != nil {
		fmt.Println(err)
		return err
	}

	if len(result) == 0 {
		fmt.Printf("Tag %s wasn't found in %s", tag, ExcelTemplateFilepath)
		err = errors.New("Tag " + tag + " wasn't found in " + ExcelTemplateFilepath)
		return err
	}

	startCell := result[0]

	startCol, _, err := excelize.CellNameToCoordinates(startCell)
	if err != nil {
		fmt.Println(err)
		return err
	}
	// fmt.Println("START CELL: ", startCol, startRow)

	// general format
	textStyle, err := file.NewStyle(&excelize.Style{
		NumFmt: 0,
		Alignment: &excelize.Alignment{
			Horizontal: "right",
		},
	})
	if err != nil {
		fmt.Println(err)
		return err
	}

	// [h]:mm:ss format
	timeStyle, err := file.NewStyle(&excelize.Style{
		NumFmt: 46,
		Alignment: &excelize.Alignment{
			Horizontal: "right",
		},
	})
	if err != nil {
		fmt.Println(err)
		return err
	}

	// dd.mm.yyyy hh:mm format
	dateStyle, err := file.NewStyle(&excelize.Style{NumFmt: 22})
	if err != nil {
		fmt.Println(err)
		return err
	}

	cell := startCell
	// col := startCol
	// row := startRow

	for i := range callsData {

		// "ID звонка"
		err = printCell(file, cell, textStyle, callsData[i].CallID)
		if err != nil {
			fmt.Println(err)
			return err
		}

		cell, err = nextCell(cell)
		if err != nil {
			fmt.Println(err)
			return err
		}

		// =-=-=-=-=-=-=

		// "От кого"
		err = printCell(file, cell, textStyle, callsData[i].From)
		if err != nil {
			fmt.Println(err)
			return err
		}

		cell, err = nextCell(cell)
		if err != nil {
			fmt.Println(err)
			return err
		}

		// =-=-=-=-=-=-=

		// "Кому"
		err = printCell(file, cell, textStyle, callsData[i].To)
		if err != nil {
			fmt.Println(err)
			return err
		}

		cell, err = nextCell(cell)
		if err != nil {
			fmt.Println(err)
			return err
		}

		// =-=-=-=-=-=-=

		// "Длит."
		err = printCell(file, cell, timeStyle, float64(callsData[i].Talktime)/86400.0)
		if err != nil {
			fmt.Println(err)
			return err
		}

		/*
			// you can use this block of code instead for time in string representation
			err = printCell(file, cell, textStyle, formatTalkTimeToPrint(callsData[i].Talktime))
			if err != nil {
				fmt.Println(err)
				return err
			}
		*/

		cell, err = nextCell(cell)
		if err != nil {
			fmt.Println(err)
			return err
		}

		// =-=-=-=-=-=-=

		// "Дата и время"
		err = printCell(file, cell, dateStyle, time.Unix(callsData[i].Timestamp, 0).Format("02.01.2006 15:04"))
		if err != nil {
			fmt.Println(err)
			return err
		}

		// =-=-=-=-=-=-=

		cell, err = nextRow(cell, startCol)
		if err != nil {
			fmt.Println(err)
			return err
		}

		/*
			for j := range reflect.TypeFor[PhoneCall]().NumField() {

				cell, err = excelize.CoordinatesToCellName(col, row)
				if err != nil {
					fmt.Println(err)
					return err
				}

				err = file.SetCellStyle(SheetName, cell, cell, textStyle)
				if err != nil {
					fmt.Println(err)
					return err
				}

				err = file.SetCellValue(SheetName, cell, callsData[i].CallID)
				if err != nil {
					fmt.Println(err)
					return err
				}

				switch j {
				case 0:
					err = file.SetCellStyle(SheetName, cell, cell, textStyle)
					if err != nil {
						fmt.Println(err)
						return
					}

					err = file.SetCellValue(SheetName, cell, callsData[i].CallID)
					if err != nil {
						fmt.Println(err)
						return
					}

				case 1:
					err = file.SetCellStyle(SheetName, cell, cell, textStyle)
					if err != nil {
						fmt.Println(err)
						return
					}

					err = file.SetCellValue(SheetName, cell, callsData[i].From)
					if err != nil {
						fmt.Println(err)
						return
					}

				case 2:
					err = file.SetCellStyle(SheetName, cell, cell, textStyle)
					if err != nil {
						fmt.Println(err)
						return
					}

					err = file.SetCellValue(SheetName, cell, callsData[i].To)
					if err != nil {
						fmt.Println(err)
						return
					}

				case 3:
					err = file.SetCellStyle(SheetName, cell, cell, timeStyle)
					if err != nil {
						fmt.Println(err)
						return
					}

					err = file.SetCellValue(SheetName, cell, float64(callsData[i].Talktime)/86400.0)
					if err != nil {
						fmt.Println(err)
						return
					}

				case 4:
					err = file.SetCellStyle(SheetName, cell, cell, dateStyle)
					if err != nil {
						fmt.Println(err)
						return
					}

					err = file.SetCellValue(SheetName, cell, time.Unix(callsData[i].Timestamp, 0).Format("02.01.2006 15:04"))
					if err != nil {
						fmt.Println(err)
						return
					}

				}

				col++
			}

			row++
			col = startCol
		*/

	}

	return nil

}

// print all required data in excel file
func printExcelFile(input_file *excelize.File, output_file string, callsData []PhoneCall) {

	setReportName(input_file, ReportName)
	setPeriod(input_file, callsData[0].Timestamp, callsData[len(callsData)-1].Timestamp)
	setGenerationDate(input_file)
	setTotalCalls(input_file, len(callsData))
	setTotalTalkTime(input_file, calcTotalTalkTime(callsData))
	setAvgTalkTime(input_file, calcAvgTalkTime(callsData))

	printCallsData(input_file, callsData)

	if err := input_file.SaveAs(output_file); err != nil {
		fmt.Println(err)
	}
}

// ===================

func main() {

	callsData, err := readCallsDataFile()
	if err != nil {
		fmt.Println(err)
		log.Fatal(err)
		return
	}

	f, err := excelize.OpenFile(ExcelTemplateFilepath)
	if err != nil {
		fmt.Println(err)
		log.Fatal(err)
		return
	}
	defer func() {
		// close the spreadsheet
		if err := f.Close(); err != nil {
			fmt.Println(err)
			log.Fatal(err)
		}
	}()

	printExcelFile(f, ExcelOutputFilepath, callsData)

	fmt.Println("Program finished. View results in output file ", ExcelOutputFilepath)

}
