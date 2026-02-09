package main

import (
	"encoding/json"
	"fmt"
	"log"
	"math"
	"os"
	"reflect"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

type PhoneCall struct {
	CallID    int    `json: "call_id"`   // call id
	From      string `json: "from"`      // user number 1
	To        string `json: "to"`        // user number 2
	Talktime  int    `json: "talktime"`  // seconds
	Timestamp int64  `json: "timestamp"` // UNIX seconds
}

type ExcelTime struct {
	Seconds int
	Minutes int
	Hours   int
}

// ===================

const ReportName string = "Отчет по звонкам"
const SheetName string = "Sheet1"

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

// set report name in excel file
func setReportName(file *excelize.File, reportName string) {
	result, err := file.SearchSheet(SheetName, "#reportName")
	if err != nil {
		fmt.Println(err)
		return
	}
	err = file.SetCellStr(SheetName, result[0], reportName)

	if err != nil {
		fmt.Println(err)
		return
	}
}

// set dates of the period in excel file
func setPeriod(file *excelize.File, from_timestamp int64, to_timestamp int64) {
	result, err := file.SearchSheet(SheetName, "#periodFrom")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = file.SetCellStr(SheetName, result[0], convertUnixTimestampToDateStr(from_timestamp))
	if err != nil {
		fmt.Println(err)
		return
	}

	result, err = file.SearchSheet(SheetName, "#periodTo")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = file.SetCellStr(SheetName, result[0], convertUnixTimestampToDateStr(to_timestamp))
	if err != nil {
		fmt.Println(err)
		return
	}
}

// set now as generation date in excel file
func setGenerationDate(file *excelize.File) {
	result, err := file.SearchSheet(SheetName, "#generationDate")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = file.SetCellStr(SheetName, result[0], time.Now().Format("02.01.2006"))
	if err != nil {
		fmt.Println(err)
		return
	}
}

// set generation date in excel file
func setGenerationDateManually(file *excelize.File, generationDate time.Time) {
	result, err := file.SearchSheet(SheetName, "#generationDate")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = file.SetCellStr(SheetName, result[0], generationDate.Format("02.01.2006"))
	if err != nil {
		fmt.Println(err)
		return
	}
}

// set total number of phone calls in excel file
func setTotalCalls(file *excelize.File, total int) {
	result, err := file.SearchSheet(SheetName, "#totalCalls")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = file.SetCellInt(SheetName, result[0], int64(total))
	if err != nil {
		fmt.Println(err)
		return
	}
}

// set total talk time in "hh ч mm мин" format in excel file
func setTotalTalkTime(file *excelize.File, seconds int) {
	result, err := file.SearchSheet(SheetName, "#totalTalkTime")
	if err != nil {
		fmt.Println(err)
		return
	}

	str := strconv.Itoa(convertSeconds(seconds, true).Hours) + " ч " + strconv.Itoa(convertSeconds(seconds, true).Minutes) + " мин"

	err = file.SetCellStr(SheetName, result[0], str)
	if err != nil {
		fmt.Println(err)
		return
	}
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

// set average talk time in "mm мин ss сек" format in excel file
func setAvgTalkTime(file *excelize.File, seconds int) {
	result, err := file.SearchSheet(SheetName, "#avgTalkTime")
	if err != nil {
		fmt.Println(err)
		return
	}

	str := strconv.Itoa(convertSeconds(seconds, false).Minutes) + " мин " + strconv.Itoa(convertSeconds(seconds, false).Seconds) + " сек"

	err = file.SetCellStr(SheetName, result[0], str)
	if err != nil {
		fmt.Println(err)
		return
	}
}

func formatTalkTimeToPrint(seconds int) string {

	talkTime := convertSeconds(seconds, true)
	if talkTime.Hours > 0 {
		return fmt.Sprintf("%d:%02d:%02d", talkTime.Hours, talkTime.Minutes, talkTime.Seconds)
	}
	return fmt.Sprintf("%d:%02d", talkTime.Minutes, talkTime.Seconds)

}

// print calls data in excel file
func printCallsData(file *excelize.File, callsData []PhoneCall) {

	result, err := file.SearchSheet(SheetName, "#callsTableStart")
	if err != nil {
		fmt.Println(err)
		return
	}

	startCell := result[0]

	startCol, startRow, err := excelize.CellNameToCoordinates(startCell)
	if err != nil {
		fmt.Println(err)
		return
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
		return
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
		return
	}

	// dd.mm.yyyy hh:mm format
	dateStyle, err := file.NewStyle(&excelize.Style{NumFmt: 22})
	if err != nil {
		fmt.Println(err)
		return
	}

	cell := startCell
	col := startCol
	row := startRow

	for i := range callsData {
		for j := range reflect.TypeFor[PhoneCall]().NumField() {

			cell, err = excelize.CoordinatesToCellName(col, row)
			if err != nil {
				fmt.Println(err)
				return
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

				/*
					// you can use this block of code instead case 3 block for time in string representation
					err = file.SetCellStyle(SheetName, cell, cell, textStyle)
					if err != nil {
						fmt.Println(err)
						return
					}

					err = file.SetCellValue(SheetName, cell, formatTalkTimeToPrint(callsData[i].Talktime))
					if err != nil {
						fmt.Println(err)
						return
					}
				*/

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

	}

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
