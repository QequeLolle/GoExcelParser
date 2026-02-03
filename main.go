package main

import (
	"encoding/json"
	"fmt"
	"log"
	"math"
	"os"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type PhoneCall struct {
	Call_id   int    `json: "call_id"`   // call id
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

const REPORT_NAME string = "Отчет по звонкам"

var EXCEL_TEMPLATE_FILEPATH string = os.Args[1]
var JSON_FILEPATH string = os.Args[2]
var EXCEL_OUTPUT_FILEPATH string = os.Args[3]

// ===================

// read call data file
func readCallsFile() []PhoneCall {
	file, err := os.Open(JSON_FILEPATH)
	if err != nil {
		log.Fatal(err)
	}

	var calls []PhoneCall

	decoder := json.NewDecoder(file)
	if err = decoder.Decode(&calls); err != nil {
		log.Fatal(err)
	}

	if err = file.Close(); err != nil {
		log.Fatal(err)
	}

	return calls
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
	result, err := file.SearchSheet("Sheet1", "#reportName")
	err = file.SetCellStr("Sheet1", result[0], reportName)

	if err != nil {
		fmt.Println(err)
		return
	}
}

// set dates of the period in excel file
func setPeriod(file *excelize.File, from_timestamp int64, to_timestamp int64) {
	result, err := file.SearchSheet("Sheet1", "#periodFrom")
	err = file.SetCellStr("Sheet1", result[0], convertUnixTimestampToDateStr(from_timestamp))

	result, err = file.SearchSheet("Sheet1", "#periodTo")
	err = file.SetCellStr("Sheet1", result[0], convertUnixTimestampToDateStr(to_timestamp))

	if err != nil {
		fmt.Println(err)
		return
	}
}

// set now as generation date in excel file
func setGenerationDate(file *excelize.File) {
	result, err := file.SearchSheet("Sheet1", "#generationDate")
	err = file.SetCellStr("Sheet1", result[0], time.Now().Format("02.01.2006"))

	if err != nil {
		fmt.Println(err)
		return
	}
}

// set generation date in excel file
func setGenerationDateManually(file *excelize.File, generationDate time.Time) {
	result, err := file.SearchSheet("Sheet1", "#generationDate")
	err = file.SetCellStr("Sheet1", result[0], generationDate.Format("02.01.2006"))

	if err != nil {
		fmt.Println(err)
		return
	}
}

// set total number of phone calls in excel file
func setTotalCalls(file *excelize.File, total int) {
	result, err := file.SearchSheet("Sheet1", "#totalCalls")
	err = file.SetCellInt("Sheet1", result[0], int64(total))

	if err != nil {
		fmt.Println(err)
		return
	}
}

// set total talk time in "hh ч mm мин" format in excel file
func setTotalTalkTime(file *excelize.File, seconds int) {
	result, err := file.SearchSheet("Sheet1", "#totalTalkTime")
	str := strconv.Itoa(convertSeconds(seconds, true).Hours) + " ч " + strconv.Itoa(convertSeconds(seconds, true).Minutes) + " мин"
	err = file.SetCellStr("Sheet1", result[0], str)

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
	result, err := file.SearchSheet("Sheet1", "#avgTalkTime")
	str := strconv.Itoa(convertSeconds(seconds, false).Minutes) + " мин " + strconv.Itoa(convertSeconds(seconds, false).Seconds) + " сек"
	err = file.SetCellStr("Sheet1", result[0], str)

	if err != nil {
		fmt.Println(err)
		return
	}
}

func formatTalkTimeToPrint(seconds int) string {
	// if seconds < 3600 {
	// 	talkTime := convertSeconds(seconds, false)
	// 	return string(strconv.Itoa(talkTime.Minutes) + ":" + strconv.Itoa(talkTime.Seconds))
	// } else {
	// 	talkTime := convertSeconds(seconds, true)
	// 	return string(strconv.Itoa(talkTime.Hours) + ":" + strconv.Itoa(talkTime.Minutes) + ":" + strconv.Itoa(talkTime.Seconds))
	// }

	talkTime := convertSeconds(seconds, true)
	return string(strconv.Itoa(talkTime.Hours) + ":" + strconv.Itoa(talkTime.Minutes) + ":" + strconv.Itoa(talkTime.Seconds))

}

// convert calls data into strings for printing in excel file
func prepareCallsDataToPrint(callsData []PhoneCall) [][]string {
	result := make([][]string, len(callsData))
	for i := range len(callsData) {
		result[i] = make([]string, 5)

		strTalkTime := strconv.FormatFloat(float64(float64(callsData[i].Talktime)/86400.0), 'f', 5, 64)
		// fmt.Println("STR TALKTIME: ", strTalkTime)
		strTalkTime = strings.ReplaceAll(strTalkTime, ".", ",")
		// fmt.Println("STR TALKTIME2: ", strTalkTime)

		result[i][0] = strconv.Itoa(callsData[i].Call_id)
		result[i][1] = callsData[i].From
		result[i][2] = callsData[i].To
		result[i][3] = strTalkTime                                                     // strconv.FormatFloat(float64(callsData[i].Talktime/86400), 'f', 6, 64) //formatTalkTimeToPrint(callsData[i].Talktime)
		result[i][4] = time.Unix(callsData[i].Timestamp, 0).Format("02.01.2006 15:04") //convertUnixTimestampToDateStr(callsData[i].Timestamp)

		// fmt.Printf("RESULT [%v]: call_id = %v, from = %v, to = %v, talktime = %v, datetime = %v\n",
		// 	i, result[i][0], result[i][1], result[i][2], result[i][3], result[i][4])

	}

	return result

}

// // print calls data in excel file
// func printCallsData(file *excelize.File, callsData []PhoneCall) {
// 	result, err := file.SearchSheet("Sheet1", "#callsTableStart")

// 	prepareadData := prepareCallsDataToPrint(callsData)

// 	// print first data row
// 	err = file.SetSheetRow("Sheet1", result[0], &prepareadData[0])

// 	if err != nil {
// 		fmt.Println(err)
// 		return
// 	}

// 	if len(callsData) == 1 {

// 		return

// 	} else {

// 		// extract column name and row number from cell name
// 		colStr, row, err := excelize.SplitCellName(result[0])

// 		if err != nil {
// 			fmt.Println(err)
// 			return
// 		}

// 		/*
// 			// extract column name from cell name
// 			colRegexp := regexp.MustCompile(`[A-Z]+`)
// 			colMatch := colRegexp.FindString(result[0])
// 			// fmt.Println("MATCH COLUMN STRING = ", colMatch)

// 			// extract row number from cell name
// 			rowRegexp := regexp.MustCompile(`\d+`)
// 			rowMatch := rowRegexp.FindString(result[0])
// 			// fmt.Println("MATCH ROW STRING = ", rowMatch)

// 			row, _ := strconv.ParseInt(rowMatch, 10, 0)
// 		*/

// 		for i := 1; i < len(callsData); i++ {

// 			// iterate to next row
// 			row++
// 			newCell, err := excelize.JoinCellName(colStr, row)
// 			// newCell := colMatch + strconv.Itoa(int(row))
// 			// fmt.Println("NEW CELL COORDS FUNC: ", newCell)

// 			if err != nil {
// 				fmt.Println(err)
// 				return
// 			}

// 			// print data row
// 			err = file.SetSheetRow("Sheet1", newCell, &prepareadData[i])

// 			if err != nil {
// 				fmt.Println(err)
// 				return
// 			}

// 		}
// 	}

// }

// print calls data in excel file
func printCallsData(file *excelize.File, callsData []PhoneCall) {

	result, err := file.SearchSheet("Sheet1", "#callsTableStart")

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
	fmt.Println("START CELL: ", startCol, startRow)

	// text format
	textStyle, err := file.NewStyle(&excelize.Style{NumFmt: 0})
	if err != nil {
		fmt.Println(err)
		return
	}

	// [h]:mm:ss format
	timeStyle, err := file.NewStyle(&excelize.Style{NumFmt: 46})
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
				err = file.SetCellStyle("Sheet1", cell, cell, textStyle)
				if err != nil {
					fmt.Println(err)
					return
				}

				err = file.SetCellValue("Sheet1", cell, callsData[i].Call_id)
				if err != nil {
					fmt.Println(err)
					return
				}

			case 1:
				err = file.SetCellStyle("Sheet1", cell, cell, textStyle)
				if err != nil {
					fmt.Println(err)
					return
				}

				err = file.SetCellValue("Sheet1", cell, callsData[i].From)
				if err != nil {
					fmt.Println(err)
					return
				}

			case 2:
				err = file.SetCellStyle("Sheet1", cell, cell, textStyle)
				if err != nil {
					fmt.Println(err)
					return
				}

				err = file.SetCellValue("Sheet1", cell, callsData[i].To)
				if err != nil {
					fmt.Println(err)
					return
				}

			case 3:
				err = file.SetCellStyle("Sheet1", cell, cell, timeStyle)
				if err != nil {
					fmt.Println(err)
					return
				}

				err = file.SetCellValue("Sheet1", cell, float64(callsData[i].Talktime)/86400.0)
				if err != nil {
					fmt.Println(err)
					return
				}

			case 4:
				err = file.SetCellStyle("Sheet1", cell, cell, dateStyle)
				if err != nil {
					fmt.Println(err)
					return
				}

				err = file.SetCellValue("Sheet1", cell, time.Unix(callsData[i].Timestamp, 0).Format("02.01.2006 15:04"))
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

// set reqired formating for calls data cells
func setStyleForCallsDataCells(file *excelize.File, rows int) {

	result, err := file.SearchSheet("Sheet1", "#callsTableStart")
	if err != nil {
		fmt.Println(err)
		return
	}

	col, row, err := excelize.CellNameToCoordinates(result[0])
	if err != nil {
		fmt.Println(err)
		return
	}

	last_row := row + rows - 1
	lastCell, err := excelize.CoordinatesToCellName(col, last_row)
	if err != nil {
		fmt.Println(err)
		return
	}

	// Number format
	intStyle, err := file.NewStyle(&excelize.Style{NumFmt: 1})
	if err != nil {
		fmt.Println(err)
		return
	}

	firstCell := result[0]

	// apply Number format to "ID звонка", "От кого", "Кому" columns
	for range 3 {
		lastCell, err = excelize.CoordinatesToCellName(col, last_row)
		if err != nil {
			fmt.Println(err)
			return
		}

		// apply format to the specified column
		err = file.SetCellStyle("Sheet1", firstCell, lastCell, intStyle)
		if err != nil {
			fmt.Println(err)
			return
		}

		// iterate to the next column
		col++
		firstCell, err = excelize.CoordinatesToCellName(col, row)
		if err != nil {
			fmt.Println(err)
			return
		}

	}

	// =-=-=-=-=-=-

	// [h]:mm:ss format
	timeStyle, err := file.NewStyle(&excelize.Style{NumFmt: 46})
	if err != nil {
		fmt.Println(err)
	}

	// first cell is already up to date
	lastCell, err = excelize.CoordinatesToCellName(col, last_row)
	if err != nil {
		fmt.Println(err)
		return
	}

	// apply [h]:mm:ss format to "Длит." column
	err = file.SetCellStyle("Sheet1", firstCell, lastCell, timeStyle)
	if err != nil {
		fmt.Println(err)
		return
	}

	// =-=-=-=-=-=-

	// dd.mm.yyyy hh:mm format
	dateStyle, err := file.NewStyle(&excelize.Style{NumFmt: 22})

	// iterate to the next column
	col++
	firstCell, err = excelize.CoordinatesToCellName(col, row)
	if err != nil {
		fmt.Println(err)
		return
	}

	lastCell, err = excelize.CoordinatesToCellName(col, last_row)
	if err != nil {
		fmt.Println(err)
		return
	}

	// apply dd.mm.yyyy hh:mm format to "Дата и время" column
	err = file.SetCellStyle("Sheet1", firstCell, lastCell, dateStyle)
	if err != nil {
		fmt.Println(err)
		return
	}
}

// ===================

func main() {

	calls := readCallsFile()

	f, err := excelize.OpenFile(EXCEL_TEMPLATE_FILEPATH)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// close the spreadsheet
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	setReportName(f, REPORT_NAME)
	setPeriod(f, calls[0].Timestamp, calls[len(calls)-1].Timestamp)
	setGenerationDate(f)
	setTotalCalls(f, len(calls))
	setTotalTalkTime(f, calcTotalTalkTime(calls))
	setAvgTalkTime(f, calcAvgTalkTime(calls))
	// prepareCallsDataToPrint(calls)
	// setStyleForCallsDataCells(f, len(calls))
	printCallsData(f, calls)

	// str, err := f.CalcCellValue("Sheet1", "D12")
	// fmt.Println("CALC: ", str)

	// // Используем Excel формат времени
	// excelTime := float64(45) / 86400.0

	// styleID, err := f.NewStyle(&excelize.Style{
	// 	NumFmt: 46, // [h]:mm:ss
	// 	Alignment: &excelize.Alignment{
	// 		Horizontal: "center",
	// 	},
	// })
	// if err != nil {
	// 	fmt.Println(err)
	// 	return
	// }

	// f.SetCellStyle("Sheet1", "D20", "D20", styleID)
	// f.SetCellValue("Sheet1", "D20", excelTime)

	if err := f.SaveAs(EXCEL_OUTPUT_FILEPATH); err != nil {
		fmt.Println(err)
	}

	fmt.Println("Program finished. View results in output file ", EXCEL_OUTPUT_FILEPATH)

}
