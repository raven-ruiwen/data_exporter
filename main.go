package main

import (
	"encoding/json"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"os"
	"os/exec"
	"strings"
	"time"
)

type Resp struct {
	Data struct {
		GetExecution struct {
			ExecutionQueued    interface{} `json:"execution_queued"`
			ExecutionRunning   interface{} `json:"execution_running"`
			ExecutionSucceeded struct {
				ExecutionId               string                   `json:"execution_id"`
				RuntimeSeconds            int                      `json:"runtime_seconds"`
				GeneratedAt               time.Time                `json:"generated_at"`
				Columns                   []string                 `json:"columns"`
				Data                      []map[string]interface{} `json:"data"`
				RequestMaxResultSizeBytes int                      `json:"request_max_result_size_bytes"`
			} `json:"execution_succeeded"`
			ExecutionFailed interface{} `json:"execution_failed"`
		} `json:"get_execution"`
	} `json:"data"`
}
type Address struct {
	Address                    string `json:"address"`
	IsScrollBridge             bool   `json:"is_scroll_bridge"`
	IsOrbiterBridge            bool   `json:"is_orbiter_bridge"`
	IsChaineyeBridge           bool   `json:"is_chaineye_bridge"`
	IsStakestoneDeposit        bool   `json:"is_stakestone_deposit"`
	IsExecuteAave2BorrowEth    bool   `json:"is_execute_aave_2_borrow_eth"`
	IsExecuteAave2BorrowWstEth bool   `json:"is_execute_aave_2_borrow_wst_eth"`
}

func main() {
	data, err := ioutil.ReadFile("curl.cmd")
	if err != nil {
		fmt.Println("can't read curl.cmd file, need create and fill command", err)
		return
	}

	curlCmd := string(data)
	curlCmd = strings.Replace(curlCmd, "curl '/v1/graphql'", "curl '/v1/graphql' -o resp.json", -1)

	cmd := exec.Command("bash", "-c", curlCmd)

	output, err := cmd.CombinedOutput()
	if err != nil {
		fmt.Println("exec curl.cmd err:", err, " check curl command")
		return
	}

	fmt.Println(string(output))
	file, err := os.Open("./resp.json")
	if err != nil {
		fmt.Println("Error opening file:", err)
		return
	}
	defer file.Close()

	decoder := json.NewDecoder(file)
	var respp Resp
	err = decoder.Decode(&respp)
	if err != nil {
		fmt.Println("Error decoding JSON:", err)
		return
	}

	f := excelize.NewFile()

	_, _ = f.NewSheet("Sheet1")

	headers := []string{}
	for _, column := range respp.Data.GetExecution.ExecutionSucceeded.Columns {
		headers = append(headers, column)
	}

	style, err := f.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Horizontal: "center"}})

	f.SetColStyle("Sheet1", "A:Q", style)
	hh := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	for col, header := range headers {
		colName, _ := excelize.ColumnNumberToName(col + 1)
		cell := colName + "1"
		f.SetCellValue("Sheet1", cell, header)
	}
	cellBeginAt := 2
	for _, data := range respp.Data.GetExecution.ExecutionSucceeded.Data {
		i := 0
		for _, _ = range data {
			h := hh[i]
			f.SetCellValue("Sheet1", fmt.Sprintf("%c%d", h, cellBeginAt), data[headers[i]])
			i++
		}

		cellBeginAt++
	}

	fileName := fmt.Sprintf("dune_%d.xlsx", time.Now().Unix())
	err = f.SaveAs(fileName)
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println("Excel file create success: ", fileName)
}
