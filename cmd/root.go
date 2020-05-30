package cmd

import (
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/spf13/cobra"
)

var sep string

var rootCmd = &cobra.Command{
	Use:  "xlsx2text <FILE> <SHEET>",
	Args: cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		fileName, sheetName := args[0], args[1]

		f, err := excelize.OpenFile(fileName)
		if err != nil {
			fmt.Println(err)
			return
		}

		rows := f.GetRows(sheetName)
		for _, row := range rows {
			nCols := len(row)
			for i, colCell := range row {
				if nCols == i+1 {
					fmt.Print(colCell)
				} else {
					fmt.Print(colCell, sep)
				}
			}
			fmt.Println()
		}
	},
}

func Execute() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}

func init() {
	rootCmd.Flags().StringVarP(&sep, "sep", "s", "\t", "separator")
}
