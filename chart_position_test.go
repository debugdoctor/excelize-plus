package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestChartPositionRoundTrip writes charts with exact positions to an xlsx file,
// reads them back, and verifies the positions are preserved exactly.
func TestChartPositionRoundTrip(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	// Populate data
	for r, row := range [][]interface{}{
		{nil, "Q1", "Q2", "Q3", "Q4"},
		{"Product A", 100, 200, 150, 300},
		{"Product B", 80, 120, 200, 180},
		{"Product C", 150, 90, 110, 250},
	} {
		cell, _ := CoordinatesToCellName(1, r+1)
		assert.NoError(t, f.SetSheetRow(sheet, cell, &row))
	}

	// Chart 1: Column chart at exact position
	pos1 := &ChartPosition{
		FromCol: 5, FromColOff: 0,
		FromRow: 0, FromRowOff: 0,
		ToCol: 13, ToColOff: 0,
		ToRow: 15, ToRowOff: 0,
	}
	assert.NoError(t, f.AddChartWithPosition(sheet, "F1", &Chart{
		Type: Col,
		Series: []ChartSeries{
			{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$E$1", Values: "Sheet1!$B$2:$E$2"},
			{Name: "Sheet1!$A$3", Categories: "Sheet1!$B$1:$E$1", Values: "Sheet1!$B$3:$E$3"},
		},
		Title:    []RichTextRun{{Text: "Chart 1 - Column"}},
		Position: pos1,
	}))

	// Chart 2: Line chart at a different exact position with offsets
	pos2 := &ChartPosition{
		FromCol: 5, FromColOff: 50000,
		FromRow: 16, FromRowOff: 100000,
		ToCol: 13, ToColOff: 150000,
		ToRow: 32, ToRowOff: 50000,
	}
	assert.NoError(t, f.AddChartWithPosition(sheet, "F17", &Chart{
		Type: Line,
		Series: []ChartSeries{
			{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$E$1", Values: "Sheet1!$B$2:$E$2"},
			{Name: "Sheet1!$A$3", Categories: "Sheet1!$B$1:$E$1", Values: "Sheet1!$B$3:$E$3"},
			{Name: "Sheet1!$A$4", Categories: "Sheet1!$B$1:$E$1", Values: "Sheet1!$B$4:$E$4"},
		},
		Title:    []RichTextRun{{Text: "Chart 2 - Line"}},
		Position: pos2,
	}))

	// Chart 3: Pie chart using normal AddChart (no Position)
	assert.NoError(t, f.AddChart(sheet, "F34", &Chart{
		Type: Pie,
		Series: []ChartSeries{
			{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$E$1", Values: "Sheet1!$B$2:$E$2"},
		},
		Title: []RichTextRun{{Text: "Chart 3 - Pie (auto position)"}},
	}))

	// Save to file
	outPath := "test_chart_position.xlsx"
	assert.NoError(t, f.SaveAs(outPath))
	assert.NoError(t, f.Close())

	// ---- Read back and verify ----
	f2, err := OpenFile(outPath)
	assert.NoError(t, err)

	allCharts, err := f2.GetCharts(sheet, "")
	assert.NoError(t, err)
	assert.Len(t, allCharts, 3)

	fmt.Println("=== Chart Position Round-Trip Results ===")
	for i, c := range allCharts {
		fmt.Printf("\nChart %d: %s (type=%d)\n", i+1, c.Title[0].Text, c.Type)
		if c.Position != nil {
			fmt.Printf("  From: col=%d colOff=%d row=%d rowOff=%d\n",
				c.Position.FromCol, c.Position.FromColOff,
				c.Position.FromRow, c.Position.FromRowOff)
			fmt.Printf("  To:   col=%d colOff=%d row=%d rowOff=%d\n",
				c.Position.ToCol, c.Position.ToColOff,
				c.Position.ToRow, c.Position.ToRowOff)
		}
	}

	// Verify Chart 1 position exactly
	p1 := allCharts[0].Position
	assert.NotNil(t, p1)
	assert.Equal(t, pos1.FromCol, p1.FromCol)
	assert.Equal(t, pos1.FromColOff, p1.FromColOff)
	assert.Equal(t, pos1.FromRow, p1.FromRow)
	assert.Equal(t, pos1.FromRowOff, p1.FromRowOff)
	assert.Equal(t, pos1.ToCol, p1.ToCol)
	assert.Equal(t, pos1.ToColOff, p1.ToColOff)
	assert.Equal(t, pos1.ToRow, p1.ToRow)
	assert.Equal(t, pos1.ToRowOff, p1.ToRowOff)

	// Verify Chart 2 position exactly
	p2 := allCharts[1].Position
	assert.NotNil(t, p2)
	assert.Equal(t, pos2.FromCol, p2.FromCol)
	assert.Equal(t, pos2.FromColOff, p2.FromColOff)
	assert.Equal(t, pos2.FromRow, p2.FromRow)
	assert.Equal(t, pos2.FromRowOff, p2.FromRowOff)
	assert.Equal(t, pos2.ToCol, p2.ToCol)
	assert.Equal(t, pos2.ToColOff, p2.ToColOff)
	assert.Equal(t, pos2.ToRow, p2.ToRow)
	assert.Equal(t, pos2.ToRowOff, p2.ToRowOff)

	// Chart 3 should also have position (from auto-calc)
	assert.NotNil(t, allCharts[2].Position)

	assert.NoError(t, f2.Close())
	fmt.Println("\n=== All positions verified! ===")
}
