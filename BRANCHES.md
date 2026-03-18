# Branch Strategy

## master

The `master` branch stays in sync with the upstream [excelize](https://github.com/xuri/excelize) repository. It contains the original, unmodified excelize source code and tracks upstream updates.

## plus-release

The `plus-release` branch contains custom extensions built on top of the upstream excelize codebase. These extensions include:

### GetCharts

`GetCharts(sheet, cell string) ([]*Chart, error)` - Read chart configurations from existing spreadsheets. Returns `[]*Chart` with chart type, series, title, legend position, and layout information.

```go
charts, err := f.GetCharts("Sheet1", "E1")
if err != nil {
    log.Fatal(err)
}
for _, chart := range charts {
    fmt.Println(chart.Type, chart.Title, chart.Legend.Position)
}
```

### Manual Layout for Charts

Added `ChartLayout` struct and layout support for chart titles, legends, and plot areas when creating charts with `AddChart`.

```go
err := f.AddChart("Sheet1", "E1", &excelize.Chart{
    Type:   excelize.Col,
    Series: series,
    Title:  []excelize.RichTextRun{{Text: "My Chart"}},
    TitleLayout: &excelize.ChartLayout{
        X: 0.1, Y: 0.05, Width: 0.8, Height: 0.1,
    },
    Legend: excelize.ChartLegend{
        Position: "right",
        Layout: &excelize.ChartLayout{
            X: 0.7, Y: 0.3, Width: 0.25, Height: 0.4,
        },
    },
    PlotArea: excelize.ChartPlotArea{
        Layout: &excelize.ChartLayout{
            X: 0.1, Y: 0.2, Width: 0.6, Height: 0.7,
        },
    },
})
```

### New Types

- `ChartLayout` - Layout configuration with `X`, `Y`, `Width`, `Height` (float64, 0.0-1.0 as fraction of chart area)

### Modified Types

- `Chart` - Added `TitleLayout *ChartLayout`
- `ChartLegend` - Added `Layout *ChartLayout`
- `ChartPlotArea` - Added `Layout *ChartLayout`
