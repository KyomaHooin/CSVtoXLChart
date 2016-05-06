#include-once

; #INDEX# =======================================================================================================================
; Title .........: ExcelChartConstants
; AutoIt Version : 3.3.12.0
; Language ......: English
; Description ...: Constants to be included in an AutoIt script when using the ExcelChart UDF.
; Author(s) .....: water, GreenCan
; Modified.......: 20150401 (YYYYMMDD)
; Contributors ..:
; Resources .....: Excel 2007 Developer Reference:  http://msdn.microsoft.com/en-us/library/bb149081(v=office.12).aspx
;                  Excel 2010 Developer Reference: http://msdn.microsoft.com/en-us/library/ff846392.aspx
;                  Excel 2013 Developer Reference: https://msdn.microsoft.com/EN-US/library/ff194068.aspx
; ===============================================================================================================================

; #CONSTANTS# ===================================================================================================================
; MsoGradientStyle Enumeration. Specifies the style for a gradient fill
; See: http://msdn.microsoft.com/en-us/library/aa432535(v=office.12).aspx
Global Const $msoGradientDiagonalDown = 4 ; Diagonal gradient moving from a top corner down to the opposite corner
Global Const $msoGradientDiagonalUp = 3 ; Diagonal gradient moving from a bottom corner up to the opposite corner
Global Const $msoGradientFromCenter = 7 ; Gradient running from the center out to the corners
Global Const $msoGradientFromCorner = 5 ; Gradient running from a corner to the other three corners
Global Const $msoGradientFromTitle = 6 ; Gradient running from the title outward
Global Const $msoGradientHorizontal = 1 ; Gradient running horizontally across the shape
Global Const $msoGradientMixed = -2 ; Gradient is mixed
Global Const $msoGradientVertical = 2 ; radient running vertically down the shape

; MsoLineDashStyle Enumeration. Dash style for a line
; See: http://msdn.microsoft.com/en-us/library/aa432639%28v=office.12%29.aspx
Global Const $msoLineDash = 4 ; Line consists of dashes only
Global Const $msoLineDashDot =5 ; Line is a dash-dot pattern
Global Const $msoLineDashDotDot = 6 ; Line is a dash-dot-dot pattern
Global Const $msoLineDashStyleMixed = -2 ; Not supported
Global Const $msoLineLongDash = 7 ; Line consists of long dashes
Global Const $msoLineLongDashDot =8 ; Line is a long dash-dot pattern
Global Const $msoLineRoundDot = 3 ; Line is made up of round dots
Global Const $msoLineSolid = 1 ; Line is solid
Global Const $msoLineSquareDot = 2 ; Line is made up of square dots

; MsoLineStyle Enumeration. Represents the style of a line
; See: http://msdn.microsoft.com/en-us/library/aa432640%28v=office.12%29.aspx
Global Const $msoLineSingle = 1 ; Single line
Global Const $msoLineStyleMixed = -2 ; Not supported
Global Const $msoLineThickBetweenThin = 5 ; Thick line with a thin line on each side
Global Const $msoLineThickThin = 4 ; Thick line next to thin line. For horizontal lines, thick line is above thin line. For vertical lines, thick line is to the left of the thin line
Global Const $msoLineThinThick = 3 ; Thick line next to thin line. For horizontal lines, thick line is below thin line. For vertical lines, thick line is to the right of the thin line
Global Const $msoLineThinThin = 2 ; Two thin lines

; MsoPresetTexture Enumeration. Specifies texture to be used to fill a shape
; See: http://msdn.microsoft.com/en-us/library/office/aa432664(v=office.12).aspx
Global Const $msoPresetTextureMixed = -2 ; Not used
Global Const $msoTextureBlueTissuePaper = 17 ; Blue tissue paper texture
Global Const $msoTextureBouquet = 20 ; Bouquet texture
Global Const $msoTextureBrownMarble = 11 ; Brown marble texture
Global Const $msoTextureCanvas = 2 ; Canvas texture
Global Const $msoTextureCork = 21 ; Cork texture
Global Const $msoTextureDenim = 3 ; Denim texture
Global Const $msoTextureFishFossil = 7 ; Fish fossil texture
Global Const $msoTextureGranite = 12 ; Granite texture
Global Const $msoTextureGreenMarble = 9 ; Green marble texture
Global Const $msoTextureMediumWood = 24 ; Medium wood texture
Global Const $msoTextureNewsprint = 13 ; Newsprint texture
Global Const $msoTextureOak = 23 ; Oak texture
Global Const $msoTexturePaperBag = 6 ; Paper bag texture
Global Const $msoTexturePapyrus = 1 ; Papyrus texture
Global Const $msoTextureParchment = 15 ; Parchment texture
Global Const $msoTexturePinkTissuePaper = 18 ; Pink tissue paper texture
Global Const $msoTexturePurpleMesh = 19 ; Purple mesh texture
Global Const $msoTextureRecycledPaper = 14 ; Recycled paper texture
Global Const $msoTextureSand = 8 ; Sand texture
Global Const $msoTextureStationery = 16 ; Stationery texture
Global Const $msoTextureWalnut = 22 ; Walnut texture
Global Const $msoTextureWaterDroplets = 5 ; Water droplets texture
Global Const $msoTextureWhiteMarble = 10 ; White marble texture
Global Const $msoTextureWovenMat = 4 ; Woven mat texture

; MsoShadowStyle Enumeration. Specifies the type of shadowing effect
; See: http://msdn.microsoft.com/en-us/library/aa432675(v=office.12).aspx
Global Const $msoShadowStyleInnerShadow = 1 ; Specifies the inner shadow effect
Global Const $msoShadowStyleMixed = -2 ; Specifies a combination of inner and outer shadow effects
Global Const $msoShadowStyleOuterShadow = 2 ; Specifies the outer shadow effect

; MsoThemeColorIndex Enumeration. Indicates the Office theme color
; See: http://msdn.microsoft.com/en-us/library/aa432702(v=office.12).aspx
Global Const $msoNotThemeColor = 0 ; No theme color
Global Const $msoThemeColorAccent1 = 5 ; Accent 1 theme color
Global Const $msoThemeColorAccent2 = 6 ; Accent 2 theme color
Global Const $msoThemeColorAccent3 = 7 ; Accent 3 theme color
Global Const $msoThemeColorAccent4 = 8 ; Accent 4 theme color
Global Const $msoThemeColorAccent5 = 9 ; Accent 5 theme color
Global Const $msoThemeColorAccent6 = 10 ; Accent 6 theme color
Global Const $msoThemeColorBackground1 = 14 ; Background 1 theme color
Global Const $msoThemeColorBackground2 = 16 ; Background 2 theme color
Global Const $msoThemeColorDark1 = 1 ; Dark 1 theme color
Global Const $msoThemeColorDark2 = 3 ; Dark 2 theme color
Global Const $msoThemeColorFollowedHyperlink = 12 ; Theme color for a clicked hyperlink
Global Const $msoThemeColorHyperlink = 11 ; Theme color for a hyperlink
Global Const $msoThemeColorLight1 = 2 ; Light 1 theme color
Global Const $msoThemeColorLight2 = 4 ; Light 2 theme color
Global Const $msoThemeColorMixed = -2 ; Mixed color theme
Global Const $msoThemeColorText1 = 13 ; Text 1 theme color
Global Const $msoThemeColorText2 = 15 ; Text 2 theme color

; XlAxisGroup Enumeration. Specifies the type of axis group
; See: http://msdn.microsoft.com/en-us/library/bb240962%28v=office.12%29.aspx
Global Const $xlPrimary = 1 ; Primary axis group
Global Const $xlSecondary = 2 ; Secondary axis group

; XlAxisType Enumeration. Represents the axis type
; See: http://msdn.microsoft.com/en-us/library/bb240966%28v=office.12%29.aspx
Global Const $xlCategory = 1 ; Axis displays categories
Global Const $xlSeriesAxis = 3 ; Axis displays data series
Global Const $xlValue = 2 ; Axis displays values

; XlBackground Enumeration. Specifies the background type for text
; See: http://msdn.microsoft.com/en-us/library/bb240967%28v=office.12%29.aspx
Global Const $xlBackgroundAutomatic = -4105 ; Excel controls the background
Global Const $xlBackgroundOpaque = 3 ; Opaque background
Global Const $xlBackgroundTransparent = 2 ; Transparent background

; XlBorderWeight Enumeration. Specifies the weight of the border around a range
; See: https://msdn.microsoft.com/EN-US/library/ff197515.aspx
Global Const $xlHairline = 1 ; Hairline (thinnest border)
Global Const $xlMedium = -4138 ; Medium
Global Const $xlThick = 4 ; Thick (widest border)
Global Const $xlThin = 2 ; Thin

; XlChartLocation Enumeration. Specifies where to relocate a chart
; See: http://msdn.microsoft.com/en-us/library/bb240998%28v=office.12%29.aspx
Global Const $xlLocationAsNewSheet = 1 ; Chart is moved to a new sheet
Global Const $xlLocationAsObject = 2 ; Chart is to be embedded in an existing sheet
Global Const $xlLocationAutomatic = 3 ; Excel controls chart location

; XlChartSplitType Enumeration. Specifies the values displayed in the second chart in a pie chart or a bar of pie chart
; See: http://msdn.microsoft.com/en-us/library/bb241007%28v=office.12%29.aspx
Global Const $xlSplitByCustomSplit = 4 ; Arbitrary slides are displayed in the second chart
Global Const $xlSplitByPercentValue = 3 ; Second chart displays values less than some percentage of the total value. The percentage is specified by the SplitValue property
Global Const $xlSplitByPosition = 1 ; Second chart displays the smallest values in the data series. The number of values to display is specified by the SplitValue property
Global Const $xlSplitByValue = 2 ; Second chart displays values less than the value specified by the SplitValue property

; XlChartType Enumeration. Represents the different chart types
; See: http://msdn.microsoft.com/en-us/library/bb241008%28v=office.12%29.aspx
Global Const $xl3DArea = -4098 ; 3D Area
Global Const $xl3DAreaStacked = 78 ; 3D Stacked Area
Global Const $xl3DAreaStacked100 = 79 ; 100% Stacked Area
Global Const $xl3DBarClustered = 60 ; 3D Clustered Bar
Global Const $xl3DBarStacked = 61 ; 3D Stacked Bar
Global Const $xl3DBarStacked100 = 62 ; 3D 100% Stacked Bar
Global Const $xl3DColumn = -4100 ; 3D Column
Global Const $xl3DColumnClustered = 54 ; 3D Clustered Column
Global Const $xl3DColumnStacked = 55 ; 3D Stacked Column
Global Const $xl3DColumnStacked100 = 56 ; 3D 100% Stacked Column
Global Const $xl3DLine = -4101 ; 3D Line
Global Const $xl3DPie = -4102 ; 3D Pie
Global Const $xl3DPieExploded = 70 ; Exploded 3D Pie
Global Const $xlArea = 1 ; Area
Global Const $xlAreaStacked = 76 ; Stacked Area
Global Const $xlAreaStacked100 = 77 ; 100% Stacked Area
Global Const $xlBarClustered = 57 ; Clustered Bar
Global Const $xlBarOfPie = 71 ; Bar of Pie
Global Const $xlBarStacked = 58 ; Stacked Bar
Global Const $xlBarStacked100 = 59 ; 100% Stacked Bar
Global Const $xlBubble = 15 ; Bubble
Global Const $xlBubble3DEffect = 87 ; Bubble with 3D effects
Global Const $xlColumnClustered = 51 ; Clustered Column
Global Const $xlColumnStacked = 52 ; Stacked Column
Global Const $xlColumnStacked100 = 53 ; 100% Stacked Column
Global Const $xlConeBarClustered = 102 ; Clustered Cone Bar
Global Const $xlConeBarStacked = 103 ; Stacked Cone Bar
Global Const $xlConeBarStacked100 = 104 ; 100% Stacked Cone Bar
Global Const $xlConeCol = 105 ; 3D Cone Column
Global Const $xlConeColClustered = 99 ; Clustered Cone Column
Global Const $xlConeColStacked = 100 ; Stacked Cone Column
Global Const $xlConeColStacked100 = 101 ; 100% Stacked Cone Column
Global Const $xlCylinderBarClustered = 95 ; Clustered Cylinder Bar
Global Const $xlCylinderBarStacked = 96 ; Stacked Cylinder Bar
Global Const $xlCylinderBarStacked100 = 97 ; 100% Stacked Cylinder Bar
Global Const $xlCylinderCol = 98 ; 3D Cylinder Column
Global Const $xlCylinderColClustered = 92 ; Clustered Cone Column
Global Const $xlCylinderColStacked = 93 ; Stacked Cone Column
Global Const $xlCylinderColStacked100 = 94 ; 100% Stacked Cylinder Column
Global Const $xlDoughnut = -4120 ; Doughnut
Global Const $xlDoughnutExploded = 80 ; Exploded Doughnut
Global Const $xlLine = 4 ; Line
Global Const $xlLineMarkers = 65 ; Line with Markers
Global Const $xlLineMarkersStacked = 66 ; Stacked Line with Markers
Global Const $xlLineMarkersStacked100 = 67 ; 100% Stacked Line with Markers
Global Const $xlLineStacked = 63 ; Stacked Line
Global Const $xlLineStacked100 = 64 ; 100% Stacked Line
Global Const $xlPie = 5 ; Pie
Global Const $xlPieExploded = 69 ; Exploded Pie
Global Const $xlPieOfPie = 68 ; Pie of Pie
Global Const $xlPyramidBarClustered = 109 ; Clustered Pyramid Bar
Global Const $xlPyramidBarStacked = 110 ; Stacked Pyramid Bar
Global Const $xlPyramidBarStacked100 = 111 ; 100% Stacked Pyramid Bar
Global Const $xlPyramidCol = 112 ; 3D Pyramid Column
Global Const $xlPyramidColClustered = 106 ; Clustered Pyramid Column
Global Const $xlPyramidColStacked = 107 ; Stacked Pyramid Column
Global Const $xlPyramidColStacked100 = 108 ; 100% Stacked Pyramid Column
Global Const $xlRadar = -4151 ; Radar
Global Const $xlRadarFilled = 82 ; Filled Radar
Global Const $xlRadarMarkers = 81 ; Radar with Data Markers
Global Const $xlStockHLC = 88 ; High-Low-Close
Global Const $xlStockOHLC = 89 ; Open-High-Low-Close
Global Const $xlStockVHLC = 90 ; Volume-High-Low-Close
Global Const $xlStockVOHLC = 91 ; Volume-Open-High-Low-Close
Global Const $xlSurface = 83 ; 3D Surface
Global Const $xlSurfaceTopView = 85 ; Surface (Top View)
Global Const $xlSurfaceTopViewWireframe = 86 ; Surface (Top View wireframe)
Global Const $xlSurfaceWireframe = 84 ; 3D Surface (wireframe)
Global Const $xlXYScatter = -4169 ; Scatter
Global Const $xlXYScatterLines = 74 ; Scatter with Lines
Global Const $xlXYScatterLinesNoMarkers = 75 ; Scatter with Lines and No Data Markers
Global Const $xlXYScatterSmooth = 72 ; Scatter with Smoothed Lines
Global Const $xlXYScatterSmoothNoMarkers = 73 ; Scatter with Smoothed Lines and No Data Markers

; XlConstants Enumeration. Global constants used in Microsoft Excel
; See: http://msdn.microsoft.com/en-us/site/ff585154
; Global Const $xlAutomatic = -4105 ; <== Already defined in Excel.au3
Global Const $xlCombination = -4111
Global Const $xlCustom = -4114
Global Const $xlBar = 2
Global Const $xlColumn = 3
Global Const $xl3DBar = -4099
Global Const $xl3DSurface = -4103
Global Const $xlDefaultAutoFormat = -1
; Global Const $xlNone = -4142 ; <== Already defined in Excel.au3
Global Const $xlAbove = 0
Global Const $xlBelow = 1
Global Const $xlBoth = 1
; Global Const $xlBottom = -4017 ; <== Already defined in Excel.au3
; Global Const $xlCenter = -4108 ; <== Already defined in Excel.au3
Global Const $xlChecker = 9
Global Const $xlCircle = 8
Global Const $xlCorner = 2
Global Const $xlCrissCross = 16
Global Const $xlCross = 4
Global Const $xlDiamond = 2
Global Const $xlDistributed = -4117
Global Const $xlFill = 5
Global Const $xlFixedValue = 1
Global Const $xlGeneral = 1
Global Const $xlGray16 = 17
Global Const $xlGray25 = -4124
Global Const $xlGray50 = -4125
Global Const $xlGray75 = -4126
Global Const $xlGray8 = 18
Global Const $xlGrid = 15
Global Const $xlHigh = -4127
Global Const $xlInside = 2
Global Const $xlJustify = -4130
; Global Const $xlLeft = -4131 ; <== Already defined in Excel.au3
Global Const $xlLightDown = 13
Global Const $xlLightHorizontal = 11
Global Const $xlLightUp = 14
Global Const $xlLightVertical = 12
Global Const $xlLow = -4134
Global Const $xlMaximum = 2
Global Const $xlMinimum = 4
Global Const $xlMinusValues = 3
Global Const $xlNextToAxis = 4
Global Const $xlOpaque = 3
Global Const $xlOutside = 3
Global Const $xlPercent = 2
Global Const $xlPlus = 9
Global Const $xlPlusValues = 2
; Global Const $xlRight = -4152 ; <== Already defined in Excel.au3
Global Const $xlScale = 3
Global Const $xlSemiGray75 = 10
Global Const $xlShowLabel = 4
Global Const $xlShowLabelAndPercent = 5
Global Const $xlShowPercent = 3
Global Const $xlShowValue = 2
Global Const $xlSingle = 2
Global Const $xlSolid = 1
Global Const $xlSquare = 1
Global Const $xlStar = 5
Global Const $xlStError = 4
; Global Const $xlTop = -4160 ; <== Already defined in Excel.au3
Global Const $xlTransparent = 2
Global Const $xlTriangle = 3

; XlDataLabelPosition Enumeration. Specifies where the data label is positioned
; See: http://msdn.microsoft.com/en-us/library/bb241064%28v=office.12%29.aspx
Global Const $xlLabelPositionAbove = 0 ; Data label is positioned above the data point
Global Const $xlLabelPositionBelow = 1 ; Data label is positioned below the data point
Global Const $xlLabelPositionBestFit = 5 ; Microsoft Office Excel 2007 sets the position of the data label
Global Const $xlLabelPositionCenter = -4108 ; Data label is centered on the data point or is inside a bar or pie chart
Global Const $xlLabelPositionCustom = 7 ; Data label is in a custom position
Global Const $xlLabelPositionInsideBase = 4 ; Data label is positioned inside the data point at the bottom edge
Global Const $xlLabelPositionInsideEnd = 3 ; Data label is positioned inside the data point at the top edge
Global Const $xlLabelPositionLeft = -4131 ; Data label is positioned to the left of the data point
Global Const $xlLabelPositionMixed = 6 ; Data labels are in multiple positions
Global Const $xlLabelPositionOutsideEnd = 2 ; Data label is positioned outside the data point at the top edge
Global Const $xlLabelPositionRight = -4152 ; Data label is positioned to the right of the data point

; XlDataLabelsType Enumeration. Specifies the type of data label to apply
; See: http://msdn.microsoft.com/en-us/library/bb241068%28v=office.12%29.aspx
Global Const $xlDataLabelsShowBubbleSizes = 6 ; Show the size of the bubble in reference to the absolute value
Global Const $xlDataLabelsShowLabel = 4 ; Category for the point
Global Const $xlDataLabelsShowLabelAndPercent = 5 ; Percentage of the total, and category for the point. Available only for pie charts and doughnut charts
Global Const $xlDataLabelsShowNone = -4142 ; No data labels
Global Const $xlDataLabelsShowPercent = 3 ; Percentage of the total. Available only for pie charts and doughnut charts
Global Const $xlDataLabelsShowValue = 2 ; Default value for the point (assumed if this argument is not specified)

; XlDisplayBlanksAs Enumeration. Specifies how blank cells are plotted on a chart
; See: http://msdn.microsoft.com/en-us/library/bb241214%28v=office.12%29.aspx
Global Const $xlInterpolated = 3 ; Values are interpolated into the chart
Global Const $xlNotPlotted = 1 ; Blank cells are not plotted
Global Const $xlZero = 2 ; Blanks are plotted as zero

; XlDisplayUnit Enumeration. Specifies the display unit label for an axis
; See: http://msdn.microsoft.com/en-us/library/bb241219%28v=office.12%29.aspx
Global Const $xlHundredMillions = -8 ; Hundreds of millions
Global Const $xlHundreds = -2 ; Hundreds
Global Const $xlHundredThousands = -5 ; Hundreds of thousands
Global Const $xlMillionMillions = -10 ; Millions of millions
Global Const $xlMillions = -6 ; Millions
Global Const $xlTenMillions = -7 ; Tens of millions
Global Const $xlTenThousands = -4 ; Tens of thousands
Global Const $xlThousandMillions = -9 ; Thousands of millions
Global Const $xlThousands = -3 ; Thousands

; XlEndStyleCap Enumeration. Specifies the end style for error bars.
; See: http://msdn.microsoft.com/en-us/library/bb241257%28v=office.12%29.aspx
Global Const $xlCap = 1 ; Caps applied
Global Const $xlNoCap = 2 ; No caps applied

; XlErrorBarDirection Enumeration. Specifies which axis values are to receive error bars
Global Const $xlX = -4168 ; Bars run parallel to the Y axis for X-axis values
Global Const $xlY = 1 ; Bars run parallel to the X axis for Y-axis values

; XlErrorBarInclude Enumeration. Specifies which error-bar parts to include.
; See: http://msdn.microsoft.com/en-us/library/bb241264%28v=office.12%29.aspx
Global Const $xlErrorBarIncludeBoth = 1 ; Both positive and negative error range
Global Const $xlErrorBarIncludeMinusValues = 3 ; Only negative error range
Global Const $xlErrorBarIncludeNone = -4142 ; No error bar range
Global Const $xlErrorBarIncludePlusValues = 2 ; Only positive error range

; XlErrorBarType Enumeration. Specifies the range marked by error bars.
; See: http://msdn.microsoft.com/en-us/library/bb241269%28v=office.12%29.aspx
Global Const $xlErrorBarTypeCustom = -4114 ; Range is set by fixed values or cell values
Global Const $xlErrorBarTypeFixedValue = 1 ; Fixed-length error bars
Global Const $xlErrorBarTypePercent = 2 ; Percentage of range to be covered by the error bars
Global Const $xlErrorBarTypeStDev = -4155 ; Shows range for specified number of standard deviations
Global Const $xlErrorBarTypeStError = 4 ; Shows standard error range

; XlFixedFormatType Enumeration. Specifies the type of file format
; See: http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.xlfixedformattype.aspx
;~ Global Const $xlTypePDF = 0 ; "PDF" — Portable Document Format file (.pdf)
;~ Global Const $xlTypeXPS = 1 ; "XPS" — XPS Document (.xps)

; XlHAlign Enumeration. Specifies the horizontal alignment of an object
; See: http://msdn.microsoft.com/en-us/library/bb241313%28v=office.12%29.aspx
Global Const $xlHAlignCenter = -4108 ; Center
Global Const $xlHAlignCenterAcrossSelection = 7 ; Center across selection
Global Const $xlHAlignDistributed = -4117 ; Distribute
Global Const $xlHAlignFill = 5 ; Fill
Global Const $xlHAlignGeneral = 1 ; Align according to data type
Global Const $xlHAlignJustify = -4130 ; Justify
Global Const $xlHAlignLeft = -4131 ; Left
Global Const $xlHAlignRight = -4152 ; Right

; XlLegendPosition Enumeration. Specifies the position of the legend on a chart
; See: http://msdn.microsoft.com/en-us/library/bb241345%28v=office.12%29.aspx
Global Const $xlLegendPositionBottom = -4107 ; Below the chart
Global Const $xlLegendPositionCorner = 2 ; In the upper right-hand corner of the chart border
Global Const $xlLegendPositionCustom = -4161 ; A custom position
; Global Const $xlLegendPositionLeft = -4131 ; Left of the chart <== Already defined in Excel.au3
Global Const $xlLegendPositionRight = -4152 ; Right of the chart
Global Const $xlLegendPositionTop = -4160 ; Above the chart

; XlMarkerStyle Enumeration. Specifies the marker style for a point or series in a line chart, scatter chart, or radar chart
; See: http://msdn.microsoft.com/en-us/library/bb241374%28v=office.12%29.aspx
Global Const $xlMarkerStyleAutomatic = -4105 ; Automatic markers
Global Const $xlMarkerStyleCircle = 8 ; Circular markers
Global Const $xlMarkerStyleDash = -4115 ; Long bar markers
Global Const $xlMarkerStyleDiamond = 2 ; Diamond-shaped markers
Global Const $xlMarkerStyleDot = -4118 ; Short bar markers
Global Const $xlMarkerStyleNone = -4142 ; No markers
Global Const $xlMarkerStylePicture = -4147 ; Picture markers
Global Const $xlMarkerStylePlus = 9 ; Square markers with a plus sign
Global Const $xlMarkerStyleSquare = 1 ; Square markers
Global Const $xlMarkerStyleStar = 5 ; Square markers with an asterisk
Global Const $xlMarkerStyleTriangle = 3 ; Triangular markers
Global Const $xlMarkerStyleX = -4168 ; Square markers with an X

; XlPageOrientation Enumeration. Specifies the page orientation when the worksheet is printed
; See: http://msdn.microsoft.com/en-us/library/bb241397%28v=office.12%29.aspx
Global Const $xlLandscape = 2 ; Landscape mode
Global Const $xlPortrait = 1 ; Portrait mode

; XlPaperSize Enumeration. Specifies the size of the paper
; See: http://msdn.microsoft.com/en-us/library/bb241398%28v=office.12%29.aspx
Global Const $xlPaper10x14 = 16 ; 10 in. x 14 in.
Global Const $xlPaper11x17 = 17 ; 11 in. x 17 in.
Global Const $xlPaperA3 = 8 ; A3 (297 mm x 420 mm)
Global Const $xlPaperA4 = 9 ; A4 (210 mm x 297 mm)
Global Const $xlPaperA4Small =10 ; A4 Small (210 mm x 297 mm)
Global Const $xlPaperA5 = 11 ; A5 (148 mm x 210 mm)
Global Const $xlPaperB4 = 12 ; B4 (250 mm x 354 mm)
Global Const $xlPaperB5 = 13 ; A5 (148 mm x 210 mm)
Global Const $xlPaperCsheet = 24 ; C size sheet
Global Const $xlPaperDsheet = 25 ; D size sheet
Global Const $xlPaperEnvelope10 = 20 ; Envelope #10 (4-1/8 in. x 9-1/2 in.)
Global Const $xlPaperEnvelope11 = 21 ; Envelope #11 (4-1/2 in. x 10-3/8 in.)
Global Const $xlPaperEnvelope12 = 22 ; Envelope #12 (4-1/2 in. x 11 in.)
Global Const $xlPaperEnvelope14 = 23 ; Envelope #14 (5 in. x 11-1/2 in.)
Global Const $xlPaperEnvelope9 = 19 ; Envelope #9 (3-7/8 in. x 8-7/8 in.)
Global Const $xlPaperEnvelopeB4 = 33 ; Envelope B4 (250 mm x 353 mm)
Global Const $xlPaperEnvelopeB5 = 34 ; Envelope B5 (176 mm x 250 mm)
Global Const $xlPaperEnvelopeB6 = 35 ; Envelope B6 (176 mm x 125 mm)
Global Const $xlPaperEnvelopeC3= 29 ; Envelope C3 (324 mm x 458 mm)
Global Const $xlPaperEnvelopeC4= 30 ; Envelope C4 (229 mm x 324 mm)
Global Const $xlPaperEnvelopeC5 = 28 ; Envelope C5 (162 mm x 229 mm)
Global Const $xlPaperEnvelopeC6 = 31 ; Envelope C6 (114 mm x 162 mm)
Global Const $xlPaperEnvelopeC65 = 32 ; Envelope C65 (114 mm x 229 mm)
Global Const $xlPaperEnvelopeDL =27 ; Envelope DL (110 mm x 220 mm)
Global Const $xlPaperEnvelopeItaly =36 ; Envelope (110 mm x 230 mm)
Global Const $xlPaperEnvelopeMonarch = 37 ; Envelope Monarch (3-7/8 in. x 7-1/2 in.)
Global Const $xlPaperEnvelopePersonal = 38 ; Envelope (3-5/8 in. x 6-1/2 in.)
Global Const $xlPaperEsheet = 26 ; E size sheet
Global Const $xlPaperExecutive = 7 ; Executive (7-1/2 in. x 10-1/2 in.)
Global Const $xlPaperFanfoldLegalGerman = 41 ; German Legal Fanfold (8-1/2 in. x 13 in.)
Global Const $xlPaperFanfoldStdGerman = 40 ; German Legal Fanfold (8-1/2 in. x 13 in.)
Global Const $xlPaperFanfoldUS = 39 ; U.S. Standard Fanfold (14-7/8 in. x 11 in.)
Global Const $xlPaperFolio = 14 ; Folio (8-1/2 in. x 13 in.)
Global Const $xlPaperLedger = 4 ; Ledger (17 in. x 11 in.)
Global Const $xlPaperLegal = 5 ; Legal (8-1/2 in. x 14 in.)
Global Const $xlPaperLetter = 1 ; Letter (8-1/2 in. x 11 in.)
Global Const $xlPaperLetterSmall = 2 ; Letter Small (8-1/2 in. x 11 in.)
Global Const $xlPaperNote = 18 ; Note (8-1/2 in. x 11 in.)
Global Const $xlPaperQuarto = 15 ; Quarto (215 mm x 275 mm)
Global Const $xlPaperStatement = 6 ; Statement (5-1/2 in. x 8-1/2 in.)
Global Const $xlPaperTabloid = 3 ; Tabloid (11 in. x 17 in.)
Global Const $xlPaperUser = 256 ; User-defined

; XlRowCol Enumeration. Specifies whether the values corresponding to a particular data series are in rows or columns
; See: http://msdn.microsoft.com/en-us/library/bb241571%28v=office.12%29.aspx
; Global Const $xlColumns = 2 ; Data series is in a row <== Already defined in Excel.au3
Global Const $xlRows = 1 ; Data series is in a column

; XlSizeRepresents Enumeration. Specifies what the bubble size represents on a bubble chart
; See: http://msdn.microsoft.com/en-us/library/bb241601%28v=office.12%29.aspx
Global Const $xlSizeIsArea = 1 ; Area of the bubble
Global Const $xlSizeIsWidth = 2 ; Width of the bubble

; XlTickLabelPosition Enumeration. Specifies the position of tick-mark labels on the specified axis
; See: http://msdn.microsoft.com/en-us/library/bb216368%28v=office.12%29.aspx
Global Const $xlTickLabelPositionHigh = -4127 ; Top or right side of the chart
Global Const $xlTickLabelPositionLow = -4134 ; Bottom or left side of the chart
Global Const $xlTickLabelPositionNextToAxis = 4 ; Next to axis (where axis is not at either side of the chart)
Global Const $xlTickLabelPositionNone = -4142 ; No tick marks

; XlTickMark Enumeration. Specifies the position of major and minor tick marks for an axis
; See: http://msdn.microsoft.com/en-us/library/bb216376%28v=office.12%29.aspx
Global Const $xlTickMarkCross = 4 ; Crosses the axis
Global Const $xlTickMarkInside = 2 ; Inside the axis
Global Const $xlTickMarkNone = -4142 ; No mark
Global Const $xlTickMarkOutside = 3 ; Outside the axis

; XlTrendlineType Enumeration. Specifies how the trendline that smoothes out fluctuations in the data is calculated
; See: http://msdn.microsoft.com/en-us/library/bb216402(v=office.12).aspx
Global Const $xlExponential = 5 ; Uses an equation to calculate the least squares fit through points (for example, y=ab^x)
; Global Const $xlLinear = -4132 ;  Uses the linear equation y = mx + b to calculate the least squares fit through points. <== Already declared in Excel.au3
Global Const $xlLogarithmic = -4133 ; Uses the equation y = c ln x + b to calculate the least squares fit through points
Global Const $xlMovingAvg = 6 ; Uses a sequence of averages computed from parts of the data series. The number of points equals the total number of points in the series minus the number specified for the period
Global Const $xlPolynomial  = 3 ; Uses an equation to calculate the least squares fit through points (for example, y = ax^6 + bx^5 + cx^4 + dx^3 + ex^2 + fx + g)
Global Const $xlPower = 4 ; Uses an equation to calculate the least squares fit through points (for example, y = ax^b)
; ===============================================================================================================================