#Tidy_Parameters= /gd 1 /gds 1 /nsdp
#AutoIt3Wrapper_Au3Check_Parameters= -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_Au3Check_Stop_OnWarning=Y
#include-once
#include <ExcelChartConstants.au3>
#include <Excel.au3>

; #INDEX# =======================================================================================================================
; Title .........: Microsoft Excel Chart Function Library (MS Excel 2007 and later)
; AutoIt Version : 3.3.12.0
; UDF Version ...: 0.4.0.0
; Language ......: English
; Description ...: A collection of functions for creating and manipulating Microsoft Excel charts
; Author(s) .....: water, GreenCan
; Modified.......: 20150401 (YYYYMMDD)
; Remarks .......: Only works for Excel 2007 and later
; Contributors ..:
; Resources .....: Excel 2007 Developer Reference: http://msdn.microsoft.com/en-us/library/bb149081(v=office.12).aspx
;                  Excel 2010 Developer Reference: http://msdn.microsoft.com/en-us/library/ff846392.aspx
;                  Excel 2013 Developer Reference: https://msdn.microsoft.com/EN-US/library/ff194068.aspx
; ===============================================================================================================================

#region #VARIABLES#
; #VARIABLES# ===================================================================================================================
Global $g__iDebug = 0 ; Debug level. 0 = no debug information, 1 = to console, 2 = to MsgBox, 3 = into File
Global $g__sDebugFile = @ScriptDir & "\ExcelChartDebug.txt" ; Debug file if $g__iDebug is set to 3
Global $g__oError ; COM Error handler
; ===============================================================================================================================
#endregion #VARIABLES#

; #CURRENT# =====================================================================================================================
;_XLChart_3D_PositionSet
;_XLChart_AreaGroupSet
;_XLChart_AxisSet
;_XLChart_BarGroupSet
;_XLChart_BubbleGroupSet
;_XLChart_ChartCreate
;_XLChart_ChartDataSet
;_XLChart_ChartDelete
;_XLChart_ChartExport
;_XLChart_ChartPositionSet
;_XLChart_ChartPrint
;_XLChart_ChartSet
;_XLChart_ChartsGet
;_XLChart_ColumnGroupSet
;_XLChart_DatalabelSet
;_XLChart_DoughnutGroupSet
;_XLChart_ErrorBarSet
;_XLChart_FillSet
;_XLChart_FontSet
;_XLChart_GridSet
;_XLChart_LayoutSet
;_XLChart_LegendSet
;_XLChart_LineGet
;_XLChart_LineGroupSet
;_XLChart_LineSet
;_XLChart_MarkerSet
;_XLChart_ObjectDelete
;_XLChart_ObjectPositionSet
;_XLChart_OfPieGroupSet
;_XLChart_PageSet
;_XLChart_PieGroupSet
;_XLChart_ScreenUpdateSet
;_XLChart_SeriesAdd
;_XLChart_SeriesSet
;_XLChart_ShadowSet
;_XLChart_TicksSet
;_XLChart_TitleGet
;_XLChart_TitleSet
;_XLChart_TrendlineSet
;_XLChart_VersionInfo
;
; ===============================================================================================================================
; To be created
; ===============================================================================================================================
;	xGroupSet: Radar, XY
;	  HasRadarAxisLabels	Radar
;	  Has3DShading			Surface
; ===============================================================================================================================

; #INTERNAL_USE_ONLY#============================================================================================================
;_XLChart_Example
;_XLChart_COMError
;_XLChart_Version
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_3D_PositionSet
; Description....: Set the 3D position of the chart.
; Syntax.........: _XLChart_3D_PositionSet($oChart[, $iRotation = Default[, $iElevation = Default[, iPerspective = Default[, $iDepthPercent = Default[, $iHeightPercent = Default[, $iGapDepth = Default]]]]]])
; Parameters ....: $oChart         - Chart object as returned by a preceding call to _XLChart_ChartCreate
;                  $iRotation      - Optional: Rotation of the plot area around the z-axis, in degrees).
;                                    Value must be from 0 to 360, except for 3-D bar charts, where the value must be from 0 to 44.
;                  $iElevation     - Optional: Height at which you view the chart, in degrees.
;                                    Value must be between -90 and 90, except for 3-D bar charts, where it must be between 0 and 44.
;                  $iPerspective   - Optional: Perspective for the 3-D chart view.
;                                    Value must be between 0 and 100.
;                                    0 is a 2D chart
;                                    -1 sets the chart axes at right angles, independent of chart rotation or elevation. Applies only to 3-D line, column, and bar charts.
;                  $iDepthPercent  - Optional: Depth of a 3-D chart as a percentage of the chart width.
;                                    Value must be between 20 and 2000 percent
;                  $iHeightPercent - Optional: Height of a 3-D chart as a percentage of the chart width.
;                                    Value must be between 20 and 10000 percent
;                  $iGapDepth      - Optional: Distance between the data series in a 3-D chart as a percentage of the marker width.
;                                    Value must be between 0 and 500
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oChart is no object
;                  |2 - $iRotation is not a number or < 0 or > 360
;                  |3 - $iElevation is not a number or < -90 or > 90
;                  |4 - $iPerspective is not a number or < -1 or > 100
;                  |5 - $iDepthPercent is not a number or < 20 or > 2000
;                  |6 - $iHeightPercent is not a number or < 5 or > 500
;                  |7 - $iGapDepth is not a number or < 0 or > 500
;                  |8 - Error setting $iRotation. $oChart might not be 3D. See @extended for details
;                  |9 - Error setting $iElevation. $oChart might not be 3D. See @extended for details
;                  |10 - Error setting $iPerspective. $oChart might not be 3D. See @extended for details
;                  |11 - Error setting $iDepthPercent. $oChart might not be 3D. See @extended for details
;                  |12 - Error setting $iHeightPercent. $oChart might not be 3D. See @extended for details
;                  |13 - Error setting $iGapDepth. $oChart might not be 3D. See @extended for details
; Authors........: GreenCan
; Modified ......: water
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_3D_PositionSet($oChart, $iRotation = Default, $iElevation = Default, $iPerspective = Default, $iDepthPercent = Default, $iHeightPercent = Default, $iGapDepth = Default)

	If Not IsObj($oChart) Then Return SetError(1, 0, "")
	If $iRotation <> Default And (Not IsNumber($iRotation) Or $iRotation < 0 Or $iRotation > 360) Then Return SetError(2, 0, 0)
	If $iElevation <> Default And (Not IsNumber($iElevation) Or $iElevation < -90 Or $iElevation > 90) Then Return SetError(3, 0, 0)
	If $iPerspective <> Default And (Not IsNumber($iPerspective) Or $iPerspective < -1 Or $iPerspective > 100) Then Return SetError(4, 0, 0)
	If $iDepthPercent <> Default And (Not IsNumber($iDepthPercent) Or $iDepthPercent < 20 Or $iDepthPercent > 2000) Then Return SetError(5, 0, 0)
	If $iHeightPercent <> Default And (Not IsNumber($iHeightPercent) Or $iHeightPercent < 5 Or $iHeightPercent > 500) Then Return SetError(6, 0, 0)
	If $iGapDepth <> Default And (Not IsNumber($iGapDepth) Or $iGapDepth < 0 Or $iGapDepth > 500) Then Return SetError(7, 0, 0)
	If $iRotation <> Default Then
		$oChart.Rotation = $iRotation
		If @error Then Return SetError(8, @error, 0)
	EndIf
	If $iElevation <> Default Then
		$oChart.Elevation = $iElevation
		If @error Then Return SetError(9, @error, 0)
	EndIf
	If $iPerspective <> Default Then
		If $iPerspective = -1 Then
			$oChart.RightAngleAxes = True
			If @error Then Return SetError(10, @error, 0)
		Else
			$oChart.Perspective = $iPerspective * 2 ; documentation says value goes from 0 to 100 but true is 0 to 200
			If @error Then Return SetError(10, @error, 0)
		EndIf
	EndIf
	If $iDepthPercent <> Default Then
		$oChart.DepthPercent = $iDepthPercent
		If @error Then Return SetError(11, @error, 0)
	EndIf
	If $iHeightPercent <> Default Then
		$oChart.AutoScaling = False
		$oChart.HeightPercent = $iHeightPercent
		If @error Then Return SetError(12, @error, 0)
	EndIf
	If $iGapDepth <> Default Then
		$oChart.GapDepth = $iGapDepth
		If @error Then Return SetError(13, @error, 0)
	EndIf
	Return 1

EndFunc   ;==>_XLChart_3D_PositionSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_AreaGroupSet
; Description....: Set properties of an area chart group.
; Syntax.........: _XLChart_AreaGroupSet($oObject[, $bHasDropLines = Default])
; Parameters ....: $oObject       - Chart group for which the properties should be set
;                  $bHasDropLines - Optional: True if the area chart has drop lines (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is not an object
;                  |2 - $bHasDropLines is not boolean
; Authors........: water
; Modified ......:
; Remarks .......: A chart contains one or more chart groups, each chart group contains one or more series, and
;                  each series contains one or more points.
;+
;                  You can either pass an item of the ChartGroups collection or an item of the AreaGroups collection (a ChartGroup object)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_AreaGroupSet($oObject, $bHasDropLines = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $bHasDropLines <> Default And Not IsBool($bHasDropLines) Then Return SetError(2, 0, 0)
	If $bHasDropLines <> Default Then $oObject.HasDropLines = $bHasDropLines
	Return 1

EndFunc   ;==>_XLChart_AreaGroupSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_AxisSet
; Description....: Set the properties of the selected axis.
; Syntax.........: _XLChart_AxisSet($oObject[, $vMinimumScale = Default[, $vMaximumScale = Default[, $iDisplayUnit = Default[, $iDisplayUnitCustom = Default[, $vDisplayUnitLabel = Default]]]]])
; Parameters ....: $oObject            - Object of an axis to set the properties
;                  $vMinimumScale      - Optional: Sets the minimum value on the Y (value) or X axis (has to be of type value not category). Can be a numeric value or "Auto" (default = Default)
;                  $vMaximumScale      - Optional: Sets the maximum value on the Y (value) or X axis (has to be of type value not category). Can be a numeric value or "Auto" (default = Default)
;                  $iDisplayUnit       - Optional: Sets the unit label for the value axis. Can be any of the XlDisplayUnit enumeration, xlCustom or xlNone (default = Default)
;                  $iDisplayUnitCustom - Optional: If the value of $iDisplayUnit is $xlCustom, $DisplayUnitCustom sets the value of the displayed units.
;                  +                     The value must be from 0 through 10E307 (default = Default)
;                  $vDisplayUnitLabel  - Optional: True if the DisplayUnitLabel is displayed. Any non-boolean data will be used as the label caption (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is no object
;                  |2 - Invalid value for $iMinimumScale. Must be numeric or "Auto"
;                  |3 - Invalid value for $iMaximumScale. Must be numeric or "Auto"
;                  |4 - Parameter $DisplayUnit, $DisplayUnitCustom and $DisplayUnitLabel are only valid for the value axis
;                  |5 - Parameter $DisplayUnit must be $xlCustom if you want to set $DisplayUnitCustom
; Authors........: water
; Modified ......:
; Remarks .......: When using parameters $vMinimumScale and $vMaximumScale for the X axis the axis has to be a value type axis not a category type axis.
;                  This is true for XYscatter type charts. For all other chart types (e.g. line, column, area) the script will crash.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_AxisSet($oObject, $iMinimumScale = Default, $iMaximumScale = Default, $iDisplayUnit = Default, $iDisplayUnitCustom = Default, $vDisplayUnitLabel = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $iDisplayUnit <> Default And $oObject.Type <> $xlValue Then Return SetError(4, 0, 0)
	If $iDisplayUnitCustom <> Default And $oObject.Type <> $xlValue Then Return SetError(4, 0, 0)
	If $iDisplayUnitCustom <> Default And $oObject.DisplayUnit <> $xlCustom Then Return SetError(5, 0, 0)
	If $vDisplayUnitLabel <> Default And $oObject.Type <> $xlValue Then Return SetError(4, 0, 0)
	; DisplayUnit				Unit label for the value axis
	; DisplayUnitCustom			Custom value of the displayed units
	; HasDisplayUnitLabel		True if the label specified by the DisplayUnit or DisplayUnitCustom property is displayed
	; DisplayUnitLabel.Caption	String value that represents the display unit label text
	If $iDisplayUnit <> Default Then $oObject.DisplayUnit = $iDisplayUnit
	If $iDisplayUnitCustom <> Default Then $oObject.DisplayUnitCustom = $iDisplayUnitCustom
	If $vDisplayUnitLabel <> Default Then
		If IsBool($vDisplayUnitLabel) Then
			$oObject.HasDisplayUnitLabel = $vDisplayUnitLabel
		Else
			$oObject.DisplayUnitLabel.Caption = $vDisplayUnitLabel
		EndIf
	EndIf
	; MinimumScale 			Minimum value on the value axis
	; MinimumScaleIsAuto 	True if Excel calculates the minimum value for the value axis
	If $iMinimumScale <> Default Then
		If $iMinimumScale == "Auto" Then
			$oObject.MinimumScaleIsAuto = True
		ElseIf IsNumber($iMinimumScale) Then
			$oObject.MinimumScale = $iMinimumScale
			$oObject.MinimumScaleIsAuto = False
		Else
			Return SetError(2, 0, 0)
		EndIf
	EndIf
	; MaximumScale 			Maximum value on the value axis
	; MaximumScaleIsAuto 	True if Excel calculates the maximum value for the value axis
	If $iMaximumScale <> Default Then
		If $iMaximumScale == "Auto" Then
			$oObject.MaximumScaleIsAuto = True
		ElseIf IsNumber($iMaximumScale) Then
			$oObject.MaximumScale = $iMaximumScale
			$oObject.MaximumScaleIsAuto = False
		Else
			Return SetError(3, 0, 0)
		EndIf
	EndIf
	Return 1

EndFunc   ;==>_XLChart_AxisSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_BarGroupSet
; Description....: Set properties of a bar chart group.
; Syntax.........: _XLChart_BarGroupSet($oObject[, $iGapWidth = Default[, $iOverlap = Default[, $bHasSeriesLines = Default]]])
; Parameters ....: $oObject         - Chart group for which the properties should be set
;                  $iGapWidth       - Optional: Sets the space between bar clusters, as a percentage of the bar width (default = Default)
;                  $iOverlap        - Optional: Specifies how bars are positioned. Can be a value between -100 and 100 (default = Default)
;                  $bHasSeriesLines - Optional: True if the bar chart has series lines (default = Default). Only works for stacked bars
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is not an object
;                  |2 - $iGapWidth is not an integer
;                  |3 - $iOverlap is not an integer
;                  |4 - $bHasSeriesLines is not boolean
; Authors........: water
; Modified ......:
; Remarks .......: A chart contains one or more chart groups, each chart group contains one or more series, and
;                  each series contains one or more points.
;+
;                  You can either pass an item of the ChartGroups collection or an item of the BarGroups collection (a ChartGroup object)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_BarGroupSet($oObject, $iGapWidth = Default, $iOverlap = Default, $bHasSeriesLines = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $iGapWidth <> Default And Not IsInt($iGapWidth) Then Return SetError(2, 0, 0)
	If $iOverlap <> Default And Not IsInt($iOverlap) Then Return SetError(4, 0, 0)
	If $bHasSeriesLines <> Default And Not IsBool($bHasSeriesLines) Then Return SetError(2, 0, 0)
	If $iGapWidth <> Default Then $oObject.GapWidth = $iGapWidth
	If $iOverlap <> Default Then $oObject.Overlap = $iOverlap
	If $bHasSeriesLines <> Default Then $oObject.HasSeriesLines = $bHasSeriesLines
	Return 1

EndFunc   ;==>_XLChart_BarGroupSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_BubbleGroupSet
; Description....: Set properties of a bubble chart group.
; Syntax.........: _XLChart_BubbleGroupSet($oObject[, $iBubbleScale = Default[, $iSizeRepresents = Default[, $ibShowNegativeBubbles = Default]]])
; Parameters ....: $oObject               - Chart group for which the properties should be set
;                  $iBubbleScale          - Optional: Value from 0 to 300, corresponding to a percentage of the default size (default = Default)
;                  $iSizeRepresents       - Optional: Sets what the bubble size represents. Can be one of the XlSizeRepresents enumeration (default = Default)
;                  $ibShowNegativeBubbles - Optional: True if negative bubbles are shown for the chart group (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is not an object
;                  |2 - $iBubbleScale is not an integer
;                  |3 - $iSizeRepresents is not an integer
;                  |4 - $ibShowNegativeBubbles is not boolean
; Authors........: water
; Modified ......:
; Remarks .......: A chart contains one or more chart groups, each chart group contains one or more series, and
;                  each series contains one or more points.
;+
;                  You have to pass an item of the ChartGroups collection
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_BubbleGroupSet($oObject, $iBubbleScale = Default, $iSizeRepresents = Default, $ibShowNegativeBubbles = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $iBubbleScale <> Default And Not IsInt($iBubbleScale) Then Return SetError(2, 0, 0)
	If $iSizeRepresents <> Default And Not IsInt($iSizeRepresents) Then Return SetError(3, 0, 0)
	If $ibShowNegativeBubbles <> Default And Not IsBool($ibShowNegativeBubbles) Then Return SetError(4, 0, 0)
	If $iBubbleScale <> Default Then $oObject.BubbleScale = $iBubbleScale
	If $iSizeRepresents <> Default Then $oObject.SizeRepresents = $iSizeRepresents
	If $ibShowNegativeBubbles <> Default Then $oObject.ShowNegativeBubbles = $ibShowNegativeBubbles
	Return 1

EndFunc   ;==>_XLChart_BubbleGroupSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ChartCreate
; Description....: Create a chart in Excel on the specified worksheet or on a separate chartsheet.
; Syntax.........: _XLChart_ChartCreate($oExcel, $vWorksheet, $iChartType, $sSizeByCells, $sChartName, $sXValueRange, $vDataRange, $vDataName[, $bShowLegend = True[, $sTitle = ""[, $sXTitle = ""[, $sYTitle = ""[, $sZTitle = ""[, $bShowDatatable = False[, $bScreenUpdate = False]]]]]]])
; Parameters ....: $oExcel         - Excel object opened by a preceding call to _Excel_BookOpen() or _Excel_BookNew()
;                  $vWorksheet     - Worksheet number or name where the chart should be created (eg. 1)
;                  $iChartType     - Chart type number to be used (see ExcelChartConstants.au3 for details, eg. $xl3DColumn)
;                  $sSizeByCells   - The left-hand top and right-hand bottom corner of the chart (eg. "B2:K24").
;                  +                 0 = create the chart on a separate chartsheet. This sheet will be inserted before the sheet specified by $vWorksheet
;                  $sChartName     - Name of the chart or the chartsheet (if $sSizeByCells = 0)
;                  $sXValueRange   - Category (X) axis label range always a single range (eg. "=Sheet1!R2C1:R6C1")
;                  $vDataRange     - The values range. Either a single range or an one-dimensional one based array
;                  $vDataName      - Header name of the range. Either a single range or an one-dimensional one based array
;                  $bShowLegend    - Optional: Set to True to show the legend or False to hide it (default = True)
;                  $sTitle         - Optional: Chart Title. If empty, no title will be displayed (default = "")
;                  $sXTitle        - Optional: X Axis title (default = "")
;                  $sYTitle        - Optional: Y Axis title (default = "")
;                  $sZTitle        - Optional: Y Axis title (default = "")
;                  $bShowDatatable - Optional: Set to True to show the data table or False to hide it (default = False)
;                  $bScreenUpdate  - Optional: Set to False to disable screen updating during chart creation to enhance performance (default = False)
;                                    After the chart is created screen updating is reset to the value it had when calling this function
; Return values .: Success - Object identifier of the created chart, sets @extended to:
;                  |0 - No COM error handler has been initialized for this UDF because another COM error handler was already active
;                  |1 - A COM error handler has been initialized for this UDF
;                  Failure - Returns 0 and sets @error:
;                  |1 - Unable to create custom error handler. See @extended for details (error returned by ObjEvent)
;                  |2 - Both parameters $vDataRange & $vDataName must be of the same type
;                  |3 - Unable to access specified Excel sheet. See @extended for details
;                  |4 - Unable to create range for the chart. Parameter $sSizeByCells is invalid. See @extended for details
;                  |5 - Unable to create new chart. See @extended for details
; Authors........: GreenCan
; Modified ......: water, GreenCan
; Remarks .......: The COM error handler will be initialized only if there doesn't already exist another error handler.
;+
;                  Data Tables are available for line, column, area, and barchart types only.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ChartCreate($oExcel, $vWorksheet, $iChartType, $sSizeByCells, $sChartName, $sXValueRange, $vDataRange, $vDataName, $bShowLegend = True, $sTitle = "", $sXTitle = "", $sYTitle = "", $sZTitle = "", $bShowDatatable = False, $bScreenUpdate = False)

	Local $iErrorHandler = 0, $oNewChart, $oChartRange, $oChart, $iScreenUpdating
	; A COM error handler will be initialised only if one does not exist
	If ObjEvent("AutoIt.Error") = "" Then
		$g__oError = ObjEvent("AutoIt.Error", "_XLChart_COMError") ; Creates a custom COM error handler
		If @error <> 0 Then Return SetError(1, @error, 0)
		$iErrorHandler = 1
	EndIf
	If $bScreenUpdate = False Then
		$iScreenUpdating = $oExcel.ScreenUpdating
		$oExcel.ScreenUpdating = False ; Disable screen updating to enhance performance
	EndIf
	If (IsArray($vDataName) And Not (IsArray($vDataRange))) Or (Not (IsArray($vDataName)) And IsArray($vDataRange)) Then Return SetError(2, 0, 0)
	; Activate worksheet
	Local $oSheet = $oExcel.Worksheets($vWorksheet)
	If @error Then Return SetError(3, @error, 0)
	If IsNumber($sSizeByCells) And $sSizeByCells = 0 Then ; Create the new chart on a separate chartsheet
		$oNewChart = $oExcel.Charts.Add()
		If @error Then Return SetError(5, @error, 0)
		$oChart = $oNewChart
		If $sChartName <> "" Then $oChart.Name = $sChartName
	Else
		$oChartRange = $oSheet.Range($sSizeByCells) ; The range where you want the chart to be created
		If @error Then Return SetError(4, @error, 0)
		$oNewChart = $oSheet.ChartObjects.Add($oChartRange.Left, $oChartRange.Top, $oChartRange.Width, $oChartRange.Height) ; Create new chart
		If @error Then Return SetError(5, @error, 0)
		$oChart = $oNewChart.Chart
		If $sChartName <> "" Then $oNewChart.Name = $sChartName
	EndIf
	If $oChart.SeriesCollection.Count = 0 Then $oChart.SeriesCollection.NewSeries
	If IsArray($vDataRange) Then
		$oChart.SeriesCollection(1).Delete
		For $iIndex = 1 To $vDataName[0]
			$oChart.SeriesCollection.NewSeries
			$oChart.SeriesCollection($iIndex).Name = $vDataName[$iIndex] ; set name of values
			$oChart.SeriesCollection($iIndex).XValues = $sXValueRange ; X values
			$oChart.SeriesCollection($iIndex).Values = $vDataRange[$iIndex]
		Next
	Else
		$oChart.SeriesCollection(1).Name = $vDataName ; set name of values
		$oChart.SeriesCollection(1).XValues = $sXValueRange ; X values
		$oChart.SeriesCollection(1).Values = $vDataRange
	EndIf
	$oChart.HasDataTable = $bShowDatatable
	; Chart title
	If $sTitle <> "" Then
		$oChart.HasTitle = True
		$oChart.ChartTitle.Text = $sTitle
	Else
		$oChart.HasTitle = 0
	EndIf
	; Legend
	$oChart.HasLegend = $bShowLegend
	; X Axis title (1)
	If $sXTitle <> "" Then
		$oChart.Axes($xlCategory).HasTitle = True
		$oChart.Axes($xlCategory).AxisTitle.Text = $sXTitle
	EndIf
	; Y Axis title (2)
	If $sZTitle <> "" Then
		$oChart.Axes($xlValue).HasTitle = True
		$oChart.Axes($xlValue).AxisTitle.Text = $sZTitle
	EndIf
	; Z Axis title (3)
	If $sYTitle <> "" Then
		$oChart.Axes($xlSeriesAxis).HasTitle = True
		$oChart.Axes($xlSeriesAxis).AxisTitle.Text = $sYTitle
	EndIf
	$oChart.ChartType = $iChartType
	If $bScreenUpdate = False Then $oExcel.ScreenUpdating = $iScreenUpdating
	Return SetError(0, $iErrorHandler, $oChart)

EndFunc   ;==>_XLChart_ChartCreate

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ChartDataSet
; Description....: Sets all data related properties of an existing chart or chartsheet.
; Syntax.........: _XLChart_ChartDataSet($oChart, $sXValueRange, $vDataRange, $vDataName)
; Parameters ....: $oChart       - Chart object as returned by a preceding call to _XLChart_ChartCreate
;                  $sXValueRange - Category (X) axis label range always a single range (eg. "=Sheet1!R2C1:R6C1")
;                  $vDataRange   - The values range. Either a single range or an one-dimensional one based array
;                  $vDataName    - Header name of the range. Either a single range or an one-dimensional one based array
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oChart is not an object
;                  |2 - Both parameters $vDataRange & $vDataName must be of the same type
; Authors........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ChartDataSet($oChart, $sXValueRange, $vDataRange, $vDataName)

	If Not IsObj($oChart) Then Return SetError(1, 0, 0)
	If (IsArray($vDataName) And Not (IsArray($vDataRange))) Or (Not (IsArray($vDataName)) And IsArray($vDataRange)) Then Return SetError(2, 0, 0)
	If $oChart.SeriesCollection.Count = 0 Then $oChart.SeriesCollection.NewSeries
	If IsArray($vDataRange) Then
		For $oSerie In $oChart.SeriesCollection
			$oSerie.Delete
		Next
		For $iIndex = 1 To $vDataName[0]
			$oChart.SeriesCollection.NewSeries
			$oChart.SeriesCollection($iIndex).Name = $vDataName[$iIndex] ; set name of values
			$oChart.SeriesCollection($iIndex).XValues = $sXValueRange ; X values
			$oChart.SeriesCollection($iIndex).Values = $vDataRange[$iIndex]
		Next
	Else
		$oChart.SeriesCollection(1).Name = $vDataName ; set name of values
		$oChart.SeriesCollection(1).XValues = $sXValueRange ; X values
		$oChart.SeriesCollection(1).Values = $vDataRange
	EndIf
	Return 1

EndFunc   ;==>_XLChart_ChartDataSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ChartDelete
; Description....: Deletes a chart or a chartsheet.
; Syntax.........: _XLChart_Delete($oExcel, $vChart)
; Parameters ....: $oExcel - Excel object opened by a preceding call to _Excel_BookOpen() or _Excel_BookNew()
;                  $vChart - Chart object as returned by a preceding call to _XLChart_ChartCreate or name of a chart
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oExcel is not an object
;                  |2 - Error occurred deleting $vChart. See @extended for details (error returned by method Delete)
;                  |3 - $vChart could not be found. See @extended for details (error returned by method Delete)
; Authors........: water
; Modified ......:
; Remarks .......: The chart name is the name of the worksheet + " " + chart name as specified in _XLChart_Create
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ChartDelete($oExcel, $vChart)

	If Not IsObj($oExcel) Then Return SetError(1, 0, 0)
	Local $iError, $bFound = False
	; Try to delete the object as a chartsheet
	Local $bDisplayAlerts = $oExcel.DisplayAlerts
	If $bDisplayAlerts <> False Then $oExcel.DisplayAlerts = False ; Suppress alerts that require user intervention
	For $oChartSheet In $oExcel.Charts
		If $oChartSheet = $vChart Or $oChartSheet.Name = $vChart Then
			$oChartSheet.Delete()
			$iError = @error
			$bFound = True
			ExitLoop
		EndIf
	Next
	$oExcel.DisplayAlerts = $bDisplayAlerts ; Reset alerts
	If $iError <> 0 Then Return SetError(2, $iError, 0)
	; Try to delete the object as an embedded chart
	If $bFound = False Then
		If IsObj($vChart) Then
			$vChart.Parent.Delete()
		Else
			For $oChart In $oExcel.ActiveSheet.ChartObjects
				If $oChart.Chart.Name = $vChart Then
					$oChart.Delete()
					$iError = @error
					$bFound = True
					ExitLoop
				EndIf
			Next
		EndIf
		If @error <> 0 Then Return SetError(3, @error, 0)
	EndIf
	Return 1

EndFunc   ;==>_XLChart_ChartDelete

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ChartExport
; Description....: Exports the chart/chartsheet in a graphic format (GIF, JPG, PNG ...) or as PDF/XPS.
; Syntax.........: _XLChart_ChartExport($oObject, $sFilename, $sFilterName[, $bInteractive = False[, $bOverwrite = False]])
; Parameters ....: $oObject      - Chart or chartsheet object as returned by a preceding call to _XLChart_ChartCreate
;                  $sFilename    - Path/name of the exported file
;                  $sFilterName  - The language-independent name of the graphic filter as it appears in the registry.
;                                  To see which graphics filters are installed, check the following registry key:
;                                  HKEY_LOCAL_MACHINE\Software\Microsoft\Shared Tools\Graphics Filters\Export or
;                                  HKEY_LOCAL_MACHINE\Software\WoW6432Node\Microsoft\Shared Tools\Graphics Filters\Export
;                                  plus "PDF" and "XPS" (method ExportAsFixedFormat is used)
;                  $bInteractive - Optional: True to display the dialog box that contains the filter-specific options.
;                                  If False then Excel uses the default values for the filter (default = False)
;                  $bOverwrite   - Optional: True to overwrite an existing version of the output file (default = False)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is not an object
;                  |2 - $bInteractive is not boolean
;                  |3 - $bOverwrite is not boolean
;                  |4 - $sFilename already exists and $bOverwrite = False
;                  |5 - $sFilename could not be deleted and $bOverwrite = True
;                  |6 - $sFilterName is empty
;                  |7 - Error exporting the chart. See @extended for details (error returned by method Export or ExportAsFixedFormat)
; Authors........: water
; Modified ......:
; Remarks .......: We noticed this function crash when creating a PDF file on Excel 2007. Problem is still under investigation.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ChartExport($oObject, $sFilename, $sFilterName, $bInteractive = False, $bOverwrite = False)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If Not IsBool($bInteractive) Then Return SetError(2, 0, 0)
	If Not IsBool($bOverwrite) Then Return SetError(3, 0, 0)
	If FileExists($sFilename) Then
		If $bOverwrite = True Then
			Local $iResult = FileDelete($sFilename)
			If $iResult = 0 Then Return SetError(5, 0, 0)
		Else
			Return SetError(4, 0, 0)
		EndIf
	EndIf
	If $sFilterName = "" Then Return SetError(6, 0, 0)
	If $sFilterName = "PDF" Then
		$oObject.Parent.Activate
		$oObject.ExportAsFixedFormat($xlTypePDF, $sFilename, 0, True)
	ElseIf $sFilterName = "XPS" Then
		$oObject.Parent.Activate
		$oObject.ExportAsFixedFormat($xlTypeXPS, $sFilename)
	Else
		$oObject.Export($sFilename, $sFilterName, $bInteractive)
	EndIf
	If @error <> 0 Then Return SetError(7, @error, 0)
	Return 1

EndFunc   ;==>_XLChart_ChartExport

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ChartPositionSet
; Description....: Resize and reposition a chart object.
; Syntax.........: _XLChart_ChartPositionSet($oExcel, $oChart[, $sSizeByCells = ""[, $iLeft = Default[, $iTop = Default[, $iWidth = Default[, $iHeight = Default[, $iFlag = 0[, $vSheet = ""]]]]]]])
; Parameters ....: $oExcel       - Excel object opened by a preceding call to _Excel_BookOpen() or _Excel_BookNew()
;                  $oChart       - Chart object as returned by a preceding call to _XLChart_ChartCreate
;                  $sSizeByCells - Optional: The left-hand top and right-hand bottom corner of the chart (eg. "B2:K24") (default = "")
;                  $iLeft        - Optional: Distance, in points, from the left edge of the chart to the left edge of column A (default = Default)
;                  $iTop         - Optional: Distance, in points, from the top edge of the chart to the top of row 1 (default = Default)
;                  $iWidth       - Optional: Width, in points, of the chart (default = Default)
;                  $iHeight      - Optional: Height, in points, of the chart default = Default)
;                  $iFlag        - Optional: 1 = add the left/top/widht/height value to the current value, 0 = set the left/top/widht/height value (default = 0)
;                  $vSheet       - Optional: Name of the sheet where to move the chart to. "-1" = move chart to a new chartsheet (default = "")
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oExcel is not an object
;                  |2 - $oChart is not an object
;                  |3 - You have to set $sSizeByCells or $iWidth, $iHeight, $iLeft, $iTop but not both
;                  |4 - Unable to create range for the chart. Parameter $sSizeByCells is invalid. See @extended for details
;                  |5 - $iLeft is not a number
;                  |6 - $iTop is not a number
;                  |7 - $iWidth is not a number
;                  |8 - $iHeight is not a number
;                  |9 - $iFlag is not an integer or < 0 or > 1
;                  |10 - $vSheet is not a string or -1
;                  |11 - Error returned by _ExcelSheetActivate. Most likely: $vSheet does not exist. See @extended for error code of _ExcelSheetActivate
;                  |12 - Error moving the chart to a new worksheet or chartsheet. See @extended for error code of the Location method
; Authors........: GreenCan
; Modified ......: water
; Remarks .......: If $vSheet = -1 all other size parameters will be ignored.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ChartPositionSet($oExcel, $oChart, $sSizeByCells = "", $iLeft = Default, $iTop = Default, $iWidth = Default, $iHeight = Default, $iFlag = 0, $vSheet = "")

	If Not IsObj($oExcel) Then Return SetError(1, 0, 0)
	If Not IsObj($oChart) Then Return SetError(2, 0, 0)
	If $vSheet <> -1 And (($sSizeByCells <> "" And ($iWidth <> Default Or $iHeight <> Default Or $iLeft <> Default Or $iTop <> Default)) Or _
			($sSizeByCells = "" And ($iWidth = Default And $iHeight = Default And $iLeft = Default And $iTop = Default))) Then _
			Return SetError(3, @error, 0)
	If $iLeft <> Default And Not IsNumber($iLeft) Then Return SetError(5, 0, 0)
	If $iTop <> Default And Not IsNumber($iTop) Then Return SetError(6, 0, 0)
	If $iWidth <> Default And Not IsNumber($iWidth) Then Return SetError(7, 0, 0)
	If $iHeight <> Default And Not IsNumber($iHeight) Then Return SetError(8, 0, 0)
	If Not IsInt($iFlag) Or $iFlag < 0 Or $iFlag > 1 Then Return SetError(9, 0, 0)
	If Not (IsString($vSheet) Or $vSheet = -1) Then Return SetError(10, 0, 0)
	If $vSheet <> "" Then
		If IsString($vSheet) Then
;			_ExcelSheetActivate($oExcel, $vSheet)
			$oExcel.ActiveWorkbook.Sheets($vSheet).Select()
			If @error Then Return SetError(11, @error, 0)
			$oChart = $oChart.Location($xlLocationAsObject, $vSheet)
			If @error Then Return SetError(12, @error, 0)
		Else
			$oChart = $oChart.Location($xlLocationAsNewSheet)
			If @error Then Return SetError(12, @error, 0)
			Return 1
		EndIf
	EndIf
	If $sSizeByCells <> "" Then
		Local $oChartRange = $oExcel.Range($sSizeByCells) ; The new range of the chart
		If @error Then Return SetError(4, @error, 0)
		$oChart.Parent.left = $oChartRange.Left
		$oChart.Parent.Top = $oChartRange.Top
		$oChart.Parent.Width = $oChartRange.Width
		$oChart.Parent.Height = $oChartRange.Height
	Else
		If $iFlag = 1 Then
			If $iLeft <> Default Then $oChart.Parent.left = $oChart.Parent.left + $iLeft
			If $iTop <> Default Then $oChart.Parent.Top = $oChart.Parent.Top + $iTop
			If $iWidth <> Default Then $oChart.Parent.Width = $oChart.Parent.Width + $iWidth
			If $iHeight <> Default Then $oChart.Parent.Height = $oChart.Parent.Height + $iHeight
		Else
			If $iLeft <> Default Then $oChart.Parent.left = $iLeft
			If $iTop <> Default Then $oChart.Parent.Top = $iTop
			If $iWidth <> Default Then $oChart.Parent.Width = $iWidth
			If $iHeight <> Default Then $oChart.Parent.Height = $iHeight
		EndIf
	EndIf
	Return 1

EndFunc   ;==>_XLChart_ChartPositionSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ChartPrint
; Description....: Print a chart or a chartsheet.
; Syntax.........: _XLChart_ChartPrint($oChart[, $sActivePrinter = Default[, $iCopies = 1[, $bPreview = False[, $bPrintToFile = False[, $sPrintToFileName = Default]]]]])
; Parameters ....: $oChart           - Object to be printed. Can be a chart or a chartsheet.
;                  $sActivePrinter   - Optional: Name of the printer to be used. Defaults to active printer (default = Default)
;                                      Example: \\Spoolservername\Printername
;                  $iCopies          - Optional: Number of copies to print (default = 1)
;                  $bPreview         - Optional: If True a print preview is displayed (default = False)
;                  $bPrintToFile     - Optional: If True the output is written to a file (default = False)
;                  $sPrintToFileName - Optional: If $bPrintToFile is True, this argument specifies the name of the file you want to print to (default = Default)
;                                          If $bPrtToFileName = Default, Excel prompts the user to enter the name of the output file
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oChart is not of type object
;                  |2 - $iCopies is not numeric or < 1
;                  |3 - $bPreview is not boolean
;                  |4 - $bPrintToFile is not boolean
;                  |5 - Error printing the specified object. See @extended for details (error returned by method PrintOut)
; Authors........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ChartPrint($oChart, $sActivePrinter = Default, $iCopies = 1, $bPreview = False, $bPrintToFile = False, $sPrintToFileName = Default)

	If Not IsObj($oChart) Then Return SetError(1, 0, 0)
	If $iCopies <> Default And (Not IsInt($iCopies) Or $iCopies < 1) Then Return SetError(2, 0, 0)
	If $bPreview <> Default And Not IsBool($bPreview) Then Return SetError(3, 0, 0)
	If $bPrintToFile <> Default And Not IsBool($bPrintToFile) Then Return SetError(4, 0, 0)
	$oChart.PrintOut(Default, Default, $iCopies, $bPreview, $sActivePrinter, $bPrintToFile, Default, $sPrintToFileName)
	If @error Then Return SetError(5, @error, 0)
	Return 1

EndFunc   ;==>_XLChart_ChartPrint

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ChartSet
; Description....: Set properties of a chart.
; Syntax.........: _XLChart_ChartSet($oChart[, $iChartType = Default[, $iDisplayBlanksAs = Default[, $iPlotBy = Default[, $bRoundedCorners  = False]]]])
; Parameters ....: $oChart           - Chart object as returned by a preceding call to _XLChart_ChartCreate
;                  $iChartType       - Optional: Chart type number to be used. Can be any of the XlChartType enumeration (default = Default)
;                  $iDisplayBlanksAs - Optional: Specifies how blank cells are plotted. Can be one of the XlDisplayBlanksAs enumeration (default = Default)
;                  $iPlotBy          - Optional: Specifies whether the values corresponding to a particular data series are in rows or columns. Can be one of the XlRowCol enumeration ($xlColumns or $xlRows) (default = Default)
;                  $bRoundedCorners  - Optional: If Set to True the chart will have rounded corners (default = False). Not valid for chartsheets
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oChart is not an object
;                  |2 - $iChartType is not an integer
;                  |3 - $iDisplayBlanksAs is not an integer
;                  |4 - $iPlotBy is not an integer
;                  |5 - $bRoundedCorners is not boolean
; Authors........: water, GreenCan
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ChartSet($oChart, $iChartType = Default, $iDisplayBlanksAs = Default, $iPlotBy = Default, $bRoundedCorners = False)

	If Not IsObj($oChart) Then Return SetError(1, 0, 0)
	If $iChartType <> Default And Not IsInt($iChartType) Then Return SetError(2, 0, 0)
	If $iDisplayBlanksAs <> Default And Not IsInt($iDisplayBlanksAs) Then Return SetError(3, 0, 0)
	If $iPlotBy <> Default And Not IsInt($iPlotBy) Then Return SetError(4, 0, 0)
	If $bRoundedCorners <> Default And Not IsBool($bRoundedCorners) Then Return SetError(5, 0, 0)
	If $iChartType <> Default Then $oChart.ChartType = $iChartType
	If $iDisplayBlanksAs <> Default Then $oChart.DisplayBlanksAs = $iDisplayBlanksAs
	If $iPlotBy <> Default Then $oChart.PlotBy = $iPlotBy
	If $bRoundedCorners <> Default Then $oChart.Parent.RoundedCorners = $bRoundedCorners
	Return 1

EndFunc   ;==>_XLChart_ChartSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ChartsGet
; Description....: Enumerate charts and chartsheets in a workbook.
; Syntax.........: _XLChart_ChartsGet($oExcel[, $iChartType = 3[, $vWorksheet = -1]])
; Parameters ....: $oExcel     - Excel object opened by a preceding call to _Excel_BookOpen() or _Excel_BookNew()
;                  $iChartType - Optional: Type of charts to return. 1 = charts, 2 = chartsheets, 3 = 1 + 2 (default = 3)
;                  $vWorksheet - Optional: Worksheet number or name. 0 = Active Worksheet, -1 = All worksheets (default = -1)
; Return values .: Success - two-dimensional one based array with the following information:
;                  |0 - Object of the chart or chartsheet
;                  |1 - Type of the object. 1 = chart, 2 = chartsheet
;                  |2 - Name of the chart or chartsheet
;                  |3 - Number of the Excel sheet where the chart or chartsheet resides
;                  |4 - Name of the Excel sheet where the chart or chartsheet resides
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oExcel is not an object
;                  |2 - $iChartType is not numeric or < 1 or > 3
;                  |3 - $vWorksheet > number of sheets in workbook
;                  |4 - $vWorksheet could not be found
; Authors........: GreenCan
; Modified ......: water
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ChartsGet($oExcel, $iChartType = 3, $vWorksheet = -1)

	If Not IsObj($oExcel) Then Return SetError(1, 0, 0)
	If Not IsInt($iChartType) Or $iChartType < 1 Or $iChartType > 3 Then Return SetError(2, 0, 0)
	If IsInt($vWorksheet) And $vWorksheet > $oExcel.Sheets.Count Then Return SetError(3, 0, 0)
	Local $iWorksheetStart, $iWorksheetEnd, $bFound = False
	If IsString($vWorksheet) Then ; Name of worksheet
		For $oSheet In $oExcel.Sheets
			If StringLower($oSheet.Name) = $vWorksheet Then
				$iWorksheetStart = $oSheet.Index
				$iWorksheetEnd = $iWorksheetStart
				$bFound = True
				ExitLoop
			EndIf
		Next
		If $bFound = False Then Return SetError(4, 0, 0)
	ElseIf $vWorksheet = -1 Then ; All Worksheets
		$iWorksheetStart = 1
		$iWorksheetEnd = $oExcel.Sheets.Count
	ElseIf $vWorksheet = 0 Then ; Current Worksheet
		$iWorksheetStart = $oExcel.ActiveSheet.Index
		$iWorksheetEnd = $iWorksheetStart
	Else ; Number of worksheet
		$iWorksheetStart = $vWorksheet
		$iWorksheetEnd = $vWorksheet
	EndIf
	Local $aResult[1][5] = [[0, 5]], $iTabIndex = 0
	For $iIndex1 = $iWorksheetStart To $iWorksheetEnd
		If BitAND($iChartType, 1) = 1 Then ; Charts
			For $oChart In $oExcel.Sheets($iIndex1).ChartObjects
				ReDim $aResult[$aResult[0][0] + 2][$aResult[0][1]]
				$iTabIndex += 1
				$aResult[0][0] = $iTabIndex
				$aResult[$iTabIndex][0] = $oChart.Chart
				$aResult[$iTabIndex][1] = 1
				$aResult[$iTabIndex][2] = $oChart.Name
				$aResult[$iTabIndex][3] = $iIndex1
				$aResult[$iTabIndex][4] = $oExcel.Sheets($iIndex1).Name
			Next
		EndIf
		If BitAND($iChartType, 2) = 2 Then ; Chartsheets
			For $oChartSheet In $oExcel.Charts
				If $oChartSheet.Index = $iIndex1 Then
					ReDim $aResult[$aResult[0][0] + 2][$aResult[0][1]]
					$iTabIndex += 1
					$aResult[0][0] = $iTabIndex
					$aResult[$iTabIndex][0] = $oChartSheet
					$aResult[$iTabIndex][1] = 2
					$aResult[$iTabIndex][2] = $oChartSheet.Name
					$aResult[$iTabIndex][3] = $iIndex1
					$aResult[$iTabIndex][4] = $oExcel.Sheets($iIndex1).Name
				EndIf
				ExitLoop
			Next
		EndIf
	Next
	Return $aResult

EndFunc   ;==>_XLChart_ChartsGet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ColumnGroupSet
; Description....: Set properties of a column chart group.
; Syntax.........: _XLChart_ColumnGroupSet($oObject[, $iGapWidth = Default[, $iOverlap = Default[, $bHasSeriesLines = Default]]])
; Parameters ....: $oObject         - Chart group for which the properties should be set
;                  $iGapWidth       - Optional: Sets the space between column clusters, as a percentage of the column width (default = Default)
;                  $iOverlap        - Optional: Specifies how columns are positioned. Can be a value between -100 and 100 (default = Default)
;                  $bHasSeriesLines - Optional: True if the column chart has series lines (default = Default). Only works for stacked columns
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is not an object
;                  |2 - $iGapWidth is not an integer
;                  |3 - $iOverlap is not an integer
;                  |4 - $bHasSeriesLines is not boolean
; Authors........: water
; Modified ......:
; Remarks .......: A chart contains one or more chart groups, each chart group contains one or more series, and
;                  each series contains one or more points.
;+
;                  You can either pass an item of the ChartGroups collection or an item of the ColumnGroups collection (a ChartGroup object)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ColumnGroupSet($oObject, $iGapWidth = Default, $iOverlap = Default, $bHasSeriesLines = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $iGapWidth <> Default And Not IsInt($iGapWidth) Then Return SetError(2, 0, 0)
	If $iOverlap <> Default And Not IsInt($iOverlap) Then Return SetError(4, 0, 0)
	If $bHasSeriesLines <> Default And Not IsBool($bHasSeriesLines) Then Return SetError(2, 0, 0)
	If $iGapWidth <> Default Then $oObject.GapWidth = $iGapWidth
	If $iOverlap <> Default Then $oObject.Overlap = $iOverlap
	If $bHasSeriesLines <> Default Then $oObject.HasSeriesLines = $bHasSeriesLines
	Return 1

EndFunc   ;==>_XLChart_ColumnGroupSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_DatalabelSet
; Description....: Set properties for the data labels of a data series.
; Syntax.........: _XLChart_DatalabelSet($oObject[, $iType = Default[, $bShowValue = Default[, $bShowPercentage = Default[, $bShowLegendKey = Default[, $bShowSeriesName = Default[, $bShowCategoryName = Default[, $iPosition = Default[, $iOrientation = Default]]]]]]]])
; Parameters ....: $oObject           - Object of the data series for which the data label properties should be set
;                  $iType             - Optional: Type of data label to apply. Can be one of the XlDataLabelsType enumeration (default = Default)
;                  $bShowValue        - Optional: True enables the value for the data label (default = Default)
;                  $bShowPercentage   - Optional: True enables the percentage for the data label (default = Default)
;                  $bShowLegendKey    - Optional: True shows the legend key next to the point (default = Default)
;                  $bShowSeriesName   - Optional: True enables the series name for the data label (default = Default)
;                  $bShowCategoryName - Optional: True enables the category name for the data label (default = Default)
;                  $iPosition         - Optional: Represents the position of the data label. Can be any of the XlDataLabelPosition enumeration (default = Default)
;                  $iOrientation      - Optional: Represents the text orientation. Can be a value from -90 to 90 degree (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is no object
;                  |2 - $iType is not an integer
;                  |3 - $iShowValue is not boolean
;                  |4 - $bShowPercentage is not boolean
;                  |5 - $bShowLegendKey is not boolean
;                  |6 - $bShowSeriesName is not boolean
;                  |7 - $bShowCategoryName is not boolean
;                  |8 - Error setting properties. See @extended for details (error returned by method ApplyDataLabels)
;                  |9 - $iPosition is not an integer
;                  |10 - $iOrientation is not an integer
; Authors........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_DatalabelSet($oObject, $iType = Default, $bShowValue = Default, $bShowPercentage = Default, $bShowLegendKey = Default, $bShowSeriesName = Default, $bShowCategoryName = Default, $iPosition = Default, $iOrientation = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $iType <> Default And Not IsInt($iType) Then Return SetError(2, 0, 0)
	If $bShowValue <> Default And Not IsBool($bShowValue) Then Return SetError(3, 0, 0)
	If $bShowPercentage <> Default And Not IsBool($bShowPercentage) Then Return SetError(4, 0, 0)
	If $bShowLegendKey <> Default And Not IsBool($bShowLegendKey) Then Return SetError(5, 0, 0)
	If $bShowSeriesName <> Default And Not IsBool($bShowSeriesName) Then Return SetError(6, 0, 0)
	If $bShowCategoryName <> Default And Not IsBool($bShowCategoryName) Then Return SetError(7, 0, 0)
	If $iPosition <> Default And Not IsInt($iPosition) Then Return SetError(9, 0, 0)
	If $iOrientation <> Default And (Not (IsInt($iOrientation)) Or $iOrientation < -90 Or $iOrientation > 90) Then Return SetError(10, 0, 0)
	$oObject.ApplyDataLabels($iType, $bShowLegendKey, Default, Default, $bShowSeriesName, $bShowCategoryName, $bShowValue, $bShowPercentage)
	If @error Then Return SetError(8, @error, 0)
	If $iPosition <> Default Then $oObject.Datalabels.Position = $iPosition
	If $iOrientation <> Default Then $oObject.Datalabels.Orientation = $iOrientation
	Return 1

EndFunc   ;==>_XLChart_DatalabelSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_DoughnutGroupSet
; Description....: Set properties of a doughnut chart group.
; Syntax.........: _XLChart_DoughnutGroupSet($oObject[, $iFirstSliceAngle = Default[, $iDoughnutHoleSize = Default]])
; Parameters ....: $oObject           - Chart group for which the properties should be set
;                  $iFirstSliceAngle  - Optional: Angle of the first doughnut slice in degrees (clockwise from vertical).
;                  +                    Can be a value from 0 through 360 (default = Default)
;                  $iDoughnutHoleSize - Optional: Size of the hole in a doughnut chart group expressed as a percentage of the chart size,
;                  +between 10 and 90 percent (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is not an object
;                  |2 - $iFirstSliceAngle is not an integer or < 0 or > 360
;                  |2 - $iDoughnutHoleSize is not an integer or < 10 or > 90
; Authors........: water
; Modified ......:
; Remarks .......: A chart contains one or more chart groups, each chart group contains one or more series, and
;                  each series contains one or more points.
;+
;                  You can either pass an item of the ChartGroups collection or an item of the DoughnutGroups collection (a ChartGroup object)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_DoughnutGroupSet($oObject, $iFirstSliceAngle = Default, $iDoughnutHoleSize = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $iFirstSliceAngle <> Default And (Not (IsInt($iFirstSliceAngle)) Or $iFirstSliceAngle < 0 Or $iFirstSliceAngle > 360) Then _
			Return SetError(2, 0, 0)
	If $iDoughnutHoleSize <> Default And (Not (IsInt($iDoughnutHoleSize)) Or $iDoughnutHoleSize < 10 Or $iDoughnutHoleSize > 90) Then _
			Return SetError(3, 0, 0)
	If $iFirstSliceAngle <> Default Then $oObject.FirstSliceAngle = $iFirstSliceAngle
	If $iDoughnutHoleSize <> Default Then $oObject.DoughnutHoleSize = $iDoughnutHoleSize
	Return 1

EndFunc   ;==>_XLChart_DoughnutGroupSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ErrorBarSet
; Description....: Add or set properties of error bars for a data series.
; Syntax.........: _XLChart_ErrorBarSet($oObject, $iDirection, $iInclude, $iType[, $iEndStyle = Default[, $iAmount = Default[, $iMinusValues = Default]]])
; Parameters ....: $oObject      - Data series object to add or change an ErrorBar
;                  $iDirection   - The error bar direction. Value has to be of the XlErrorBarDirection enumeration
;                  $iInclude     - The error bar parts to include. Value has to be of the XlErrorBarInclude enumeration
;                  $iType        - The error bar type. Value has to be of the XlErrorBarType enumeration
;                  $iEndStyle    - Optional: End style for the error bars. Can be one of the XlEndStyleCap enumeration (default = Default)
;                  $iAmount      - Optional: The positive error amount when $iType is xlErrorBarTypeCustom (default = Default)
;                  $iMinusValues - Optional: The negative error amount when $iType is xlErrorBarTypeCustom (default = Default)
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is no object
;                  |2 - $iDirection is not an integer
;                  |3 - $iInclude is not an integer
;                  |4 - $iType is not an integer
;                  |5 - $iEndStyle is not an integer
;                  |6 - $iAmount is not an integer
;                  |7 - $iMinusValues is not an integer
;                  |8 - Error creating errorbars. See @extended for details
; Authors........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ErrorBarSet($oObject, $iDirection, $iInclude, $iType, $iEndStyle = Default, $iAmount = Default, $iMinusValues = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If Not IsInt($iDirection) Then Return SetError(2, 0, 0)
	If Not IsInt($iInclude) Then Return SetError(3, 0, 0)
	If Not IsInt($iType) Then Return SetError(4, 0, 0)
	If $iEndStyle <> Default And Not IsInt($iEndStyle) Then Return SetError(5, 0, 0)
	If $iAmount <> Default And Not IsInt($iAmount) Then Return SetError(6, 0, 0)
	If $iMinusValues <> Default And Not IsInt($iMinusValues) Then Return SetError(7, 0, 0)
	$oObject.HasErrorBars = True
	$oObject.ErrorBar($iDirection, $iInclude, $iType, $iAmount, $iMinusValues)
	If @error Then Return SetError(8, @error, 0)
	If $iEndStyle <> Default Then $oObject.ErrorBars.EndStyle = $iEndStyle
	Return 1

EndFunc   ;==>_XLChart_ErrorBarSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_FillSet
; Description....: Set fill properties for the specified object.
; Syntax.........: _XLChart_FillSet($oObject[, $iForeColor = Default[, $iBackColor = Default[, $bThemeColor = False[, $iGradientStyle = Default[, $iGradientVariant = Default[, $iTransparency = Default[, $sBitmap = Default[, $bTextureTile = Default]]]]]]]])
; Parameters ....: $oObject          - Object for which the fill properties should be set (...)
;                  $iForeColor       - Optional: Sets the foreground fill color (default = Default)
;                  +                   You can set colors to an explicit red-green-blue value (e.g. 0xFF00FF) or to a color in the color scheme (negative numbers -1 to -56)
;                  $iBackColor       - Optional: Sets the background fill color (default = Default)
;                  +                   You can set colors to an explicit red-green-blue value (e.g. 0xFF00FF) or to a color in the color scheme (negative numbers -1 to -56)
;                  $bThemeColor      - Optional: True specifies that $iForeColor and $iBackColor are interpreted as theme colors (default = False).
;                  +                   If set to True $iForeColor and $iBackColor values have to be one of the MsoThemeColorIndex enumeration
;                  $iGradientStyle   - Optional: The gradient style as sepcified by the MsoGradientStyle enumeration (default = Default).
;                  +                   If $iGradientStyle is sepcified you have to specify $iGradientVariant as well
;                  $iGradientVariant - Optional: The gradient variant. Can be a value from 1 through 4, corresponding to one of the four variants on the Gradient tab in the Fill Effects dialog box.
;                                      If $iGradientStyle is msoGradientFromCenter, $iGradientVariant can only be 1 or 2 (default = Default)
;                  $iDegree          - Optional: The gradient degree. Can be a value from 0.0 (dark) through 1.0 (light) (default = Default)
;                                      If $iGradientStyle is set and this value is Default then a two color gradient will be used otherwise a one color gradient
;                  $iTransparency    - Optional: Sets the degree of transparency from 0.0 (opaque) through 1.0 (clear) (default = Default)
;                                      To make it work you have to specify a foreground color
;                  $sBitmap          - Optional: Picture full path or preset texture used to fill the object background
;                  $bTextureTile     - Optional: if True set picture or preset texture defined by $sBitmap tiled
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is no object
;                  |2 - $iForeColor is not an integer
;                  |3 - $iBackColor is not an integer
;                  |4 - $iGradientStyle is not an integer
;                  |5 - $iGradientVariant is not an integer, < 1 or > 4 or > 2 if $iGradientStyle = $msoGradientFromCenter
;                  |6 - $iGradientVariant and $iGradientVariant have to be specified both or none
;                  |7 - $iTransparency is not a number or < 0 or > 1
;                  |8 - $iDegree is not a number or < 0 or > 1
;                  |9 - $bThemeColor is not boolean
;                  |10 - $sBitmap path/filename not found or error setting $sBitmap
;                  |11 - $sBitmap preset Texture does not exist
;                  |12 - $bTextureTile is not boolean or error setting $bTextureTile
; Authors........: water
; Modified ......: Greencan
; Remarks .......: Color 0 (white) has to be specified as RGB value or use color 2 which has the same RGB value of 0xFFFFFF
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_FillSet($oObject, $iForeColor = Default, $iBackColor = Default, $bThemeColor = False, $iGradientStyle = Default, $iGradientVariant = Default, $iDegree = Default, $iTransparency = Default, $sBitmap = Default, $bTextureTile = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $bThemeColor = Default Then $bThemeColor = False
	If Not IsBool($bThemeColor) Then Return SetError(9, 0, 0)
	If $iGradientStyle <> Default Then
		If $iGradientVariant = Default Then Return SetError(6, 0, 0)
		If Not IsInt($iGradientStyle) Then Return SetError(4, 0, 0)
		If $iDegree <> Default And (Not (IsNumber($iDegree)) Or $iDegree < 0 Or $iDegree > 1) Then Return SetError(8, 0, 0)
		If Not IsInt($iGradientVariant) Or _
				$iGradientVariant < 1 Or $iGradientVariant > 4 Or _
				($iGradientStyle = $msoGradientFromCenter And $iGradientVariant > 2) Then Return SetError(5, 0, 0)
		If $iDegree = Default Then
			$oObject.Format.Fill.TwoColorGradient($iGradientStyle, $iGradientVariant)
		Else
			$oObject.Format.Fill.OneColorGradient($iGradientStyle, $iGradientVariant, $iDegree)
		EndIf
	EndIf
	If $iForeColor <> Default Then
		If Not IsInt($iForeColor) Then Return SetError(2, 0, 0)
		If $bThemeColor Then
			$oObject.Format.Fill.ForeColor.ObjectThemeColor = $iForeColor
		Else
			If $iForeColor < 0 Then
				; Add 7 to ColorIndex to convert to SchemeColor: http://www.ozgrid.com/forum/showthread.php?t=53791
				$oObject.Format.Fill.ForeColor.SchemeColor = Abs($iForeColor) + 7
			Else
				$oObject.Format.Fill.ForeColor.RGB = _XLChart_RGB($iForeColor)
			EndIf
		EndIf
	EndIf
	If $iBackColor <> Default And $iDegree = Default Then ; Ignore BackgroundColor for one color gradients
		If Not IsInt($iBackColor) Then Return SetError(3, 0, 0)
		If $bThemeColor Then
			$oObject.Format.Fill.BackColor.ObjectThemeColor = $iBackColor
		Else
			If $iBackColor < 0 Then
				; Add 7 to ColorIndex to convert to SchemeColor: http://www.ozgrid.com/forum/showthread.php?t=53791
				$oObject.Format.Fill.BackColor.SchemeColor = Abs($iBackColor) + 7
			Else
				$oObject.Format.Fill.BackColor.RGB = _XLChart_RGB($iBackColor)
			EndIf
		EndIf
	EndIf
	If $sBitmap <> Default Then
		If StringInStr($sBitmap, "\") > 0 Then ; Bitmap
			If Not FileExists($sBitmap) Then Return SetError(10, 0, 0)
			$oObject.Format.Fill.UserPicture($sBitmap)
			If @error Then Return SetError(10, @error, 0)
		Else
			$oObject.Format.Fill.PresetTextured($sBitmap)
			If @error Then Return SetError(11, @error, 0)
		EndIf
	EndIf
	If $bTextureTile = Default Then $bTextureTile = False
	If Not IsBool($bTextureTile) Then Return SetError(12, 0, 0)
	If $bTextureTile Then
		$oObject.Format.Fill.TextureTile = True
		If @error Then Return SetError(12, @error, 0)
	EndIf
	If $iTransparency <> Default Then
		If Not IsNumber($iTransparency) Or $iTransparency < 0 Or $iTransparency > 1 Then Return SetError(7, 0, 0)
		$oObject.Format.Fill.Transparency = $iTransparency
	EndIf
	Return 1

EndFunc   ;==>_XLChart_FillSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_FontSet
; Description....: Set font properties for the specified object.
; Syntax.........: _XLChart_FontSet($oObject[, $sName = Default[, $iSize = Default[, $bBold = Default[, $bItalic = Default[, $bUnderline = Default[, $iColor = Default,[$bThemeColor = False]]]]]]])
; Parameters ....: $oObject     - Object for which the font properties should be set (ChartTitle, AxisTitle, Legend ...)
;                  $sName       - Optional: Font name like "Courier New" or "Arial" (default = Default)
;                  $iSize       - Optional: Size of the font in points (default = Default)
;                  $bBold       - Optional: If True the font will be displayed bold (default = Default)
;                  $bItalic     - Optional: If True the font will be displayed italic (default = Default)
;                  $bUnderline  - Optional: If True the font will be displayed underlined (default = Default)
;                  $iColor      - Optional: Color of the font (default = Default)
;                  +              You can set colors to an explicit red-green-blue value (e.g. 0xFF00FF) or to a color in the color scheme (negative numbers -1 to -56)
;                  $bThemeColor - Optional: True specifies that $iColor is interpreted as theme color (default = False).
;                  +              If set to True the $iColor value has to be one of the MsoThemeColorIndex enumeration
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is no object
;                  |2 - $iSize is not a number
;                  |3 - $bBold is not boolean
;                  |4 - $bItalic is not boolean
;                  |5 - $bUnderline is not boolean
;                  |6 - $iColor is not an integer
;                  |7 - $bThemeColor is not boolean
; Authors........: water
; Modified ......: GreenCan
; Remarks .......: Color 0 (white) has to be specified as RGB value or use color 2 which has the same RGB value of 0xFFFFFF
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_FontSet($oObject, $sName = Default, $iSize = Default, $bBold = Default, $bItalic = Default, $bUnderline = Default, $iColor = Default, $bThemeColor = False)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $bThemeColor = Default Then $bThemeColor = False
	If $iSize <> Default And Not IsNumber($iSize) Then Return SetError(2, 0, 0)
	If $bBold <> Default And Not IsBool($bBold) Then Return SetError(3, 0, 0)
	If $bItalic <> Default And Not IsBool($bItalic) Then Return SetError(4, 0, 0)
	If $bUnderline <> Default And Not IsBool($bUnderline) Then Return SetError(5, 0, 0)
	If $iColor <> Default And Not IsInt($iColor) Then Return SetError(6, 0, 0)
	If Not IsBool($bThemeColor) Then Return SetError(7, 0, 0)
	If $sName <> Default Then $oObject.Font.Name = $sName
	If $iSize <> Default Then $oObject.Font.Size = $iSize
	If $bBold <> Default Then $oObject.Font.Bold = $bBold
	If $bItalic <> Default Then $oObject.Font.Italic = $bItalic
	If $bUnderline <> Default Then $oObject.Font.Underline = $bUnderline
	If $iColor <> Default Then
		If $bThemeColor Then
			$oObject.Font.ThemeColor = $iColor
		Else
			If $iColor < 0 Then
				$oObject.Font.ColorIndex = Abs($iColor)
			Else
				$oObject.Font.Color = _XLChart_RGB($iColor)
			EndIf
		EndIf
	EndIf
	Return 1

EndFunc   ;==>_XLChart_FontSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_GridSet
; Description....: Set gridlines of a chart.
; Syntax.........: _XLChart_GridSet($oObject[, $iMajorGridline = Default[, $iMinorGridline = Default]])
; Parameters ....: $oObject        - Object of the axis for which the grid properties should be set
;                  $bMajorGridline - Optional: True if the axis has major gridlines. Only axes in the primary axis group can have gridlines (default = Default)
;                  $bMinorGridline - Optional: True if the axis has minor gridlines. Only axes in the primary axis group can have gridlines (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is no object
;                  |2 - $iMajorGridline is not boolean
;                  |3 - $iMinorGridline is not boolean
; Authors........: water
; Modified ......: GreenCan
; Remarks .......: To change the units between major and/or minor grid lines please see _XLChart_TicksSet
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_GridSet($oObject, $bMajorGridline = Default, $bMinorGridline = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $bMajorGridline <> Default And Not IsBool($bMajorGridline) Then Return SetError(2, 0, 0)
	If $bMinorGridline <> Default And Not IsBool($bMinorGridline) Then Return SetError(3, 0, 0)
	If $bMajorGridline <> Default Then $oObject.HasMajorGridlines = $bMajorGridline
	If $bMinorGridline <> Default Then $oObject.HasMinorGridlines = $bMinorGridline
	Return 1

EndFunc   ;==>_XLChart_GridSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_LayoutSet
; Description....: Set layout, style or template for a chart.
; Syntax.........: _XLChart_LayoutSet($oChart[, $iLayout = Default[, $iStyle = Default[, $iTemplate = ""]]])
; Parameters ....: $oChart    - Chart object as returned by _XLChart_Create()
;                  $iLayout   - Applies one of the layouts shown in the ribbon. Number between 1 and about 11
;                  $iStyle    - Chart style for the chart. Number between 1 and 48
;                  $iTemplate - Applies a standard or custom template to a chart (path and filename)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oChart is no object
;                  |2 - Layout could not be applied. Please check @extended for detailed error code
;                  |3 - Style could not be applied. Please check @extended for detailed error code
;                  |4 - Template could not be applied. Please check @extended for detailed error code
; Authors........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_LayoutSet($oChart, $iLayout = Default, $iStyle = Default, $sTemplate = "")

	If Not IsObj($oChart) Then Return SetError(1, 0, 0)
	If $iLayout <> Default Then
		$oChart.Applylayout($iLayout)
		If @error Then Return SetError(2, @error, 0)
	EndIf
	If $iStyle <> Default Then
		$oChart.ClearToMatchStyle()
		$oChart.ChartStyle = $iStyle
		If @error Then Return SetError(3, @error, 0)
	EndIf
	If $sTemplate <> "" Then
		$oChart.ApplyChartTemplate($sTemplate)
		If @error Then Return SetError(4, @error, 0)
	EndIf
	Return 1

EndFunc   ;==>_XLChart_LayoutSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_LegendSet
; Description....: Set properties of the legend.
; Syntax.........: _XLChart_LegendSet($oChart[, $iPosition = Default[, $iLeft = Default[, $iTop = Default[, $iWidth = Default[, $iHeight = Default[, $bShadow = Default]]]]]])
; Parameters ....: $oChart    - Chart for which the legend properties should be set
;                  $iPosition - Optional: Represents the position of the legend on the chart. Please see the XlLegendPosition enumeration (default = Default)
;                               -1 = Do not show a legend. All other parameters are ignored
;                  $iLeft     - Optional: Value in points representing the distance from the left edge of the legend to the left edge of the chart area (default = Default)
;                  $iTop      - Optional: Value in points representing the distance from the top edge of the legend to the top of the chart area (default = Default)
;                  $iWidth    - Optional: Value in points representing the width of the legend (default = Default)
;                  $iHeight   - Optional: Value in points representing the height of the legend (default = Default)
;                  $bFrame    - Optional: True if the legend has a frame (default = True)
;                  $bShadow   - Optional: True if the legend has a shadow (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oChart is no object
;                  |2 - $iLeft is not an integer value
;                  |3 - $iTop is not an integer value
;                  |4 - $iWidth is not an integer value
;                  |5 - $iHeight is not an integer value
;                  |6 - $bShadow is not a boolean value
;                  |7 - $iPosition is not an integer value
;                  |8 - $bFrame is not a boolean value
;                  |9 - Error setting the legend position (position, left or top). Please check @extended for detailed error code
;                  |10 - Error setting the legend width. Please check @extended for detailed error code
;                  |11 - Error setting the legend height. Please check @extended for detailed error code
;                  |12 - Error setting the legend frame. Please check @extended for detailed error code
;                  |13 - Error setting the legend shadow. Please check @extended for detailed error code
; Authors........: water
; Modified ......:
; Remarks .......: If $iPosition is set parameters $iLeft and $iTop will be ignored
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_LegendSet($oChart, $iPosition = Default, $iLeft = Default, $iTop = Default, $iWidth = Default, $iHeight = Default, $bFrame = Default, $bShadow = Default)

	If Not IsObj($oChart) Then Return SetError(1, 0, 0)
	If $iPosition <> Default Then
		If Not IsInt($iPosition) Then Return SetError(7, 0, 0)
		If $iPosition = -1 Then
			$oChart.HasLegend = False
			Return 1
		Else
			$oChart.HasLegend = True
			$oChart.Legend.Position = $iPosition
			If @error <> 0 Then Return SetError(9, @error, 0)
		EndIf
	Else
		If $iLeft <> Default Then
			If Not IsInt($iLeft) Then Return SetError(2, 0, 0)
			$oChart.Legend.Left = $iLeft
			If @error <> 0 Then Return SetError(9, @error, 0)
		EndIf
		If $iTop <> Default Then
			If Not IsInt($iTop) Then Return SetError(3, 0, 0)
			$oChart.Legend.Top = $iTop
			If @error <> 0 Then Return SetError(9, @error, 0)
		EndIf
	EndIf
	If $iWidth <> Default Then
		If Not IsInt($iWidth) Then Return SetError(4, 0, 0)
		$oChart.Legend.Width = $iWidth
		If @error <> 0 Then Return SetError(10, @error, 0)
	EndIf
	If $iHeight <> Default Then
		If Not IsInt($iHeight) Then Return SetError(5, 0, 0)
		$oChart.Legend.Height = $iHeight
		If @error <> 0 Then Return SetError(11, @error, 0)
	EndIf
	If $bFrame <> Default Then
		If Not IsBool($bFrame) Then Return SetError(8, 0, 0)
		$oChart.Legend.Format.Line.Visible = $bFrame
		If @error <> 0 Then Return SetError(12, @error, 0)
	EndIf
	If $bShadow <> Default Then
		If Not IsBool($bShadow) Then Return SetError(6, 0, 0)
		$oChart.Legend.Shadow = $bShadow
		If @error <> 0 Then Return SetError(13, @error, 0)
	EndIf
	Return 1

EndFunc   ;==>_XLChart_LegendSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_LineGet
; Description....: Get properties of a line (axis line, grid line, data line ...).
; Syntax.........: _XLChart_LineGet($oObject)
; Parameters ....: $oObject - Object for which the properties should be returned (axis, grid line, data line ...)
; Return values .: Success - Returns a one-based one dimensional array with the following properties:
;                  |1 - Weight: Integer that represents the weight of the line in points
;                  |2 - Color
;                  |3 - ThemeColor
;                  |4 - Style: Integer that represents the style of the line. MsoLineStyle enumeration
;                  |5 - DashStyle: Integer that represents the dash style for the line. MsoLineDashStyle enumeration
;                  |6 - Visible
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is not an object
; Authors........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_LineGet($oObject)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	Local $aProperties[7] = [6]
	$aProperties[1] = $oObject.Format.Line.Weight
	$aProperties[2] = _XLChart_RGB($oObject.Border.Color)
	$aProperties[3] = $oObject.Border.ThemeColor
	$aProperties[4] = $oObject.Format.Line.Style
	$aProperties[5] = $oObject.Format.Line.DashStyle
	$aProperties[6] = $oObject.Format.Line.Visible
	Return $aProperties

EndFunc   ;==>_XLChart_LineGet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_LineGroupSet
; Description....: Set properties of a line chart group.
; Syntax.........: _XLChart_LineGroupSet($oObject[, $bHasUpDownBars = Default[, $bHasHiLoLines = Default[, $bHasDropLines = Default]]])
; Parameters ....: $oObject        - Chart group for which the properties should be set
;                  $bHasUpDownBars - Optional: True if the line chart has up and down bars (default = Default)
;                  $bHasHiLoLines  - Optional: True if the line chart has high-low lines (default = Default)
;                  $bHasDropLines  - Optional: True if the line chart has drop lines (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is not an object
;                  |2 - $bHasUpDownBars is not boolean
;                  |3 - $bHasHiLoLines is not boolean
;                  |4 - $bHasDropLines is not boolean
; Authors........: water
; Modified ......:
; Remarks .......: A chart contains one or more chart groups, each chart group contains one or more series, and
;                  each series contains one or more points.
;+
;                  You can either pass an item of the ChartGroups collection or an item of the LineGroups collection (a ChartGroup object)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_LineGroupSet($oObject, $bHasUpDownBars = Default, $bHasHiLoLines = Default, $bHasDropLines = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $bHasUpDownBars <> Default And Not IsBool($bHasUpDownBars) Then Return SetError(2, 0, 0)
	If $bHasHiLoLines <> Default And Not IsBool($bHasHiLoLines) Then Return SetError(3, 0, 0)
	If $bHasDropLines <> Default And Not IsBool($bHasDropLines) Then Return SetError(4, 0, 0)
	If $bHasUpDownBars <> Default Then $oObject.HasUpDownBars = $bHasUpDownBars
	If $bHasHiLoLines <> Default Then $oObject.HasHiLoLines = $bHasHiLoLines
	If $bHasDropLines <> Default Then $oObject.HasDropLines = $bHasDropLines
	Return 1

EndFunc   ;==>_XLChart_LineGroupSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_LineSet
; Description....: Set properties of a line (axis line, grid line, data line ...).
; Syntax.........: _XLChart_LineSet($oObject[, $iWeight = Default[, $iColor = Default[, $bThemeColor = False[, $iStyle = Default[, $iDash = Default[, $iVisible = Default]]]]]])
; Parameters ....: $oObject     - Object for which the properties should be set (axis, grid line, data line ...)
;                  $iWeight     - Optional: Integer that represents the weight of the line (default = Default)
;                  $iColor      - Optional: Color of the line (default = Default)
;                  +              You can set colors to an explicit red-green-blue value (e.g. 0xFF00FF) or to a color in the color scheme (negative numbers -1 to -56)
;                  $bThemeColor - Optional: True specifies that $iColor is interpreted as theme color (default = False).
;                  +              If set to True the $iColor value has to be one of the MsoThemeColorIndex enumeration
;                  $iStyle      - Optional: Integer that represents the style of the line. Please check the MsoLineStyle enumeration (default = Default)
;                  $iDash       - Optional: Integer that represents the dash style for the line. Please check the MsoLineDashStyle enumeration (default = Default)
;                  $iVisible    - Optional: Determines whether the line is visible. Please check the XlSheetVisibility enumeration.
;                  +              Can be $xlSheetHidden or $xlSheetVisible (default = $xlSheetVisible)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is no object
;                  |2 - $iStyle and $iDash can't be used together
;                  |3 - $iWeight is not an integer
;                  |4 - $iColor is not an integer
;                  |5 - $bThemeColor is not boolean
;                  |6 - $iStyle is not an integer
;                  |7 - $iDash is not an integer
;                  |8 - $iVisible is not an integer
; Authors........: water
; Modified ......:
; Remarks .......: You either set $iStyle or $iDash but not both.
;+                 Color 0 (white) has to be specified as RGB value or use color 2 which has the same RGB value of 0xFFFFFF
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_LineSet($oObject, $iWeight = Default, $iColor = Default, $bThemeColor = False, $iStyle = Default, $iDash = Default, $iVisible = $xlSheetVisible)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $bThemeColor = Default Then $bThemeColor = False
	If $iStyle <> Default And $iDash <> Default Then Return SetError(2, 0, 0)
	If $iWeight <> Default And Not IsInt($iWeight) Then Return SetError(3, 0, 0)
	If $iColor <> Default And Not IsInt($iColor) Then Return SetError(4, 0, 0)
	If Not IsBool($bThemeColor) Then Return SetError(5, 0, 0)
	If $iStyle <> Default And Not IsInt($iStyle) Then Return SetError(6, 0, 0)
	If $iDash <> Default And Not IsInt($iDash) Then Return SetError(7, 0, 0)
	If $iVisible <> Default And Not IsInt($iVisible) Then Return SetError(8, 0, 0)
	If $iVisible <> Default Then $oObject.Format.Line.Visible = $iVisible
	If $iColor <> Default Then ; apply color first
		If $bThemeColor Then
			$oObject.Border.ThemeColor = $iColor
		Else
			If $iColor < 0 Then
				$oObject.Border.ColorIndex = Abs($iColor)
			Else
				$oObject.Border.Color = _XLChart_RGB($iColor)
			EndIf
		EndIf
	EndIf
	If $iWeight <> Default Then $oObject.Format.Line.Weight = $iWeight
	If $iStyle <> Default Then $oObject.Format.Line.Style = $iStyle
	If $iDash <> Default Then $oObject.Format.Line.DashStyle = $iDash
	Return 1

EndFunc   ;==>_XLChart_LineSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_MarkerSet
; Description....: Set properties for the marker objects of line, scatter or radar charts.
; Syntax.........: _XLChart_MarkerSet($oObject[, $iSize = Default[, $iStyle = Default[, $iForeColor = Default[, $iBackColor = Default[, $bThemeColor = False]]]]])
; Parameters ....: $oObject     - Object for which the fill properties should be set (...)
;                  $iSize       - Optional: Size of the data-marker in points. Can be a value from 2 through 72 (default = Default)
;                  $iStyle      - Optional: Marker style. Can be any of the xlMarkerStyle enumeration (default = Default)
;                  $iForeColor  - Optional: Sets the foreground fill color (default = Default)
;                  +              You can set colors to an explicit red-green-blue value (e.g. 0xFF00FF) or to a color in the color scheme (negative numbers -1 to -56)
;                  $iBackColor  - Optional: Sets the background fill color (default = Default)
;                  +              You can set colors to an explicit red-green-blue value (e.g. 0xFF00FF) or to a color in the color scheme (negative numbers -1 to -56)
;                  $bThemeColor - Optional: True specifies that $iForeColor and $iBackColor are interpreted as theme colors (default = False).
;                  +              If set to True $iForeColor and $iBackColor values have to be one of the MsoThemeColorIndex enumeration
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is no object
;                  |2 - $iForeColor is not an integer
;                  |3 - $iBackColor is not an integer
;                  |4 - $iSize is not an integer or < 2 or > 72
;                  |5 - $iStyle is not an integer
;                  |9 - $bThemeColor is not boolean
; Authors........: water
; Modified ......:
; Remarks .......: Color 0 (white) has to be specified as RGB value or use color 2 which has the same RGB value of 0xFFFFFF
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_MarkerSet($oObject, $iSize = Default, $iStyle = Default, $iForeColor = Default, $iBackColor = Default, $bThemeColor = False)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $bThemeColor = Default Then $bThemeColor = False
	If Not IsBool($bThemeColor) Then Return SetError(9, 0, 0)
	If $iSize <> Default And (Not (IsInt($iSize)) Or $iSize < 2 Or $iSize > 72) Then Return SetError(4, 0, 0)
	If $iStyle <> Default And Not IsInt($iStyle) Then Return SetError(5, 0, 0)
	If $iSize <> Default Then $oObject.MarkerSize = $iSize
	If $iStyle <> Default Then $oObject.MarkerStyle = $iStyle
	If $iForeColor <> Default Then
		If Not IsInt($iForeColor) Then Return SetError(2, 0, 0)
		If $iForeColor < 0 Then
			$oObject.MarkerForegroundColorIndex = Abs($iForeColor)
		Else
			$oObject.MarkerForegroundColor = _XLChart_RGB($iForeColor)
		EndIf
	EndIf
	If $iBackColor <> Default Then
		If Not IsInt($iBackColor) Then Return SetError(3, 0, 0)
		If $iBackColor < 0 Then
			$oObject.MarkerBackgroundColorIndex = Abs($iBackColor)
		Else
			$oObject.MarkerBackgroundColor = _XLChart_RGB($iBackColor)
		EndIf
	EndIf

	Return 1

EndFunc   ;==>_XLChart_MarkerSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ObjectDelete
; Description....: Delete an object from a chart.
; Syntax.........: _XLChart_ObjectDelete($oObject)
; Parameters ....: $oObject - Object of the data series to be deleted
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is not an object
;                  |2 - Unable to delete the specified object. Please check @extended for detailed error code
; Authors........: GreenCan
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ObjectDelete($oObject)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	$oObject.Delete()
	If @error Then Return SetError(2, @error, 0)
	Return 1

EndFunc   ;==>_XLChart_ObjectDelete

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ObjectPositionSet
; Description....: Resize and reposition an object (plot area, legend ...) on a chart.
; Syntax.........: _XLChart_ObjectPositionSet($oObject[, $iLeft = Default[, $iTop = Default[, $iWidth = Default[, $iHeight = Default[, $iFlag = 0]]]]])
; Parameters ....: $oObject  - Chart object as returned by a preceding call to _XLChart_ChartCreate
;                  $iLeft    - Optional: Distance, in points, from the left edge of the chart (default = Default)
;                  $iTop     - Optional: Distance, in points, from the top edge of the chart (default = Default)
;                  $iWidth   - Optional: Width, in points, of the object (default = Default)
;                  $iHeight  - Optional: Height, in points, of the object default = Default)
;                  $iFlag    - Optional: 1 = add the left/top/widht/height value to the current value, 0 = set the left/top/widht/height value (default = 0)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is not an object
;                  |2 - $iLeft is not a number
;                  |3 - $iTop is not a number
;                  |4 - $iWidth is not a number
;                  |5 - $iHeight is not a number
;                  |6 - $iFlag is not an integer or < 0 or > 1
; Authors........: GreenCan, water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ObjectPositionSet($oObject, $iLeft = Default, $iTop = Default, $iWidth = Default, $iHeight = Default, $iFlag = 0)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $iLeft <> Default And Not IsNumber($iLeft) Then Return SetError(2, 0, 0)
	If $iTop <> Default And Not IsNumber($iTop) Then Return SetError(3, 0, 0)
	If $iWidth <> Default And Not IsNumber($iWidth) Then Return SetError(4, 0, 0)
	If $iHeight <> Default And Not IsNumber($iHeight) Then Return SetError(5, 0, 0)
	If Not IsInt($iFlag) Or $iFlag < 0 Or $iFlag > 1 Then Return SetError(6, 0, 0)
	If $iFlag = 1 Then
		If $iLeft <> Default Then $oObject.left = $oObject.left + $iLeft
		If $iTop <> Default Then $oObject.Top = $oObject.Top + $iTop
		If $iWidth <> Default Then $oObject.Width = $oObject.Width + $iWidth
		If $iHeight <> Default Then $oObject.Height = $oObject.Height + $iHeight
	Else
		If $iLeft <> Default Then $oObject.left = $iLeft
		If $iTop <> Default Then $oObject.Top = $iTop
		If $iWidth <> Default Then $oObject.Width = $iWidth
		If $iHeight <> Default Then $oObject.Height = $iHeight
	EndIf
	Return 1

EndFunc   ;==>_XLChart_ObjectPositionSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_OfPieGroupSet
; Description....: Set properties of a pie of pie or bar of pie chart group.
; Syntax.........: _XLChart_OfPieGroupSet($oObject[, $bHasSeriesLines = Default[, $iGapWidth = Default[, $iSecondPlotSize = Default[, $iSplitType = Default[, $iSplitValue = Default]]]]])
; Parameters ....: $oObject         - Chart group for which the properties should be set
;                  $bHasSeriesLines - True if a Pie of Pie or Bar of Pie chart has connector lines between the two sections (default = Default)
;                  $iGapWidth       - The space between the primary and secondary sections of the chart (default = Default)
;                  $iSecondPlotSize - The size of the secondary section as a percentage of the size of the primary pie. Can be a value from 5 to 200 (default = Default)
;                  $iSplitType      - The way the two sections are split. Can be any of the XlChartSplitType enumeration (default = Default)
;                  $iSplitValue     - The threshold value separating the two sections of the chart (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is not an object
;                  |2 - $bHasSeriesLines is not boolean
;                  |3 - $iGapWidth is not an integer
;                  |4 - $iSecondPlotSize is not an integer or y 5 or > 200
;                  |5 - $iSplitType is not an integer
;                  |6 - $iSplitValue is not an integer
; Authors........: water
; Modified ......:
; Remarks .......: A chart contains one or more chart groups, each chart group contains one or more series, and
;                  each series contains one or more points.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_OfPieGroupSet($oObject, $bHasSeriesLines = Default, $iGapWidth = Default, $iSecondPlotSize = Default, $iSplitType = Default, $iSplitValue = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $bHasSeriesLines <> Default And Not IsBool($bHasSeriesLines) Then Return SetError(2, 0, 0)
	If $iGapWidth <> Default And Not IsInt($iGapWidth) Then Return SetError(3, 0, 0)
	If $iSecondPlotSize <> Default And (Not (IsInt($iSecondPlotSize)) Or $iSecondPlotSize < 5 Or $iSecondPlotSize > 200) Then _
			Return SetError(4, 0, 0)
	If $iSplitType <> Default And Not IsInt($iSplitType) Then Return SetError(5, 0, 0)
	If $iSplitValue <> Default And Not IsInt($iSplitValue) Then Return SetError(6, 0, 0)
	If $bHasSeriesLines <> Default Then $oObject.HasSeriesLines = $bHasSeriesLines
	If $iGapWidth <> Default Then $oObject.GapWidth = $iGapWidth
	If $iSecondPlotSize <> Default Then $oObject.SecondPlotSize = $iSecondPlotSize
	If $iSplitType <> Default Then $oObject.SplitType = $iSplitType
	If $iSplitValue <> Default Then $oObject.SplitValue = $iSplitValue
	Return 1

EndFunc   ;==>_XLChart_OfPieGroupSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_PageSet
; Description....: Set the page setup attributes (paper size, orientation, margins etc.) for a chart or chartsheet.
; Syntax.........: _XLChart_PageSet($oChart[, $iPaperSize = Default[, $iOrientation = Default[, $iTopMargin = Default[, $iBottomMargin = Default[, $iLeftMargin = Default[, $iRightMargin = Default[, $iBlackAndWhite = Default]]]]]]])
; Parameters ....: $oChart         - Chart object or chartsheet object for which the page setup properties should be set
;                  $iPaperSize     - Optional: Sets the size of the paper. See enumeration XlPaperSize (default = Default)
;                  $iOrientation   - Optional: Sets a XlPageOrientation value that represents the portrait or landscape printing mode (default = Default)
;                  $iTopMargin     - Optional: Sets the top margin in centimeters (default = Default)
;                  $iBottomMargin  - Optional: Sets the bottom margin in centimeters (default = Default)
;                  $iLeftMargin    - Optional: Sets the left margin in centimeters (default = Default)
;                  $iRightMargin   - Optional: Sets the right margin in centimeters (default = Default)
;                  $bBlackAndWhite - Optional: True if the chart will be printed in black and white (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oChart is not an object
;                  |2 - $iPaperSize is invalid. Must be an integer
;                  |3 - $iOrientation is invalid. Must be an integer
;                  |4 - $iTopMargin is invalid. Must be numeric
;                  |5 - $iBottomMargin is invalid. Must be numeric
;                  |6 - $iLeftMargin is invalid. Must be numeric
;                  |7 - $iRightMargin is invalid. Must be numeric
;                  |8 - $bBlackAndWhite is not boolean
; Authors........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_PageSet($oChart, $iPaperSize = Default, $iOrientation = Default, $iTopMargin = Default, $iBottomMargin = Default, $iLeftMargin = Default, $iRightMargin = Default, $bBlackAndWhite = Default)

	If Not IsObj($oChart) Then Return SetError(1, 0, 0)
	If $iPaperSize <> Default And Not IsInt($iPaperSize) Then Return SetError(2, 0, 0)
	If $iOrientation <> Default And Not IsInt($iOrientation) Then Return SetError(3, 0, 0)
	If $iTopMargin <> Default And Not IsNumber($iTopMargin) Then Return SetError(4, 0, 0)
	If $iBottomMargin <> Default And Not IsNumber($iBottomMargin) Then Return SetError(5, 0, 0)
	If $iLeftMargin <> Default And Not IsNumber($iLeftMargin) Then Return SetError(6, 0, 0)
	If $iRightMargin <> Default And Not IsNumber($iRightMargin) Then Return SetError(7, 0, 0)
	If $bBlackAndWhite <> Default And Not IsBool($bBlackAndWhite) Then Return SetError(8, 0, 0)
	$oChart.Parent.Activate ; Activate the chartobject
	If $iPaperSize <> Default Then $oChart.PageSetup.PaperSize = $iPaperSize
	If $iOrientation <> Default Then $oChart.PageSetup.Orientation = $iOrientation
	If $iTopMargin <> Default Then $oChart.PageSetup.TopMargin = $iTopMargin / 0.035 ; centimeters -> points
	If $iBottomMargin <> Default Then $oChart.PageSetup.BottomMargin = $iBottomMargin / 0.035
	If $iLeftMargin <> Default Then $oChart.PageSetup.leftMargin = $iLeftMargin / 0.035
	If $iRightMargin <> Default Then $oChart.PageSetup.RightMargin = $iRightMargin / 0.035
	If $bBlackAndWhite <> Default Then $oChart.PageSetup.BlackAndWhite = $bBlackAndWhite
	Return 1

EndFunc   ;==>_XLChart_PageSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_PieGroupSet
; Description....: Set properties of a pie or 3D-pie chart group.
; Syntax.........: _XLChart_PieGroupSet($oObject[, $iFirstSliceAngle = Default])
; Parameters ....: $oObject          - Chart group for which the properties should be set
;                  $iFirstSliceAngle - Optional: Angle of the first pie-chart slice in degrees (clockwise from vertical).
;                  +Can be a value from 0 through 360 (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is not an object
;                  |2 - $iFirstSliceAngle is not an integer or < 0 or > 360
; Authors........: water
; Modified ......:
; Remarks .......: A chart contains one or more chart groups, each chart group contains one or more series, and
;                  each series contains one or more points.
;+
;                 For 2D-Pies you can either pass an item of the ChartGroups collection or an item of the PieGroups collection (a ChartGroup object).
;                 For 3D-Pies you have to use an item of the ChartGroups collection.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_PieGroupSet($oObject, $iFirstSliceAngle = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $iFirstSliceAngle <> Default And (Not (IsInt($iFirstSliceAngle)) Or $iFirstSliceAngle < 0 Or $iFirstSliceAngle > 360) Then _
			Return SetError(2, 0, 0)
	If $iFirstSliceAngle <> Default Then $oObject.FirstSliceAngle = $iFirstSliceAngle
	Return 1

EndFunc   ;==>_XLChart_PieGroupSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ScreenUpdateSet
; Description....: Turning screen updating on/off to improve performance.
; Syntax.........: _XLChart_ScreenUpdateSet($oObject[, $iScreenUpdate = 0])
; Parameters ....: $oExcel        - Excel object opened by a preceding call to _Excel_BookOpen() or _Excel_BookNew()
;                  $iScreenUpdate - Optional: Enables/Disables the Excel screen updating during chart creation to enhance performance (default = 0)
;                  |0 - Disable screen updating. Is ignored when debugging is activated ($g__iDebug > 0)
;                  |1 - Enable screen updating
;                  |2 - Disable screen updating (forced). Disables screen updating even when debugging is activated ($g__iDebug > 0)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oExcel is not an object
;                  |2 - $iScreenUpdate is invalid. Must be 0, 1 or 2
;                  |3 - Debugging enabled, setting overrules ScreenUpdate. ScreenUpdate not disabled
;                  |4 - Error updating ScreenUpdating property. See @extended for details
; Authors........: water
; Modified ......:
; Remarks .......: Remember to enable ScreenUpdating after creation of your charts.
;                  This function is valid not only for the Excel Chart UDF but is a general Excel function.
;+
;                  If you get wierd results with Excel 2007 please try again and set $iScreenUpdate = 1
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ScreenUpdateSet($oExcel, $iScreenUpdate = 0)

	If Not IsObj($oExcel) Then Return SetError(1, 0, 0)
	If Not IsInt($iScreenUpdate) Or $iScreenUpdate < 0 Or $iScreenUpdate > 2 Then Return SetError(2, 0, 0)
	If $g__iDebug <> 0 Then
		If $iScreenUpdate = 0 Then Return SetError(3, 0, 0)
		If $iScreenUpdate = 1 Then $oExcel.ScreenUpdating = True
		If $iScreenUpdate = 2 Then $oExcel.ScreenUpdating = False
	Else
		$oExcel.ScreenUpdating = $iScreenUpdate
	EndIf
	If @error Then Return SetError(4, @error, 0)
	Return 1

EndFunc   ;==>_XLChart_ScreenUpdateSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_SeriesAdd
; Description....: Add a data series to a chart.
; Syntax.........: _XLChart_SeriesAdd($oChart, $sXValueRange, $vDataRange, $vDataName)
; Parameters ....: $oChart       - Chart object as returned by a preceding call to _XLChart_ChartCreate
;                  $sXValueRange - Category (X) axis label range always a single range (eg. "=Sheet1!R2C1:R6C1")
;                  $vDataRange   - The values range. Either a single range or an one-dimensional one based array
;                  $vDataName    - Header name of the range. Either a single range or an one-dimensional one based array
; Return values .: Success - Object identifier of the created data series
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oChart is not an object
;                  |2 - Unable to set the Data Range of the new series
;                  |3 - Unable to set the Value Range of the new series
;                  |4 - Unable to set the Data Name of the new series
; Authors........: GreenCan
; Modified ......: water
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_SeriesAdd($oChart, $sXValueRange, $vDataRange, $vDataName)

	If Not IsObj($oChart) Then Return SetError(1, 0, 0)
	Local $oNewSeries = $oChart.SeriesCollection.NewSeries()
	With $oNewSeries
		.Values = $vDataRange
		If @error Then Return SetError(2, @error, 0)
		.XValues = $sXValueRange
		If @error Then Return SetError(3, @error, 0)
		.Name = $vDataName
		If @error Then Return SetError(4, @error, 0)
	EndWith
	Return $oNewSeries

EndFunc   ;==>_XLChart_SeriesAdd

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_SeriesSet
; Description....: Set properties of a data series.
; Syntax.........: _XLChart_SeriesSet($oObject[, $iChartType = Default[, $bSmooth = Default[, $bSecondary = Default]]])
; Parameters ....: $oObject    - Object of the data series for which the properties should be set
;                  $iChartType - Optional: Sets the chart type. Can be one of the XlChartType enumeration (default = Default)
;                  $bSmooth    - Optional: True if curve smoothing is turned on. Only valid for line and scatter charts (default = Default)
;                  $iSecondary - Optional: Sets the axis group. Can be one of the XlAxisGroup enumeration (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is no object
;                  |2 - $iChartType is not a number
;                  |3 - $iSmooth is not boolean
;                  |4 - $iSecondary is not an integer
; Authors........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_SeriesSet($oObject, $iChartType = Default, $bSmooth = Default, $iSecondary = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $iChartType <> Default And Not IsNumber($iChartType) Then Return SetError(2, 0, 0)
	If $bSmooth <> Default And Not IsBool($bSmooth) Then Return SetError(3, 0, 0)
	If $iSecondary <> Default And Not IsInt($iSecondary) Then Return SetError(4, 0, 0)
	If $iChartType <> Default Then $oObject.ChartType = $iChartType
	If $bSmooth <> Default Then $oObject.Smooth = $bSmooth
	If $iSecondary <> Default Then $oObject.AxisGroup = $iSecondary
	Return 1

EndFunc   ;==>_XLChart_SeriesSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_ShadowSet
; Description....: Set properties of a shadow.
; Syntax.........: _XLChart_ShadowSet($oObject[, $iForeColor = Default[, $bThemeColor = False[, $iOffsetX = Default[, $iOffsetY = Default[, $iTransparency = Default[, $iVisible = $xlSheetVisible]]]]]])
; Parameters ....: $oObject       - Object to apply the shadow properties to (e.g. $oChart or $oChart.Axes($xlCategory))
;                  $iForeColor    - Optional: Sets the foreground fill color (default = Default)
;                  +                You can set colors to an explicit red-green-blue value (e.g. 0xFF00FF) or to a color in the color scheme (negative numbers -1 to -56)
;                  $bThemeColor   - Optional: True specifies that $iForeColor is interpreted as theme color (default = False).
;                  +                If set to True the $iForeColor value has to be one of the MsoThemeColorIndex enumeration
;                  $iStyle        - Optional: Shadow style. Can be any of the MsoShadowStyle enumeration (default = Default)
;                  $iBlur         - Optional: sets the degree of blurriness of the specified shadow (default = Default)
;                  $iOffsetX      - Optional: The horizontal offset of the shadow from the specified shape, in points (default = Default).
;                  +                A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left
;                  $iOffsetY      - Optional: The vertical offset of the shadow from the specified shape, in points (default = Default).
;                  +                A positive value offsets the shadow to the top of the shape; a negative value offsets it to the bottom
;                  $iTransparency - Optional: The degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear) (default = Default)
;                  $iVisible      - Optional: Determines whether the shadow is visible. Please check the XlSheetVisibility enumeration.
;                  +                Can be $xlSheetHidden or $xlSheetVisible (default = $xlSheetVisible)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is no object
;                  |2 - $iForeColor is not an integer
;                  |3 - $bThemeColor is not boolean
;                  |4 - $iStyle is not an integer
;                  |5 - $iBlur is not an integer
;                  |6 - $sOffsetX is not a number
;                  |7 - $sOffsetY is not a number
;                  |8 - $iTransparency is not a number
;                  |9 - $bVisible is not an integer
; Authors........: water
; Modified ......: GreenCan
; Remarks .......: Excel 2007: Just draws a shadow on the border. So you have to make sure the object (chart title, legend etc.) has a border drawn. Else you won't see a shadow.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_ShadowSet($oObject, $iForeColor = Default, $bThemeColor = False, $iStyle = Default, $iBlur = Default, $iOffsetX = Default, $iOffsetY = Default, $iTransparency = Default, $iVisible = $xlSheetVisible)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $iForeColor <> Default And Not IsInt($iForeColor) Then Return SetError(2, 0, 0)
	If $bThemeColor <> Default And Not IsBool($bThemeColor) Then Return SetError(3, 0, 0)
	If $iStyle <> Default And Not IsNumber($iStyle) Then Return SetError(4, 0, 0)
	If $iBlur <> Default And Not IsNumber($iBlur) Then Return SetError(5, 0, 0)
	If $iOffsetX <> Default And Not IsNumber($iOffsetX) Then Return SetError(6, 0, 0)
	If $iOffsetY <> Default And Not IsNumber($iOffsetY) Then Return SetError(7, 0, 0)
	If $iTransparency <> Default And Not IsNumber($iTransparency) Then Return SetError(8, 0, 0)
	If $iVisible <> Default And Not IsInt($iVisible) Then Return SetError(9, 0, 0)
	If $iStyle <> Default Then $oObject.Format.Shadow.Style = $iStyle
	If $iForeColor <> Default Then
		If $bThemeColor Then
			$oObject.Format.Shadow.ForeColor.ObjectThemeColor = $iForeColor
		Else
			If $iForeColor < 0 Then
				; Add 7 to ColorIndex to convert to SchemeColor: http://www.ozgrid.com/forum/showthread.php?t=53791
				$oObject.Format.Shadow.ForeColor.SchemeColor = Abs($iForeColor) + 7
			Else
				$oObject.Format.Shadow.ForeColor.RGB = _XLChart_RGB($iForeColor)
			EndIf
		EndIf
	EndIf
	If $iOffsetX <> Default Then $oObject.Format.Shadow.OffsetX = $iOffsetX
	If $iOffsetY <> Default Then $oObject.Format.Shadow.OffsetY = $iOffsetY
	If $iBlur <> Default Then $oObject.Format.Shadow.Blur = $iBlur
	If $iTransparency <> Default Then $oObject.Format.Shadow.Transparency = $iTransparency
	If $iVisible <> Default Then $oObject.Format.Shadow.Visible = $iVisible
	Return 1

EndFunc   ;==>_XLChart_ShadowSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_TicksSet
; Description....: Set tick marks and tick labels of a chart.
; Syntax.........: _XLChart_TicksSet($oObject[, $iMajorTickmark = Default[, $iMinorTickmark = Default[, $vMajorUnit = Default[, $vMinorUnit = Default[, $iLabelPosition = Default[, $vLabelSpacing = Default[, $sLabelNumberFormat = Default]]]]]]])
; Parameters ....: $oObject            - Object of the axis for which the tick properties should be set
;                  $iMajorTickmark     - Optional: Sets the type of major tick marks. Can be one of the enumeration XlTickMark (default = Default)
;                  $iMinorTickmark     - Optional: Sets the type of minor tick marks. Can be one of the enumeration XlTickMark (default = Default)
;                  $vMajorUnit         - Optional: Sets the major units for the value axis. Can be numeric or "Auto" (default = Default)
;                  $vMinorUnit         - Optional: Sets the minor units for the value axis. Can be numeric or "Auto"  (default = Default)
;                  $iLabelPosition     - Optional: Sets the position of tick-mark labels. Can be one of the enumeration XlTickLabelPosition (default = Default)
;                  $vLabelSpacing      - Optional: Sets the spacing between tick labels. Can be a numeric value or "Auto" (default = Default)
;                                        Applies only to category and series axes
;                  $sLabelNumberFormat - Optional: Sets the number format of the tick labels (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is no object
;                  |2 - Parameter $vMajorUnit and $vMinorUnit are only valid for the value axis
;                  |3 - Invalid value for $vMajorUnit. Must be numeric or "Auto"
;                  |4 - Invalid value for $vMinorUnit. Must be numeric or "Auto"
;                  |5 - Invalid value for $iLabelPosition. Must be numeric
;                  |6 - Invalid value for $vLabelSpacing. Must be numeric or "Auto"
;                  |7 - Error setting $vMajorUnit. Please check @extended for detailed error code
;                  |8 - Error setting $vMinorUnit. Please check @extended for detailed error code
;                  |9 - Error setting $iLabelPosition. Please check @extended for detailed error code
;                  |10 - Error setting $iLabelSpacing. Please check @extended for detailed error code
;                  |11 - $iMajorTickmark is not an integer
;                  |12 - $iMinorTickmark is not an integer
;                  |13 - $sLabelNumberFormat format error
; Authors........: water
; Modified ......: Greencan
; Remarks .......: $sLabelNumberFormat: The format code is the same string as the Format Codes option in the Format Cells dialog box
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_TicksSet($oObject, $iMajorTickmark = Default, $iMinorTickmark = Default, $vMajorUnit = Default, $vMinorUnit = Default, $iLabelPosition = Default, $vLabelSpacing = Default, $sLabelNumberFormat = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $iMajorTickmark <> Default And Not IsInt($iMajorTickmark) Then Return SetError(11, 0, 0)
	If $iMinorTickmark <> Default And Not IsInt($iMinorTickmark) Then Return SetError(12, 0, 0)
	If $iMajorTickmark <> Default Then $oObject.MajorTickMark = $iMajorTickmark
	If $iMinorTickmark <> Default Then $oObject.MinorTickMark = $iMinorTickmark
	If $vMajorUnit <> Default Then
		; If $oObject.Type <> $xlValue Then Return SetError(2, 0, 0)  	;==> removed by GreenCan on 1/1/2013 because uncorrect
		If $vMajorUnit == "Auto" Then
			$oObject.MajorUnitIsAuto = True
		ElseIf IsNumber($vMajorUnit) Then
			$oObject.MajorUnit = $vMajorUnit
		Else
			Return SetError(3, 0, 0)
		EndIf
		If @error Then Return SetError(7, 0, 0)
	EndIf
	If $vMinorUnit <> Default Then
		; If $oObject.Type <> $xlValue Then Return SetError(2, 0, 0)  	;==> removed by GreenCan on 1/1/2013 because uncorrect
		If $vMinorUnit == "Auto" Then
			$oObject.MinorUnitIsAuto = True
		ElseIf IsNumber($vMinorUnit) Then
			$oObject.MinorUnit = $vMinorUnit
		Else
			Return SetError(4, 0, 0)
		EndIf
		If @error Then Return SetError(8, 0, 0)
	EndIf
	If $iLabelPosition <> Default Then
		If Not IsInt($iLabelPosition) Then Return SetError(5, 0, 0)
		$oObject.TickLabelPosition = $iLabelPosition
		If @error Then Return SetError(9, 0, 0)
	EndIf
	If $vLabelSpacing <> Default Then
		If $vLabelSpacing == "Auto" Then
			$oObject.TickLabelSpacingIsAuto = True
		ElseIf IsNumber($vLabelSpacing) Then
			$oObject.TicklabelSpacing = $vLabelSpacing
		Else
			Return SetError(6, 0, 0)
		EndIf
		If @error Then Return SetError(10, 0, 0)
	EndIf
	If $sLabelNumberFormat <> Default Then
		$oObject.TickLabels.NumberFormat = $sLabelNumberFormat
		If @error Then Return SetError(13, 0, 0)
	EndIf
	Return 1

EndFunc   ;==>_XLChart_TicksSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_TitleGet
; Description....: Return information about a title. This can be the chart or any axis title.
; Syntax.........: _XLChart_TitleGet($oObject)
; Parameters ....: $oObject - Title object (e.g. $oChart.ChartTitle or $oChart.Axes([$xlCategory|$xlValue|$xlSeriesAxis]).AxisTitle)
; Return values .: Success - One-dimensional one based array with the following information:
;                  |1 - AutoScaleFont: True if the text in the object changes font size when the object size changes
;                  |2 - Caption: String representing the chart or axis title text
;                  |3 - Format: ChartFormat object
;                  |4 - FormulaLocal: String representing the formula in A1-style notation (in the language of the user)
;                  |5 - FormulaR1C1Local: String representing the formula in R1C1-style notation (in the language of the user)
;                  |6 - Height: Height in points
;                  |7 - HorizontalAlignment: Variant representing the horizontal alignment. See the XlHAlign enumeration
;                  |8 - IncludeInLayout: True if a chart or axis title will occupy the chart layout space when a chart layout is being determined (default = True)
;                  |9 - Left: Value in points representing the distance from the left edge of the title to the left edge of the chart area
;                  |10 - Orientation: Variant representing the text orientation
;                  |11 - Position: Position of the title on the chart. XlChartElementPosition enumeration
;                  |12 - Shadow: Boolean that determines if the object has a shadow
;                  |13 - Top: Value in points representing the distance from the top edge of the title to the top of the chart area
;                  |14 - VerticalAlignment: Variant representing the vertical alignment
;                  |15 - Width: Width in points
;                  Failure - Returns "" and sets @error:
;                  |1 - $oObject is not an object
; Authors........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_TitleGet($oObject)

	If Not IsObj($oObject) Then Return SetError(1, 0, "")
	Local $asTitle[16] = [15]
	$asTitle[1] = $oObject.AutoScaleFont
	$asTitle[2] = $oObject.Caption
	$asTitle[3] = $oObject.Format
	$asTitle[4] = $oObject.FormulaLocal
	$asTitle[5] = $oObject.FormulaR1C1Local
	$asTitle[6] = $oObject.Height
	$asTitle[7] = $oObject.HorizontalAlignment
	$asTitle[8] = $oObject.IncludeInLayout
	$asTitle[9] = $oObject.Left
	$asTitle[10] = $oObject.Orientation
	$asTitle[11] = $oObject.Position
	$asTitle[12] = $oObject.Shadow
	$asTitle[13] = $oObject.Top
	$asTitle[14] = $oObject.VerticalAlignment
	$asTitle[15] = $oObject.Width
	Return $asTitle

EndFunc   ;==>_XLChart_TitleGet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_TitleSet
; Description....: Set properties of a title. This can be the chart or any axis title.
; Syntax.........: _XLChart_TitleSet($oObject[, $sCaption = Default[, $bShadow = Default[, $iLeft = Default[, $iTop = Default[, $iHorizontalAlignment = Default[, $iVerticalAlignment = Default[, $iOrientation = Default[, $bIncludeInLayout = Default]]]]]]]])
; Parameters ....: $oObject              - Title object (e.g. $oChart.ChartTitle or $oChart.Axes([$xlCategory|$xlValue|$xlSeriesAxis]).AxisTitle)
;                  $sCaption             - Optional: String representing the title text. "" removes an existing title caption (default = Default)
;                  $bShadow              - Optional: Boolean that determines if the object has a shadow (default = Default)
;                  $iLeft                - Optional: Value in points representing the distance from the left edge of the title to the left edge of the chart area (default = Default)
;                  $iTop                 - Optional: Value in points representing the distance from the top edge of the title to the top of the chart area (default = Default)
;                  $iHorizontalAlignment - Optional: Variant representing the horizontal alignment. A subset of the XlHAlign enumeration (default = Default). See the Remarks
;                  $iVerticalAlignment   - Optional: Variant representing the vertical alignment (default = Default)
;                  $iOrientation         - Optional: Variant representing the text orientation (default = Default)
;                  $bIncludeInLayout     - Optional: True if a chart or axis title will occupy the chart layout space when a chart layout is being determined  (default = Default)
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is no object
;                  |2 - $bShadow is not boolean
;                  |3 - $iLeft is not a number
;                  |4 - $iTop is not a number
;                  |5 - $iHorizontalAlignment is not a number
;                  |6 - $iVerticalAlignment is not a number
;                  |7 - $iOrientation is not a number
;                  |8 - $bIncludeInLayout is not boolean
; Authors........: water
; Modified ......:
; Remarks .......: The value of $iHorizontalAlignment can be one of the following out of the XlHAlign enumeration:
;                  $xlHAlignCenter, $xlHAlignDistributed, $xlHAlignJustify, $xlHAlignLeft, $xlHAlignRight
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_TitleSet($oObject, $sCaption = Default, $bShadow = Default, $iLeft = Default, $iTop = Default, $iHorizontalAlignment = Default, $iVerticalAlignment = Default, $iOrientation = Default, $bIncludeInLayout = Default)

	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If $bShadow <> Default And Not IsBool($bShadow) Then Return SetError(2, 0, 0)
	If $iLeft <> Default And Not IsNumber($iLeft) Then Return SetError(3, 0, 0)
	If $iTop <> Default And Not IsNumber($iTop) Then Return SetError(4, 0, 0)
	If $iHorizontalAlignment <> Default And Not IsNumber($iHorizontalAlignment) Then Return SetError(5, 0, 0)
	If $iVerticalAlignment <> Default And Not IsNumber($iVerticalAlignment) Then Return SetError(6, 0, 0)
	If $iOrientation <> Default And Not IsNumber($iOrientation) Then Return SetError(7, 0, 0)
	If $bIncludeInLayout <> Default And Not IsBool($bIncludeInLayout) Then Return SetError(8, 0, 0)
	If $sCaption <> Default Then $oObject.Caption = $sCaption
	If $bShadow <> Default Then $oObject.Shadow = $bShadow
	If $iLeft <> Default Then $oObject.Left = $iLeft
	If $iTop <> Default Then $oObject.Top = $iTop
	If $iHorizontalAlignment <> Default Then $oObject.HorizontalAlignment = $iHorizontalAlignment
	If $iVerticalAlignment <> Default Then $oObject.VerticalAlignment = $iVerticalAlignment
	If $iOrientation <> Default Then $oObject.Orientation = $iOrientation
	If $bIncludeInLayout <> Default Then $oObject.IncludeInLayout = $bIncludeInLayout
	Return 1

EndFunc   ;==>_XLChart_TitleSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_TrendlineSet
; Description....: Add a new trendline or set properties of an existing trendline of a data series.
; Syntax.........: _XLChart_TrendlineSet($oObject, $iNumber[, $iType = Default[, $iOption = Default[, $sName = Default]]])
; Parameters ....: $oObject - Data series object to add or change a trendline
;                  $iNumber - Number of the trendline to change or 0 to add a new trendline
;                  $iType   - Optional: Integer representing the trendline type. Can be any of the XlTrendlineType enumeration (default = Default)
;                  $iOption - Optional: Sets trendline option (default = Default). Parameter is only valid for:
;                  |$iType = $xlPolynomial. $iOption is interpreted to set property "Order"
;                  |$iType = $xlMovingAvg. $iOption is interpreted to set property "Period"
;                  $sName   - Optional: Name of the trendline. "" deletes the name (default = Default)
; Return values .: Success - Object identifier of the created/changed trend line
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oObject is no object
;                  |2 - $iNumber is not an integer or < 0
;                  |3 - $iType is not an integer
;                  |4 - Error creating trendline. See @extended for details
;                  |5 - Error accessing the specified trendline. See @extended for details
;                  |6 - Error setting the trendline option (order or period). See @extended for details
; Authors........: Greencan, water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_TrendlineSet($oObject, $iNumber, $iType = Default, $iOption = Default, $sName = Default)

	Local $oTrendline
	If Not IsObj($oObject) Then Return SetError(1, 0, 0)
	If Not IsInt($iNumber) Or $iNumber < 0 Then Return SetError(2, 0, 0)
	If $iType <> Default And Not IsInt($iType) Then Return SetError(3, 0, 0)
	If $iNumber = 0 Then ; Create new trendline
		$oTrendline = $oObject.Trendlines.Add
		If @error Then Return SetError(4, 0, 0)
	Else
		$oTrendline = $oObject.Trendlines.Item($iNumber)
		If @error Then Return SetError(5, 0, 0)
		If $iType <> Default Then $oTrendline.Type = $iType
	EndIf
	If $iOption <> Default Then
		If $oTrendline.Type = $xlPolynomial Then $oTrendline.Order = $iOption
		If $oTrendline.Type = $xlMovingAvg Then $oTrendline.Period = $iOption
		If @error Then Return SetError(6, @error, 0)
	EndIf
	If $sName <> Default Then $oTrendline.Name = $sName
	Return $oTrendline

EndFunc   ;==>_XLChart_TrendlineSet

; #FUNCTION# ====================================================================================================
; Name...........: _XLChart_VersionInfo
; Description ...: Returns an array of information about the ExcelChart UDF.
; Syntax.........: _XLChart_VersionInfo()
; Parameters ....: None
; Return values .: Success - One-dimensional one based array with the following information:
;                  |1 - Release Type (T=Test or V=Production)
;                  |2 - Major Version
;                  |3 - Minor Version
;                  |4 - Sub Version
;                  |5 - Release Date (YYYYMMDD)
;                  |6 - AutoIt version required
;                  |7 - List of authors separated by ","
;                  |8 - List of contributors separated by ","
; Author ........: water
; Modified.......:
; Remarks .......: Based on function _IE_VersionInfo written bei Dale Hohm
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================
Func _XLChart_VersionInfo()

	Local $avVersionInfo[9] = [8, "V", 0, 4, 0.0, "20150401", "3.3.12.0", "water, GreenCan", ""]
	Return $avVersionInfo

EndFunc   ;==>_XLChart_VersionInfo

; #INTERNAL_USE_ONLY#============================================================================================
; Name ..........: _XLChart_COMError
; Description ...: Called if an ObjEvent error occurs.
; Syntax.........: _XLChart_COMError()
; Parameters ....: None
; Return values .: Sets @error to 999 and @error to the COM error number (decimal)
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================
Func _XLChart_COMError()

	Local $bHexNumber = Hex($g__oError.number, 8)
	Local $avVersionInfo = _XLChart_VersionInfo()
	Local $sError = "COM Error Encountered in " & @ScriptName & @CRLF & _
			"ExcelChart UDF version = " & $avVersionInfo[2] & "." & $avVersionInfo[3] & "." & $avVersionInfo[4] & @CRLF & _
			"@AutoItVersion = " & @AutoItVersion & @CRLF & _
			"@AutoItX64 = " & @AutoItX64 & @CRLF & _
			"@Compiled = " & @Compiled & @CRLF & _
			"@OSArch = " & @OSArch & @CRLF & _
			"@OSVersion = " & @OSVersion & @CRLF & _
			"Scriptline = " & $g__oError.scriptline & @CRLF & _
			"NumberHex = " & $bHexNumber & @CRLF & _
			"Number = " & $g__oError.number & @CRLF & _
			"WinDescription = " & StringStripWS($g__oError.WinDescription, 2) & @CRLF & _
			"Description = " & StringStripWS($g__oError.description, 2) & @CRLF & _
			"Source = " & $g__oError.Source & @CRLF & _
			"HelpFile = " & $g__oError.HelpFile & @CRLF & _
			"HelpContext = " & $g__oError.HelpContext & @CRLF & _
			"LastDllError = " & $g__oError.LastDllError
	If $g__iDebug > 0 Then
		If $g__iDebug = 1 Then ConsoleWrite($sError & @CRLF & "========================================================" & @CRLF)
		If $g__iDebug = 2 Then MsgBox(64, "ExcelChart UDF - Debug Info", $sError)
		If $g__iDebug = 3 Then FileWrite($g__sDebugFile, @YEAR & "." & @MON & "." & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " " & @CRLF & _
				"-------------------" & @CRLF & $sError & @CRLF & "========================================================" & @CRLF)
	EndIf
	Return SetError(999, $g__oError.number, 0)

EndFunc   ;==>_XLChart_COMError

; #INTERNAL_USE_ONLY#============================================================================================
; Name...........: _XLChart_Example
; Description....: Creates a new Excel workbook and populates it with some example data and example charts.
; Syntax.........: _XLChart_Example([$bCreateChart = False[, $iChart1 = 0[, $iChart2 = 0[, $iChart3 = 0[, $iChart4 = 0[, $iChart5 = 0[, $iChart6 = 0]]]]]]])
; Parameters ....: $bCreateChart - Optional: True if the sample charts should be created (default = False)
;                  $iChart1      - Optional: Type of chart1. -1 means no chart at this position (default = 0 which means a chart of type $xlLineMarkers)
;                  $iChart2      - Optional: Type of chart2. -1 means no chart at this position (default = 0 which means a chart of type $xlColumnClustered)
;                  $iChart3      - Optional: Type of chart3. -1 means no chart at this position (default = 0 which means a chart of type $xl3DColumn)
;                  $iChart4      - Optional: Type of chart4. -1 means no chart at this position (default = 0 which means a chart of type $xlPyramidCol)
;                  $iChart5      - Optional: Type of chart5. -1 means no chart at this position (default = -1)
;                  $iChart6      - Optional: Type of chart6. -1 means no chart at this position (default = -1)
; Return values .: Success - One-dimensional zero based array with the following information:
;                  |0 - Excel object
;                  |1 - Object of the 1st chart
;                  |2 - Object of the 2nd chart
;                  |3 - Object of the 3rd chart
;                  |4 - Object of the 4th chart
;                  |5 - Object of the 5th chart
;                  |6 - Object of the 6th chart
;                  Failure - Returns "" and sets @error:
;                  |1 - Unable to create the Excel COM object (as returned by _Excel_BookNew). The error returned by _Excel_BookNew is returned in @extended
;                  |2 - The installed Excel version is not supported by this UDF. Version must be >= 12 (Excel 2007)
; Authors........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================
Func _XLChart_Example($bCreateChart = False, $iChart1 = 0, $iChart2 = 0, $iChart3 = 0, $iChart4 = 0, $iChart5 = -1, $iChart6 = -1)

	Local $aoCharts[7], $sChartName
	; Create a new workbook
	Local $oExcel = ObjCreate("Excel.Application")
	If @error <> 0 Or Not IsObj($oExcel) Then Return SetError(1, @error, "")
	If _XLChart_Version($oExcel) < 12 Then Return SetError(2, 0, "")
	$oExcel.Visible = 1
	$oExcel.WorkBooks.Add
	$oExcel.ActiveWorkbook.Sheets(1).Select()
	$aoCharts[0] = $oExcel
	; Create data
	$oExcel.Activesheet.Cells(1, 1).Value = "Date"
	$oExcel.Activesheet.Cells(2, 1).Value = "'01/2011"
	$oExcel.Activesheet.Cells(3, 1).Value = "'02/2011"
	$oExcel.Activesheet.Cells(4, 1).Value = "'03/2011"
	$oExcel.Activesheet.Cells(5, 1).Value = "'04/2011"
	$oExcel.Activesheet.Cells(6, 1).Value = "'05/2011"
	$oExcel.ActiveSheet.Columns(1).AutoFit
	$oExcel.Activesheet.Cells(1, 2).Value = "Sales Store 1"
	$oExcel.Activesheet.Cells(2, 2).Value = "10"
	$oExcel.Activesheet.Cells(3, 2).Value = "23"
	$oExcel.Activesheet.Cells(4, 2).Value = "15"
	$oExcel.Activesheet.Cells(5, 2).Value = "20"
	$oExcel.Activesheet.Cells(6, 2).Value = "34"
	$oExcel.ActiveSheet.Columns(2).AutoFit
	$oExcel.Activesheet.Cells(1, 3).Value = "Sales Store 2"
	$oExcel.Activesheet.Cells(2, 3).Value = "18"
	$oExcel.Activesheet.Cells(3, 3).Value = "23"
	$oExcel.Activesheet.Cells(4, 3).Value = "40"
	$oExcel.Activesheet.Cells(5, 3).Value = "32"
	$oExcel.Activesheet.Cells(6, 3).Value = "28"
	$oExcel.ActiveSheet.Columns(3).AutoFit
	$oExcel.ActiveSheet.Columns(7).ColumnWidth = 2
	; Set the names of the worksheets
	$oExcel.ActiveSheet.Name = "_XLChart_Example"
	; Create sample charts
	If $bCreateChart Then
		Local $asPosition[7] = [6, "A8:F26", "H8:M26", "A28:F46", "H28:M46", "A48:F66", "H48:M66"]
		Local $aiType[7] = [6, $iChart1, $iChart2, $iChart3, $iChart4, $iChart5, $iChart6]
		If $aiType[1] = 0 Then $aiType[1] = 1
		If $aiType[2] = 0 Then $aiType[2] = 2
		If $aiType[3] = 0 Then $aiType[3] = 3
		If $aiType[4] = 0 Then $aiType[4] = 4
		Local $XValueRange = "=_XLChart_Example!R2C1:R6C1"
		Local $asDataRange[3] = [2, "=_XLChart_Example!R2C2:R6C2", "=_XLChart_Example!R2C3:R6C3"]
		Local $asDataName[3] = [2, "=_XLChart_Example!B1", "=_XLChart_Example!C1"]
		For $iIndex = 1 To $aiType[0]
			$sChartName = "Example Chart " & $iIndex
			Switch $aiType[$iIndex]
				Case -1
				Case 1
					; LINES MARKERS: With legend
					$aoCharts[$iIndex] = _XLChart_ChartCreate($oExcel, 1, $xlLineMarkers, $asPosition[$iIndex], $sChartName, $XValueRange, $asDataRange, $asDataName, True, "Sales", "Date", "", "Quantity")
				Case 2
					; COLUMN CLUSTERED: With legend
					$aoCharts[$iIndex] = _XLChart_ChartCreate($oExcel, 1, $xlColumnClustered, $asPosition[$iIndex], $sChartName, $XValueRange, $asDataRange, $asDataName, True, "Sales", "Date", "", "Quantity")
				Case 3
					; 3D COLUMN: Without legend, with data table
					$aoCharts[$iIndex] = _XLChart_ChartCreate($oExcel, 1, $xl3DColumn, $asPosition[$iIndex], $sChartName, $XValueRange, $asDataRange, $asDataName, False, "Sales", "Date", "Location", "Quantity")
				Case 4
					; 3D PYRAMID COLUMN: With legend
					$aoCharts[$iIndex] = _XLChart_ChartCreate($oExcel, 1, $xlPyramidCol, $asPosition[$iIndex], $sChartName, $XValueRange, $asDataRange, $asDataName, True, "Sales", "Date", "Location", "Quantity")
				Case 5
					; 3D STACKED AREA: Without legend
					$aoCharts[$iIndex] = _XLChart_ChartCreate($oExcel, 1, $xl3DAreaStacked, $asPosition[$iIndex], $sChartName, $XValueRange, $asDataRange, $asDataName, True, "Sales", "Date")
				Case 6
					; 3D EXPLODED PIE: With legend. Create the chart on a separate chartsheet
					$aoCharts[$iIndex] = _XLChart_ChartCreate($oExcel, 1, $xl3DPieExploded, 0, $sChartName, $XValueRange, "=_XLChart_Example!R2C2:R6C2", "=_XLChart_Example!B1", True, "Sales")
				Case 7
					; 3D EXPLODED PIE: With legend
					$aoCharts[$iIndex] = _XLChart_ChartCreate($oExcel, 1, $xl3DPieExploded, $asPosition[$iIndex], $sChartName, $XValueRange, "=_XLChart_Example!R2C2:R6C2", "=_XLChart_Example!B1", True, "Sales")
				Case Else
					; See Case 1
					$aoCharts[$iIndex] = _XLChart_ChartCreate($oExcel, 1, $xlLineStacked, $asPosition[$iIndex], $sChartName, $XValueRange, $asDataRange, $asDataName, True, "Sales", "Date", "", "Quantity")
			EndSwitch
		Next
	EndIf
	Return $aoCharts

EndFunc   ;==>_XLChart_Example

; #INTERNAL_USE_ONLY#============================================================================================
; Name...........: _XLChart_Version
; Description....: Returns the installed Excel version as version number or text.
; Syntax.........: _XLChart_Version($oExcel[, $bText = False])
; Parameters ....: $oExcel - Excel object opened by a preceding call to _Excel_BookOpen() or _Excel_BookNew()
;                  $bText  - Optional: If set to True the Excel version is returned as text e.g. "Excel 2007" (default = False)
; Return values .: Success - Excel version number as integer or text
;                  Failure - Returns 0 and sets @error:
;                  |1 - $oExcel is not an object
; Authors........: GreenCan
; Modified ......: water
; Remarks .......: The ExcelChart UDF is written for Excel 2007 or later
;                  Example: For Excel 2010 you either get 14 ($bText = False) or "Excel 2010" ($bText = True)
;+
;                  This function is valid not only for the Excel Chart UDF but is a general Excel function.
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================
Func _XLChart_Version($oExcel, $bText = False)

	If Not IsObj($oExcel) Then Return SetError(1, 0, 0)
	If $bText = False Then Return Int($oExcel.Version)
	Switch $oExcel.Version
		Case 5
			Return "Excel 5"
		Case 7
			Return "Excel 95"
		Case 8
			Return "Excel 97"
		Case 9
			Return "Excel 2000"
		Case 10
			Return "Excel 2002"
		Case 11
			Return "Excel 2003"
		Case 12
			Return "Excel 2007"
		Case 14
			Return "Excel 2010"
		Case 15
			Return "Excel 2013"
		Case Else
			Return "Unknown version"
	EndSwitch
	Return "Unknown version"

EndFunc   ;==>_XLChart_Version

; #INTERNAL_USE_ONLY#============================================================================================
; Name...........: _XLChart_RGB
; Description....: Translate a RGB color definition to the format used by Excel (BGR).
; Syntax.........: _XLChart_RGB($iRGB)
; Parameters ....: $iRGB - Integer representing the RGB value
; Return values .: Success - Integer representing the BGR value
; Authors........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================
Func _XLChart_RGB($iRGB)

	$iRGB = Int($iRGB) ; make sure the input value is of type Int. Important for 3.3.8.0 because function Hex returns 16 characters for type Double
	Local $sTemp1, $sTemp2
	$sTemp1 = StringRight("000000" & Hex($iRGB), 6)
	$sTemp2 = StringRight($sTemp1, 2) & StringMid($sTemp1, 3, 2) & StringLeft($sTemp1, 2)
	Return Dec($sTemp2)

EndFunc   ;==>_XLChart_RGB