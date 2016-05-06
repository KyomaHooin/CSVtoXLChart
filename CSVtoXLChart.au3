;
; Program: CSV to XL Chart parser
; Author: Richard Bruna (c) 2016
;
; Description: Parses data from CSV file and create MS Excel line graph
;

;TUNE

#AutoIt3Wrapper_Icon=toolbox.ico
#NoTrayIcon

;INCLUDE

#include <GUIListBox.au3>
#include <GUIConstantsEx.au3>
#include <File.au3>
#include <Excel.au3>
#include <C:\Users\brunari\Desktop\CSVtoXLChart\ExcelChart 0.4.0.0\ExcelChart.au3>
#include <C:\Users\brunari\Desktop\CSVtoXLChart\ExcelChart 0.4.0.0\ExcelChartConstants.au3>
#include <WindowsConstants.au3>

;VAR

global $csv_file = '' ;CSV soubor pro import
global $csv_array = '' ; CSV array
global $vsechny_mistnosti = '' ;Pole vsech mistnosti
global $vybrane_mistnosti = ''; vybrane pole mistnosti
global $column[6] = ['B2','C2','D2','E2','F2','G2'] ;Excel data columns
global $graph_pos[6] = ['D3:M25','E3:N25', 'F3:O25','G3:P25','H3:Q25','I3:R25']; Excel Graph position

; GUI

$gui = GUICreate("CSVtoXLChart 1.1", 337, 168, -1, -1)
$csv_button = GUICtrlCreateButton("CSV", 248, 13, 75, 22)
$graf_button = GUICtrlCreateButton("GRAF", 248, 70, 75, 22)
$text_mistnosti = GUICtrlCreateLabel("MISTNOST", 16, 46, 60, 17)
$text_error = GUICtrlCreateLabel('', 188, 141, 133, 17)
$input = GUICtrlCreateInput('', 16, 13, 209, 21) ;
$list_mistnosti = GUICtrlCreateList('', 16, 70, 153, 84, $LBS_SORT + $LBS_EXTENDEDSEL + $WS_VSCROLL)

;CONTROL

;Allready running?
If UBound(ProcessList("CSVtoXLChart.exe"), $UBOUND_ROWS) > 2 Then
	MsgBox(48, "", "Program už byl spuštěn!")
	Exit
EndIf

;MAIN

;nastaveni GUI
GUISetState(@SW_SHOW)
GUICtrlSetColor($text_error, 0xFF0000)

;main loop
While 1
	;catch GUI Event
	$event = GUIGetMsg()
	;buttons
	If $event = $csv_button Then
		;clear outputs
		GUICtrlSetData($input, '')
		GUICtrlSetData($list_mistnosti, '')
		;get file
		$csv_file = FileOpenDialog("Vyber CSV soubor", @HomeDrive, "CSV soubor (*.csv)", $FD_FILEMUSTEXIST,'',$gui)
		If @error Then
			GUICtrlSetData($text_error, "Nebyl vybrán žádný soubor!")
		Else
			;clear errror text
			GUICtrlSetData($text_error, '')
			;display file
			GUICtrlSetData($input, $csv_file)
			;parse CSV to array
			_FileReadToArray($csv_file, $csv_array, $FRTA_NOCOUNT + $FRTA_INTARRAYS, ';')
			;parse location
			parse_location()
		EndIf
	EndIf
	If $event = $graf_button Then
		;entry control
		If Not $csv_file Then
			GUICtrlSetData($text_error, "Nebyl vybrán žádný soubor!")
		ElseIf ubound(_GUICtrlListBox_GetSelItems($list_mistnosti)) < 2  Then
			GUICtrlSetData($text_error, "Nebyla vybrána místnost!")
		Elseif ubound(_GUICtrlListBox_GetSelItems($list_mistnosti)) > 4 then;max 4 rooms
			GUICtrlSetData($text_error, "Příliš mnoho místností!")
		else
			;clear errror text
			GUICtrlSetData($text_error, '')
			;get select
			$vybrane_mistnosti = _GUICtrlListBox_GetSelItemsText($list_mistnosti)
			;define global data  array by slect size
			global $vsechna_data[UBound($vybrane_mistnosti)]
			;get dte array
			$vsechna_data[0] = parse_date()
			;get the rest of the data arrays
			for $i=1 to ubound($vybrane_mistnosti) - 1
				$vsechna_data[$i] = parse_data($vybrane_mistnosti[$i])
			next
			;open excel
			create_excel()
			;fill data
			create_excel_data()
			;create graph
			create_excel_graph()
		EndIf
	EndIf
	;exit
	If $event = $GUI_EVENT_CLOSE Then Exit
WEnd

;FUNC

;parse locations from CSV file
Func parse_location()
	;fourth line of file split by semicolon
	$f = FileOpen($csv_file)
	;create location array
	Global $vsechny_mistnosti = StringSplit(FileReadLine($csv_file, 4), ';', $STR_NOCOUNT)
	;from third value
	For $i = 2 To UBound($vsechny_mistnosti) - 1
		;find all location by temperature capital
		If StringLeft($vsechny_mistnosti[$i], 1) == 'T' Then
			;remove prefix and populate list
			GUICtrlSetData($list_mistnosti, StringTrimLeft($vsechny_mistnosti[$i], 2))
		EndIf
	Next
	FileClose($f)
EndFunc

;parse data by location from CSV file
Func parse_date()
	local $d_data = ''
	;_ArrayDisplay($csv_array)
	For $i = 6 To UBound($csv_array, $UBOUND_ROWS) - 1
		$d_data &= StringRegExpReplace(($csv_array[$i])[0], '([0-9]+.)([0-9]+.).*','$1$2') & ';'
	Next
	$d_data = StringSplit($d_data, ';', $STR_NOCOUNT)
	_ArrayPush($d_data, 'datum', 1)
	Return $d_data
EndFunc

;parse data by location from CSV file
Func parse_data($selection)
	local $teplota = '', $vlhkost = '';
	;get back real CSV string value from selected room
	$teplota = 'T ' & $selection
	if StringIsUpper(StringMid($selection,1,1))  then
		$vlhkost = 'H ' & StringLeft($selection, 1) & '2' & StringTrimLeft($selection, 1) ;StringRegExpReplace($selection,"([A-Z]).*","\1[2]")
	else
		$vlhkost = 'H ' & $selection
	EndIf
	;get column indexes, start at second value
	$t_col_index = _ArraySearch($vsechny_mistnosti, $teplota, 2)
	$h_col_index = _ArraySearch($vsechny_mistnosti, $vlhkost, 2)
	;get data arrays
	local $t_data = '', $h_data = ''
	;populate arrays/strings
	For $i = 6 To UBound($csv_array, $UBOUND_ROWS) - 1
		;test if index exist
		if $t_col_index > -1 then $t_data &= ($csv_array[$i])[$t_col_index] & ';'
		if $h_col_index > -1 then $h_data &= ($csv_array[$i])[$h_col_index] & ';'
	Next
	$t_data = StringSplit($t_data, ';', $STR_NOCOUNT)
	$h_data = StringSplit($h_data, ';', $STR_NOCOUNT)
	;insert header and remove trailer
	_ArrayPush($t_data, $teplota, 1)
	_ArrayPush($h_data, $vlhkost, 1)
	local $data[2] = [$t_data,$h_data]
	return $data
EndFunc   ;==>parse_data

func create_excel()
	;create excel object and open it
	Global $excel = _Excel_Open()
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, 'Chyba', "Nepodařilo se spustit aplikaci Excel." & @CRLF & @CRLF & 'Chyba [' & @error & ']')
	;create workbook object with 0 sheets
	Global $workbook = _Excel_BookNew($excel, 1)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, 'Chyba', "Nepodařilo se vytvořit WorkBook." & @CRLF & @CRLF & 'Chyba [' & @error & ']')
	;name default sheet
	$excel.ActiveSheet.Name = 'mistnost'
EndFunc

func create_excel_data()
	;data col counter
	global $col = 0
	_Excel_RangeWrite($workbook, $workbook.Activesheet, $vsechna_data[0], 'A2')
	for $i=1 to ubound($vybrane_mistnosti) - 1
	;	;temperature
		_Excel_RangeWrite($workbook, $workbook.Activesheet, ($vsechna_data[$i])[0], $column[$col])
		$col += 1
	;	;humidity
		if UBound(($vsechna_data[$i])[1]) > 1 Then
			_Excel_RangeWrite($workbook, $workbook.Activesheet, ($vsechna_data[$i])[1], $column[$col])
			$col += 1
		EndIf
	next
EndFunc

Func create_excel_graph()
	;define x line chart
	global $graph_location = $graph_pos[$col - 1] ;2977 = ubound => 2..2977 3..2978
	global $graph_x_range = "=mistnost!R3C1:R" & UBound($vsechna_data[0]) + 1 & "C1" ; R3C1:R2978C1 => A3:A2978
	;get all data location
	global $graph_y_range[$col + 1],  $graph_y_name[$col + 1]
	$graph_y_range[0] = $col
	$graph_y_name[0] = $col
	for $i=1 to $col
		$graph_y_range[$i] = "mistnost!R3C" & $i + 1 & ":R" & UBound($vsechna_data[0]) + 1 & "C" & $i + 1
		$graph_y_name[$i] = "=mistnost!" & $column[$i - 1]
	next
	;create graph
	$graf = _XLChart_ChartCreate($excel, 1, $xlLine, $graph_location, '', $graph_x_range, $graph_y_range, $graph_y_name, True)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, 'Chyba', "Nepodařilo se vytvořit graf." & @CRLF & @CRLF & 'Chyba [' & @error & ']')
	;patch the Y scale
	_XLChart_AxisSet($graf.Axes($xlValue), 0, 100)
	If @error Then MsgBox($MB_SYSTEMMODAL, 'Chyba',"Nepodařilo se nastavit osu." & @CRLF & @CRLF & 'Chyba [' & @error & ']' )
EndFunc

