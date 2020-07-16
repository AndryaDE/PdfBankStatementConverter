#include <Array.au3>

; ==== CONFIGS ==== ;

Global Const $_OUTPUT_PER_MONTH = True
Global Const $_OUTPUT_FILE_PREFIX	= "PB_Kontoverlauf_"
Global Const $_OUTPUT_FILE_SUFFIX	= ".txt"

Global Const $_PDF_ORDNER	= @ScriptDir
Global Const $_PDF_NAMES	= "PB_Kontoauszug_KtoNr_*.pdf"

Global Const $_WAEHRUNG_SUFFIX = "€"
Global Const $_DATUM_SUFFIX = ""

; ==== REGULAR EXPESSIONS ==== ;

Global Const $_AUSZUG	= "(?m)(\d\d\.\d\d\.\/\d\d\.\d\d\..{65}.*(\-|\+)\s(\d+\.)?\d{1,3}\,\d\d(\R.*)*?)\R\R"
Global Const $_BETRAG	= "(?m)\A.*((\-|\+)\s(\d+\.)?\d{1,3}\,\d\d)$"
Global Const $_VERWENDUNG_TRIM_LEFT = "(?m)^(.{17})"
Global Const $_VERWENDUNG_TRIM_RIGHT = "(?m)(\s+(\-|\+)\s(\d+\.)?\d{1,3}\,\d\d)$"

; ==== SCRIPT CODE ==== ;

#include <Array.au3>
#include <File.au3>
#include <StringConstants.au3>

Func Main()

	Local $pdfs = _FileListToArray($_PDF_ORDNER, $_PDF_NAMES) ;0 = item count
	Dim $array[0]

	Local $debug_output = ""

	For $i = $pdfs[0] To 1 Step -1

		Local $data = ReadPDF($pdfs[$i], "-simple")
		Local $auszug = StringRegExp($data, $_AUSZUG, $STR_REGEXPARRAYGLOBALMATCH)

		WriteFile($data, "debug_full")

		For $j = UBound($auszug)-1 To 0 Step -1

			If StringLen($auszug[$j]) <= 93 Then
				ContinueLoop
			EndIf

			$debug_output &= $auszug[$j] & @LF

			Local $date
			Local $verwendung
			Local $betrag
			AnalyseAuszug($auszug[$j], $date, $verwendung, $betrag)


			ReDim $array[UBound($array)+1]
			$array[UBound($array)-1] = $date & @TAB & '"' & $verwendung & '"' & @TAB & $betrag	;Export Format

		Next

	Next

	If StringLen($debug_output) >= 1 Then
		WriteFile($debug_output, "debug")
	EndIf

	If $_OUTPUT_PER_MONTH Then
		OutputPerMonth($array)
	Else
		OutputAsFile($array)
	EndIf

EndFunc

Func AnalyseAuszug($auszug, ByRef $date, ByRef $verwendung, ByRef $betrag)

	$auszug = " " & $auszug

	$date = StringMid($auszug, 9, 6)
	$date &= $_DATUM_SUFFIX

	$betrag = StringRegExp($auszug, $_BETRAG, $STR_REGEXPARRAYMATCH)
	$betrag = StringReplace($betrag[0], " ", "")
	$betrag = StringReplace($betrag, "+", "")
	$betrag &= $_WAEHRUNG_SUFFIX

	$verwendung = StringRegExpReplace($auszug, $_VERWENDUNG_TRIM_LEFT, "")
	$verwendung = StringRegExpReplace($verwendung, $_VERWENDUNG_TRIM_RIGHT, "", 1)
	$verwendung = StringReplace($verwendung, '"', "''")

EndFunc

#Region Output

Func OutputPerMonth($array)

	For $i = 1 To 12

		Dim $month[0]

		For $j = UBound($array)-1 To 0 Step -1

			Local $x = Number(StringMid($array[$j], 4, 2))
			If $x <> $i Then
				ContinueLoop
			EndIf

			ReDim $month[UBound($month)+1]
			$month[UBound($month)-1] = $array[$j]

		Next

		If UBound($month) >= 1 Then
			OutputAsFile($month, $i)
		EndIf

	Next

EndFunc

Func OutputAsFile($array, $file = @YEAR&@MON&@MDAY&@HOUR&@MIN&@SEC)

	_Sort($array)
	Local $output = _ArrayToString($array, @LF)
	WriteFile($output, $file)

EndFunc

Func WriteFile($output, $file = @YEAR&@MON&@MDAY&@HOUR&@MIN&@SEC)

	$file = $_OUTPUT_FILE_PREFIX & $file & $_OUTPUT_FILE_SUFFIX
	FileWrite($file, $output)

EndFunc

#EndRegion

#Region Sort

;~ Local $aArray = ["INFO [13.06.2017 11:48:01] [Thread-13] [ConGenImpUsb -> waitForConnection]", _
;~         "INFO [07.06.2017 08:55:44] [main] MDU5 - Ver 5.1x", _
;~         "INFO [07.06.2017 12:55:11] [main] Dummy String1", _
;~         "INFO [07.06.2016 09:55:11] [main] Dummy String2", _
;~         "INFO [07.06.2017 09:55:12] [main] Dummy String3", _
;~         "INFO [07.06.2017 09:55:11] [main] Dummy String4"]


;~ _ArrayDisplay($aArray, "No sorted")
;~ _Sort($aArray)
;~ _ArrayDisplay($aArray, "Sorted")

Func _Sort(ByRef $aArray)
    For $i = UBound($aArray) - 1 To 1 Step -1
        For $j = 1 To $i
            If _GetNumber($aArray[$j - 1]) > _GetNumber($aArray[$j]) Then
                $temp = $aArray[$j - 1]
                $aArray[$j - 1] = $aArray[$j]
                $aArray[$j] = $temp
            EndIf
        Next
    Next
    Return $aArray
EndFunc   ;==>_Sort


Func _GetNumber($String)
    Return Number(StringRegExpReplace($string, '(\d{2})\.(\d{2})\.(\d{4})', "$3$2$1", 1))
EndFunc   ;==>_GetNumber

#EndRegion

#Region PDF

Func ReadPDF($pdf, $options = "-layout")

	Local $txt = StringTrimRight($pdf, 4) & "_temp.txt"

	_XPDF_ToText($pdf, $txt, $options)

	Local $handle = FileOpen($txt)
	Local $content = FileRead($handle)

	FileClose($handle)
	FileDelete($txt)

	Return $content

EndFunc


Func _XPDF_ToText($sPDFFile, $sTXTFile, $sOptions = "-layout")
    Local $sXPDFToText = @ScriptDir & "\pdftotext.exe"

    If NOT FileExists($sPDFFile) Then Return SetError(1, 0, 0)
    If NOT FileExists($sXPDFToText) Then Return SetError(2, 0, 0)

    Local $iReturn = ShellExecuteWait ( $sXPDFToText , $sOptions & ' "' & $sPDFFile & '" "' & $sTXTFile & '"', @ScriptDir, "", @SW_HIDE)
    If $iReturn = 0 Then Return 1

    Return 0

EndFunc ; ---> _XPDF_ToText

#EndRegion

Main()