#include-once

; #INDEX# =======================================================================================================================
; Title .........: YDXpdf
; AutoIt Version : 3.3.14.5
; Language ......: French
; Description ...: Script .au3 développé pour gérer les fichiers pdf
; Author(s) .....: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================

; #ENVIRONMENT# =================================================================================================================
; Settings
#AutoIt3Wrapper_Run_AU3Check=Y
#AutoIt3Wrapper_AU3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
; Includes

#include <YDLogger.au3>
;~ #include <FileConstants.au3>
;~ #include <File.au3>
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name...........: _XFDF_Info
; Description....: Retrives informations from a PDF file
; Syntax.........: _XFDF_Info ( "File" [, "Info"] )
; Parameters.....: File    - PDF File.
;                  Info    - The information to retrieve
; Return values..: Success - If the Info parameter is not empty, returns the desired information for the specified Info parameter
;                          - If the Info parameter is empty, returns an array with all available informations
;                  Failure - 0, and sets @error to :
;                   1 - PDF File not found
;                   2 - Unable to find the external programm
; Remarks........: The array returned is two-dimensional and is made up as follows:
;                   $array[1][0] = Label of the first information (title, author, pages...)
;                   $array[1][1] = value of the first information
;                   ...
; ===============================================================================================================================
Func _XFDF_Info($sPDFFile, $sInfo = "")
    Local $sXPDFInfo = @ScriptDir & "\pdfinfo.exe"

    If NOT FileExists($sPDFFile) Then Return SetError(1, 0, 0)
    If NOT FileExists($sXPDFInfo) Then Return SetError(2, 0, 0)
    $sXPDFInfo = FileGetShortName($sXPDFInfo)

    Local $iPid = Run(@ComSpec & ' /c ' &  $sXPDFInfo & ' "' & $sPDFFile & '"', @ScriptDir, @SW_HIDE, 2)

    Local $sResult
    While 1
        $sResult &= StdoutRead($iPid)
        If @error Then ExitLoop
    WEnd

    Local $aInfos = StringRegExp($sResult, "(?m)^(.+?): +(.*)$", 3)
    If @error Or Mod( UBound($aInfos, 1), 2) = 1 Then Return SetError(3, 0, 0)

    Local $aResult [ UBound($aInfos, 1) / 2][2]

    For $i = 0 To UBound($aInfos) - 1 Step 2
        If $sInfo <> "" AND $aInfos[$i] = $sInfo Then Return $aInfos[$i + 1]
        $aResult[$i / 2][0] = $aInfos[$i]
        $aResult[$i / 2][1] = $aInfos[$i + 1]
    Next

    If $sInfo <> "" Then Return ""

    Return $aResult
EndFunc ; ---> _XFDF_Info


; #FUNCTION# ====================================================================================================================
; Name...........: _XPDF_Search
; Description....: Retrives informations from a PDF file
; Syntax.........: _XFDF_Info ( "File" [, "String" [, Case = 0 [, Flag = 0 [, FirstPage = 1 [, LastPage = 0]]]]] )
; Parameters.....: File    - PDF File.
;                  String    - String to search for
;                  Case      - If set to 1, search is case sensitive (default is 0)
;                  Flag      - A number to indicate how the function behaves. See below for details. The default is 0.
;                  FirstPage  - First page to convert (default is 1)
;                  LastPage   - Last page to convert (default is 0 = last page of the document)
; Return values..: Success -
;                   Flag = 0 - Returns 1 if the search string was found, or 0 if not
;                   Flag = 1 - Returns the number of occcurrences found in the whole PDF File
;                   Flag = 2 - Returns an array containing the number of occurrences found for each page
;                              (only pages containing the search string are returned)
;                              $array[0][0] - Number of matching pages
;                              $array[0][1] - Number of occcurrences found in the whole PDF File
;                              $array[n][0] - Page number
;                              $array[n][1] - Number of occcurrences found for the page
;                  Failure - 0, and sets @error to :
;                   1 - PDF File not found
;                   2 - Unable to find the external programm
; ===============================================================================================================================
Func _XPDF_Search($sPDFFile, $sSearch, $iCase = 0, $iFlag = 0, $iStart = 1, $iEnd = 0)
    Local $sXPDFToText = @ScriptDir & "\pdftotext.exe"
    Local $sOptions = " -layout -f " & $iStart
    Local $iCount = 0, $aResult[1][2] = [[0, 0]], $aSearch, $sContent, $iPageOccCount
   
    If NOT FileExists($sPDFFile) Then Return SetError(1, 0, 0)
    If NOT FileExists($sXPDFToText) Then Return SetError(2, 0, 0)
   
    If $iEnd > 0 Then $sOptions &= " -l " & $iEnd
   
    Local $iPid = Run($sXPDFToText & $sOptions & ' "' & $sPDFFile & '" -', @ScriptDir, @SW_HIDE, 2)
    While 1
        $sContent &= StdoutRead($iPid)
        If @error Then ExitLoop
    WEnd
   
   
    Local $aPages = StringSplit($sContent, chr(12) )
   
    For $i = 1 To $aPages[0]
        $iPageOccCount = 0
        While StringInStr($aPages[$i], $sSearch, $iCase, $iPageOccCount + 1)
            If $iFlag <> 1 AND $iFlag <> 2 Then
                $aResult[0][1] = 1
                ExitLoop
            EndIf
            $iPageOccCount += 1
        WEnd

        If $iPageOccCount Then
            Redim $aResult[ UBound($aResult, 1) + 1][2]
            $aResult[0][1] += $iPageOccCount
            $aResult[0][0] = UBound($aResult) - 1
            $aResult[ UBound($aResult, 1) - 1 ][0] = $i + $iStart - 1
            $aResult[ UBound($aResult, 1) - 1 ][1] = $iPageOccCount
        EndIf
    Next
   
    If $iFlag = 2 Then Return $aResult
    Return $aResult[0][1]
   
EndFunc ; ---> _XPDF_Search



; #FUNCTION# ====================================================================================================================
; Name...........: _XPDF_ToText
; Description....: Converts a PDF file to plain  text.
; Syntax.........: _XPDF_ToText ( "PDFFile" , "TxtFile" [ , FirstPage [, LastPage [, Layout ]]] )
; Parameters.....: PDFFile    - PDF Input File.
;                  TxtFile    - Plain text file to convert to
;                  FirstPage  - First page to convert (default is 1)
;                  LastPage   - Last page to convert (default is last page of the document)
;                  Layout     - If true, maintains (as  best as possible) the original physical layout of the text
;                               If false, the behavior is to 'undo'  physical  layout  (columns, hyphenation, etc.)
;                                 and output the text in reading order.
;                               Default is True
; Return values..: Success - 1
;                  Failure - 0, and sets @error to :
;                   1 - PDF File not found
;                   2 - Unable to find the external program
; ===============================================================================================================================
Func _XPDF_ToText($sPDFFile, $sTXTFile, $iFirstPage = 1, $iLastPage = 0, $bLayout = True)
    Local $sXPDFToText = @ScriptDir & "\pdftotext.exe"
    Local $sOptions
   
    If NOT FileExists($sPDFFile) Then Return SetError(1, 0, 0)
    If NOT FileExists($sXPDFToText) Then Return SetError(2, 0, 0)
   
    If $iFirstPage <> 1 Then $sOptions &= " -f " & $iFirstPage
    If $iLastPage <> 0 Then $sOptions &= " -l " & $iLastPage
    If $bLayout = True Then $sOptions &= " -layout"
   
    Local $iReturn = ShellExecuteWait ( $sXPDFToText , $sOptions & ' "' & $sPDFFile & '" "' & $sTXTFile & '"', @ScriptDir, "", @SW_HIDE)
    If $iReturn = 0 Then Return 1
   
    Return 0
   
EndFunc ; ---> _XPDF_ToText