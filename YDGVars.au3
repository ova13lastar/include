#include-once

; #INDEX# =======================================================================================================================
; Title .........: YDGVars
; AutoIt Version : 3.3.14.5
; Language ......: French
; Description ...: Script .au3 développé pour améliorer la gestion des dictionnaires
; Author(s) .....: yann.daniel@assurance-maladie.fr
; Related .......: Inspired by https://github.com/dmwyatt/AutoItDict2
; ===============================================================================================================================

; #ENVIRONMENT# =================================================================================================================
; Settings
#AutoIt3Wrapper_Run_AU3Check=Y
#AutoIt3Wrapper_AU3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
Global $__g_oGVars
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name...........: _YDGVars_Init
; Description ...: Initialise le dictionnaire $__g_oGVars
; Syntax.........: _YDGVars_Init()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 20/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDGVars_Init($_sFuncName = "")
	Local $sFuncName = "_YDGVars_Init"
	Local $sLogDateTime = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
	Local $sMsg = "Initialisation via " & $_sFuncName & "()"
	ConsoleWrite($sLogDateTime & " : [" & $sFuncName & "] " & $sMsg & @CRLF)
	$__g_oGVars = ObjCreate("Scripting.Dictionary")
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDGVars_Set
; Description ...: Ajoute ou Modifie une cle-valeur au dictionnaire
; Syntax.........: _YDGVars_Set($key, $value)
; Parameters ....: $key      	- Cle
;				   $value      	- Valeur
; Return values .: 
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 20/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDGVars_Set($key, $value)
	Local $sFuncName = "_YDGVars_Set"
	If Not IsObj($__g_oGVars) Then _YDGVars_Init($sFuncName)
	If $__g_oGVars.Exists($key) Then
		$__g_oGVars.Item($key) = $value
	Else
		$__g_oGVars.Add($key, $value)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDGVars_Get
; Description ...: Renvoi la valeur d une cle du dictionnaire
; Syntax.........: _YDGVars_Get($key)
; Parameters ....: $key			- Cle
; Return values .: Success      - Valeur
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 20/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDGVars_Get($key)
	Local $sFuncName = "_YDGVars_Get"
	If Not IsObj($__g_oGVars) Then _YDGVars_Init($sFuncName)
	If $__g_oGVars.Exists($key) Then
		Return $__g_oGVars.Item($key)
	Else
		Return ""
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDGVars_Exists
; Description ...: verifie si la cle existe dans le dictionnaire
; Syntax.........: _YDGVars_Exists($key)
; Parameters ....: $key			- Cle
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 21/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDGVars_Exists($key)
	Local $sFuncName = "_YDGVars_Exists"
	If Not IsObj($__g_oGVars) Then _YDGVars_Init($sFuncName)
	If $__g_oGVars.Exists($key) Then
		Return True
	Else
		Return False
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDGVars_GetAllKeys
; Description ...: Renvoi la valeur d une cle du dictionnaire
; Syntax.........: _YDGVars_GetAllKeys()
; Parameters ....:
; Return values .: Array	- Tableau de cles
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 20/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDGVars_GetAllKeys()
	Local $sFuncName = "_YDGVars_GetAllKeys"
	If Not IsObj($__g_oGVars) Then _YDGVars_Init($sFuncName)
	Return $__g_oGVars.Keys
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDGVars_GetAllValues
; Description ...: Renvoi la valeur d une cle du dictionnaire
; Syntax.........: _YDGVars_GetAllValues()
; Parameters ....:
; Return values .: Array	- Tableau de valeurs
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 20/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDGVars_GetAllValues()
	Local $sFuncName = "_YDGVars_GetAllValues"
	If Not IsObj($__g_oGVars) Then _YDGVars_Init($sFuncName)
	Return $__g_oGVars.Items
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDGVars_Del
; Description ...: Supprime une cle-valeur du dictionnaire
; Syntax.........: _YDGVars_Del($key)
; Parameters ....: $key			- Cle
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 20/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDGVars_Del($key)
	Local $sFuncName = "_YDGVars_Del"
	If Not IsObj($__g_oGVars) Then _YDGVars_Init($sFuncName)
	If $__g_oGVars.Exists($key) Then
		$__g_oGVars.Remove($key)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDGVars_DelAll
; Description ...: Supprime toutes les cles-valeurs du dictionnaire
; Syntax.........: _YDGVars_DelAll()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 20/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDGVars_DelAll()
	Local $sFuncName = "_YDGVars_DelAll"
	If Not IsObj($__g_oGVars) Then _YDGVars_Init($sFuncName)
	Local $aKeys = $__g_oGVars.keys()
	For $key In $aKeys
		_YDGVars_Del($key)
	Next
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDGVars_Len
; Description ...: Renvoi le nombre d elements du dictionnaire
; Syntax.........: _YDGVars_Len()
; Parameters ....:
; Return values .: Integer		- Nombre elements
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 20/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDGVars_Len()
	Local $sFuncName = "_YDGVars_Len"
	If Not IsObj($__g_oGVars) Then _YDGVars_Init($sFuncName)
	Return $__g_oGVars.Count
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDGVars_GetArray
; Description ...: Renvoi un tableau de cles-valeurs
; Syntax.........: _YDGVars_GetArray()
; Parameters ....: 
; Return values .: Success          - Tableau de cles-valeurs
;                  Failure          - Tableau vide
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 20/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDGVars_GetArray()
	Local $sFuncName = "_YDGVars_GetArray"
	If Not IsObj($__g_oGVars) Then _YDGVars_Init($sFuncName)
	Local $aFalse[0]
	If _YDGVars_Len() > 0 Then
		Return $__g_oGVars
	Else
		Return $aFalse
	EndIf
EndFunc







