#include-once

; #INDEX# =======================================================================================================================
; Title .........: YDLogger
; AutoIt Version : 3.3.14.5
; Language ......: French
; Description ...: Script .au3 développé pour gérer les log dans les programmes AutoIT (console + fichier de log)
; Author(s) .....: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================

; #ENVIRONMENT# =================================================================================================================
; Settings
#AutoIt3Wrapper_Run_AU3Check=Y
#AutoIt3Wrapper_AU3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
; Includes
#include <YDGVars.au3>
#include <FileConstants.au3>
#include <File.au3>
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
Global Const $YDLOGGER_LOGTYPE_CONSOLE = "console"
Global Const $YDLOGGER_LOGTYPE_FILE    = "file"
Global Const $YDLOGGER_LOGTYPE_BOTH    = "both"
Global Const $YDLOGGER_LOGEXT          = ".log"
Global Const $YDLOGGER_LOGLEVEL        = 1
Global Const $YDLOGGER_LOGCASE_ERROR   = "error"
Global Const $YDLOGGER_SEP_NUMBER      = 50
Global Const $YDLOGGER_SEP_CAR          = "-"
Global Const $YDLOGGER_FUNCNAME        = "MAIN"

Global $__g_sLastFuncName   = $YDLOGGER_FUNCNAME
Global $__g_sLogType        = $YDLOGGER_LOGTYPE_BOTH
Global $__g_bLogInit        = False
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name...........: _YDLogger_Init
; Description ...: Initialise le fichier de log
; Syntax.........: _YDLogger_Init([$_sLogPath], [$_sLogType])
; Parameters ....: $_oAppVars   - Objet dictionnaire contenant les variables globales issues de l'application appelante
;                  $_sLogPath   - Path du fichier log
;                  $_sLogType   - console | file | both
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......: Cette fonction doit etre appellee en second
; Related .......: 
; ===============================================================================================================================
Func _YDLogger_Init($_sLogPath = "", $_sLogType = $YDLOGGER_LOGTYPE_BOTH)
    ; On sort si pas de variables globales
    If _YDGVars_Len = 0 Then 
        MsgBox($MB_SYSTEMMODAL + $MB_ICONERROR + $MB_OK, "ERREUR", "Variables globales vides !! L'application ne peut pas s'executer.")
        Exit
    EndIf
    ; On détermine le niveau de log
    If FileExists(_YDGVars_Get("sAppConfFile")) Then
        Local $iLoglevel = IniRead(_YDGVars_Get("sAppConfFile"), "general", "loglevel", $YDLOGGER_LOGLEVEL)
        If @error Then
            _YDLogger_Error("Lecture impossible du fichier : " & _YDGVars_Get("sAppConfFile"))
            Return False
        EndIf
        ; On affecte la valeur trouvee a la variable globale
        _YDGVars_Set("sLogLevel", $iLoglevel)
    Else
        _YDGVars_Set("sLogLevel", $YDLOGGER_LOGLEVEL)
    EndIf    
    ; On affecte le type choisi a la variable globale
    Switch (StringLower($_sLogType))
        Case "console"
            $__g_sLogType = $YDLOGGER_LOGTYPE_CONSOLE
        Case "file"
            $__g_sLogType = $YDLOGGER_LOGTYPE_FILE
        Case Else
            $__g_sLogType = $YDLOGGER_LOGTYPE_BOTH
    EndSwitch
    ; On affecte le chemin choisi a la variable globale
    If $_sLogPath = "" Then
        Local $sPath = _YDGVars_Get("sAppDirLogsPath")
        $_sLogPath = $sPath & "\" & @YEAR & @MON & @MDAY & "_" & Stringreplace(StringReplace(@ScriptName, ".au3", ""), ".exe", "") & $YDLOGGER_LOGEXT
        ; On cree le dossier "logs" s'il nexiste pas
        If DirGetSize($sPath) = -1 And DirCreate($sPath) = 0 Then
            _YDLogger_Error("Creation dossier impossible : " & $sPath)
        EndIf
    EndIf
    _YDGVars_Set("sLogPath", $_sLogPath)
    ; On encadre le debut du log
    $__g_bLogInit = True
    _YDLogger_Sep(50, "*")
    _YDLogger_Log("- SCRIPT         : " & @ScriptName & " (" & _YDGVars_Get("sAppVersionV") & ")")
    _YDLogger_Log("- LOG            : " & _YDGVars_Get("sLogPath"))
    _YDLogger_Log("- LOGLEVEL       : " & _YDGVars_Get("sLogLevel"))
    _YDLogger_Log("- COMPUTERNAME   : " & @ComputerName)
    _YDLogger_Log("- USERNAME       : " & @UserName)
    _YDLogger_Sep(50, "*")
    $__g_bLogInit = False
    Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDLogger_Log
; Description ...: Ecrit une ligne de log simple
; Syntax.........: _YDLogger_Log($_sLogMsg, [$_sFuncName])
; Parameters ....: $_sLogMsg     - Message a logger
;                  $_sFuncName      - Nom de la fonction appelante
;                  $_iLogLevel      - Niveau de log
; Return values .: Success          - True
;                  Failure          - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......: 
; Related .......: 
; ===============================================================================================================================
Func _YDLogger_Log($_sLogMsg, $_sFuncName = $YDLOGGER_FUNCNAME, $_iLogLevel = _YDGVars_Get("sLogLevel"))
    ; On ne loggue que si demande
    If $_iLogLevel > _YDGVars_Get("sLogLevel") Then Return False
    ; On affiche un separateur a chaque changement de fonction
    If ($__g_sLastFuncName <> $_sFuncName And $__g_bLogInit = False) Then
        _YDLogger_Sep()
    EndIf
    ; On logge
    __YDLogger_SetLog("[" & $_sFuncName & "] " & $_sLogMsg)
    ; On garde en memoire le nom de la fonction precedente
    $__g_sLastFuncName = $_sFuncName
    Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDLogger_Error
; Description ...: Ecrit une ligne de log d'erreur
; Syntax.........: _YDLogger_Error($_sLogMsg, [$_sFuncName], [$_iError], [$_iExtended], [$_iScriptLineNumber])
; Parameters ....: $_sLogMsg             - Message a logger
;                  $_sFuncName           - Nom de la fonction appelante
;                  $_iError              - Numero de l'erreur
;                  $_iExtended           - Numero de l'erreur extended
;                  $_iScriptLineNumber   - Numero de la ligne du script
; Return values .: Success      - True
;                  Failure      - 0
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......: 
; Related .......: 
; ===============================================================================================================================
Func _YDLogger_Error($_sLogMsg, $_sFuncName = $YDLOGGER_FUNCNAME, Const $_iError = @error, Const $_iExtended = @extended, Const $_iScriptLineNumber = @ScriptLineNumber)
    ; On logge
    _YDLogger_Sep()
    __YDLogger_SetLog("[" & $_sFuncName & "] " & "/!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\ /!\", $YDLOGGER_LOGCASE_ERROR)
    __YDLogger_SetLog("[" & $_sFuncName & "] " & $_sLogMsg, $YDLOGGER_LOGCASE_ERROR)
    __YDLogger_SetLog("[" & $_sFuncName & "] " & "Erreur: " & $_iError & " - Extended: " & $_iExtended & " - ScriptLine: " & $_iScriptLineNumber, $YDLOGGER_LOGCASE_ERROR)
    ; On garde en memoire le nom de la fonction precedente
    $__g_sLastFuncName = $_sFuncName
    Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDLogger_Sep
; Description ...: Ecrit un separateur dans une ligne de log
; Syntax.........: _YDLogger_Sep([$_iSepNumber], [$_sSepCar])
; Parameters ....: $_iSepNumber  - Nombre de caractère du séparateur
;                  $_sSepCar     - Caractère du séparateur
; Return values .: Success      - True
;                  Failure      - 0
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......: 
; Related .......: 
; ===============================================================================================================================
Func _YDLogger_Sep($_iSepNumber = $YDLOGGER_SEP_NUMBER, $_sSepCar = $YDLOGGER_SEP_CAR)
    Local $sSep = ""
    For $i = 1 To $_iSepNumber Step +1
        $sSep = $sSep & $_sSepCar
    Next 
    __YDLogger_SetLog($sSep)
    Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDLogger_Var
; Description ...: Ecrit une ligne de log du type nomvariable = valeurvariable
; Syntax.........: _YDLogger_Var($_sLogVarName, $_sLogVarValue, [$_sFuncName, $_iLogLevel])
; Parameters ....: $_sLogVarName    - Nom de la variable
;                  $_sLogVarValue   - Valeur de la variable
;                  $_sFuncName      - Nom de la fonction appelante
;                  $_iLogLevel      - Niveau de log
; Return values .: Success          - True
;                  Failure          - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......: 
; Related .......: 
; ===============================================================================================================================
Func _YDLogger_Var($_sLogVarName, $_sLogVarValue, $_sFuncName = $YDLOGGER_FUNCNAME, $_iLogLevel = _YDGVars_Get("sLogLevel"))
    ; On ne loggue que si demande
    If $_iLogLevel > _YDGVars_Get("sLogLevel") Then Return False
    ; On affiche un separateur a chaque changement de fonction
    If ($__g_sLastFuncName <> $_sFuncName) Then
        _YDLogger_Sep()
    EndIf
    ; On construit le message a afficher
    Local $sLogMsg = "[" & $_sFuncName & "] " & $_sLogVarName & " = " & $_sLogVarValue
    ; On logge
    __YDLogger_SetLog($sLogMsg)
    ; On garde en memoire le nom de la fonction precedente
    $__g_sLastFuncName = $_sFuncName
    Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDLogger_LogAllGVars
; Description ...: Ecrit une ligne de log pour chaque variable globale declaree
; Syntax.........: _YDLogger_LogAllGVars([$_iLogLevel])
; Parameters ....: 
; Return values .: Success          - True
;                  Failure          - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 20/02/2019
; Remarks .......: 
; Related .......: 
; ===============================================================================================================================
Func _YDLogger_LogAllGVars($_iLogLevel = _YDGVars_Get("sLogLevel"))
    ; On ne loggue que si demande
    If $_iLogLevel > _YDGVars_Get("sLogLevel") Then Return False
    ; On logge toutes les cles-valeurs
    For $vKey In _YDGVars_GetArray()
        _YDLogger_Var($vKey, _YDGVars_Get($vKey))
    Next
    Return True
EndFunc


; #INTERNAL FUNCTIONS# ===========================================================================================================


; #FUNCTION# ====================================================================================================================
; Name...........: __YDLogger_SetLog
; Description ...: Ecrit une ligne de log
; Syntax.........: __YDLogger_SetLog($_sLogMsg, [sLogLevel])
; Parameters ....: $_sLogMsg    - Message a logger
;                  $_sLogCase   - Level du message
; Return values .: Success      - True
;                  Failure      - 0
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......: Internal Use
; Related .......: 
; ===============================================================================================================================
Func __YDLogger_SetLog($_sLogMsg, $_sLogCase = "")
    ; On ne genere pas de log si loglevel = 0 (sortie de la fonction)
    If (_YDGVars_Get("sLogLevel") = 0 And $__g_bLogInit = False) Then Return True
    ; On sort en erreur si le Logger n'a pas été initialisé
    If _YDGVars_Get("sLogPath") = "" Then Exit MsgBox($MB_SYSTEMMODAL, "", "Le Logger n'a pas été initialisé !")
    Local $sLogDateTime = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC    
    ; On gere le cas $YDLOGGER_LOGCASE_ERROR en rajoutant des infos + ConsoleWriteError
    If $_sLogCase = $YDLOGGER_LOGCASE_ERROR Then
        ConsoleWriteError($sLogDateTime & " : " &  $_sLogMsg & @CRLF)
        Sleep(500)
    Else
        ConsoleWrite($sLogDateTime & " : " &  $_sLogMsg & @CRLF)
    EndIf
    ; On ecrit la log dans le fichier si defini
    If ($__g_sLogType = $YDLOGGER_LOGTYPE_FILE Or $__g_sLogType = $YDLOGGER_LOGTYPE_BOTH) Then
        ; Boucle pour gerer l ecriture par plusieurs utilisateurs en meme temps
        Local $iCount = 0
        While 1 And $iCount < 100
            If _FileWriteLog(_YDGVars_Get("sLogPath"), $_sLogMsg) Then ExitLoop
            Sleep(20)
            $iCount += 1
        WEnd
    EndIf
    Return True
EndFunc