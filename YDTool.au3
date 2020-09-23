#include-once

; #INDEX# =======================================================================================================================
; Title .........: YDTool
; AutoIt Version : 3.3.14.5
; Language ......: French
; Description ...: Script .au3 développé pour gérer les outils communs
; Author(s) .....: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================

; #ENVIRONMENT# =================================================================================================================
; Settings
#AutoIt3Wrapper_Run_AU3Check=Y
#AutoIt3Wrapper_AU3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
; Includes YD
#include <YDGVars.au3>
#include <YDLogger.au3>
; Includes Constants
#include <MsgBoxConstants.au3>
#include <GUIConstants.au3>
; Includes
#include <Misc.au3>
#include <Array.au3>
#include <Date.au3>
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
_YDGVars_Set("sProgramFilesPath", RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir"))
_YDGVars_Set("sPulseRootPathServer", "\\55.126.5.250\bei\PULSE2")
_YDGVars_Set("sPulseAntoolsPathServer", _YDGVars_Get("sPulseRootPathServer") & "\_antools")
_YDGVars_Set("sPsExecExeFullPathServer", _YDGVars_Get("sPulseAntoolsPathServer") & "\" & "PsExec.exe")
_YDGVars_Set("sAntoolsPathLocal", "C:\APPLILOC\_antools")
_YDGVars_Set("sAntoolsPathLocalAdminShare", _YDTool_ConvertPathToAdminShare(_YDGVars_Get("sAntoolsPathLocal")))
_YDGVars_Set("sPulseLogPathLocal", "C:\PMF\RAPPINST\_pulse.log")
_YDGVars_Set("sPulseLogPathLocalAdminShare", _YDTool_ConvertPathToAdminShare(_YDGVars_Get("sPulseLogPathLocal")))
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_IsPing
; Description ...: Lance un Ping sur un Host donne et renvoi si ping OK ou pas
; Syntax.........: _YDTool_IsPing($_sHost)
; Parameters ....: $_sHost      - Nom du PMF ou IP
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_IsPing($_sHost)
    Local $sFuncName = "_YDTool_IsPing"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    Local $iPing = Ping($_sHost, 150)
    If $iPing = 0 Then
        _YDLogger_Log("Pas de reponse au ping", $sFuncName)
        Return False
    Else
        _YDLogger_Log("Reponse au ping en " & $iPing & " ms", $sFuncName)
        Return True
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_IsAdminShare
; Description ...: Verifie si l'utilisateur a acces au partage administratif sur un Host donne
; Syntax.........: _YDTool_IsAdminShare($_sHost)
; Parameters ....: $_sHost       - Nom du PMF ou IP
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 01/10/2019
; Remarks .......: DEPRECATED !! -> Remplacé par _YDTool_IsHostAdminShare()
; Related .......:
; ===============================================================================================================================
Func _YDTool_IsAdminShare($_sHost)
    Local $sFuncName = "_YDTool_CheckAdministrativeShare"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    If FileExists("\\" & $_sHost & "\c$") = 0 Then
        _YDLogger_Error("Partage administratif non accessible !", $sFuncName)
        Return False
    Else
        _YDLogger_Log("Partage administratif accessible", $sFuncName)
        Return True
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_Is7zipInstalled
; Description ...: Verifie si 7zip est bien installe sur un Host donne
; Syntax.........: _YDTool_Is7zipInstalled($_sHost)
; Parameters ....: $_sHost       - Nom du PMF ou IP
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_Is7zipInstalled($_sHost)
    Local $sFuncName = "_YDTool_Is7zipInstalled"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    Local $s7zExeFullPath = "\\" & $_sHost & "\" & _YDTool_Get7zipExeFullPath(True)
    _YDLogger_Var("$s7zExeFullPath", $s7zExeFullPath, $sFuncName)
    If FileExists($s7zExeFullPath) = 0 Then
        _YDLogger_Error("7zip non installe", $sFuncName)
        Return False
    Else
        _YDLogger_Log("7zip installe", $sFuncName)
        Return True
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_IsOpenSshInstalled
; Description ...: Verifie si l'agent OpenSSH est bien installe sur un Host donne
; Syntax.........: _YDTool_IsOpenSshInstalled($_sHost)
; Parameters ....: $_sHost           - Nom du PMF ou IP
;                  $_sOpenSshClient  - Nom du client OpenSSH (Nytrio | Mandriva)
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_IsOpenSshInstalled($_sHost, $_sOpenSshClient = "Nytrio")
    Local $sFuncName = "_YDTool_IsOpenSshInstalled"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    Local $sOpenSshExeFullPathLocal = _YDTool_GetOpenSshExeFullPath(True, $_sOpenSshClient)
    _YDLogger_Var("$sOpenSshExeFullPathLocal", $sOpenSshExeFullPathLocal, $sFuncName)
    If FileExists("\\" & $_sHost & "\" & $sOpenSshExeFullPathLocal) = 0 Then
        _YDLogger_Error("Client OpenSSH " & $_sOpenSshClient & " non installe", $sFuncName)
        Return False
    Else
        _YDLogger_Log("Client OpenSSH " & $_sOpenSshClient & " installe", $sFuncName)
        Return True
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_IsServiceRunning
; Description ...: Verifie si un service Windows est actif sur un Host donne
; Syntax.........: _YDTool_IsServiceRunning($_sHost, $_sServiceName, [$_sFilterKey], [$_sFilterValue])
; Parameters ....: $_sHost           - Nom du PMF ou IP
;                  $_sServiceName    - Nom du service Windows
;                  $_sFilterKey      - Clé du filtre (ex : Caption)
;                  $_sFilterValue    - Valeur du filtre (ex : Nytrio SSH agent)
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_IsServiceRunning($_sHost, $_sServiceName, $_sFilterKey = "", $_sFilterValue = "")
    Local $sFuncName = "_YDTool_IsServiceRunning"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    _YDLogger_Var("$_sServiceName", $_sServiceName, $sFuncName, 2)
    _YDLogger_Var("$_sFilterKey", $_sFilterKey, $sFuncName, 2)
    _YDLogger_Var("$_sFilterValue", $_sFilterValue, $sFuncName, 2)
    Local $bReturn = False
    If _YDTool_IsPing($_sHost) Then
        Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sHost & "\root\cimv2")
        If Not @error Then
            Local $sQuery = "SELECT State FROM Win32_Service WHERE Name='" & $_sServiceName & "'"
            If ($_sFilterKey <> "") Then
                $sQuery = $sQuery & " AND " & $_sFilterKey & "='" & $_sFilterValue & "'"
            EndIf
             _YDLogger_Var("$sQuery", $sQuery, $sFuncName)
            Local $colItems = $objWMIService.ExecQuery($sQuery, "WQL", 0x30)
            If IsObj($colItems) Then
                For $objItem In $colItems
                    If $objItem.State = "Running" Then
                        $bReturn = True
                    EndIf
                Next
            Endif
        EndIf
    EndIf
    _YDLogger_Var("$bReturn", $bReturn, $sFuncName)
    Return $bReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_IsProcessRunning
; Description ...: Verifie si un process Windows est actif sur un Host donne
; Syntax.........: _YDTool_IsProcessRunning($_sHost, $_sProcessName, [$_sPathItem])
; Parameters ....: $_sHost           - Nom du PMF ou IP
;                  $_sServiceName    - Nom du processus Windows
;                  $_sPathItem       - Item present dans le chemin de l'executable
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_IsProcessRunning($_sHost, $_sProcessName, $_sPathItem = "")
    Local $sFuncName = "_YDTool_IsProcessRunning"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    _YDLogger_Var("$_sProcessName", $_sProcessName, $sFuncName, 2)
    Local $bReturn = False
    If _YDTool_IsPing($_sHost) Then
        Local $oWMI = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sHost & "\root\cimv2")
        If Not @error Then
            Local $oProcessList = $oWMI.ExecQuery ("SELECT * FROM Win32_Process Where Name = '" & $_sProcessName & "'", "WQL", 0x30)
            If IsObj($oProcessList) Then
                For $sProcess in $oProcessList
                    _YDLogger_Var("$sProcess.Name", $sProcess.Name, $sFuncName, 2)
                    _YDLogger_Var("$sProcess.ExecutablePath", $sProcess.ExecutablePath, $sFuncName, 2)
                    If $_sPathItem == "" Or StringInStr($sProcess.ExecutablePath, $_sPathItem) Then
                        $bReturn = True
                    EndIf
                Next
            Endif
        EndIf
    EndIf
    _YDLogger_Var("$bReturn", $bReturn, $sFuncName)
    Return $bReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_ExecuteCommandWithPsExec
; Description ...: Lance une commande PsExec sur un Host donne
; Syntax.........: _YDTool_ExecuteCommandWithPsExec($_sHost, $_sCommand, $_sPsExecFullPath)
; Parameters ....: $_sHost              - Nom du PMF ou IP
;                  $_sCommand           - Commande a lancer
;                  $_sPsExecFullPath    - Path de l'executable PsExec
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_ExecuteCommandWithPsExec($_sHost, $_sCommand, $_sPsExecFullPath)
    Local $sFuncName = "_YDTool_ExecuteCommandWithPsExec"
    _YDLogger_Var("$_sPsExecFullPath", $_sPsExecFullPath, $sFuncName, 2)
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    _YDLogger_Var("$_sCommand", $_sCommand, $sFuncName, 2)
    Local $bReturn = False
    ; On verifie que le le chemin du PsExec est accessible
    If FileExists($_sPsExecFullPath) = 0 Then
        _YDLogger_Error($_sPsExecFullPath & " innacessible !", $sFuncName)
        Return False
    Endif
    ; On verifie que le Host est accessible
    If Not _YDTool_IsPing($_sHost) Then Return False
    ; Si la commande est un chemin, on verifie qu il est bien accessible en mode partage administratif
    If StringInStr($_sCommand, "\") > 0 And StringInStr($_sCommand, "cmd.exe") = 0  Then
        If FileExists("\\" & $_sHost & "\" & _YDTool_ConvertPathToAdminShare($_sCommand)) = 0 Then
            _YDLogger_Error($_sCommand & " innacessible !", $sFuncName)
            Return False
        Endif
    EndIf
    ; Tout semble OK, on prepare la commandesyntaxe complete de la commande a lancer
    Local $sCommandToRun = $_sPsExecFullPath & " \\" & $_sHost & " -s -h -accepteula -nobanner " & $_sCommand
    _YDLogger_Var("$sCommandToRun", $sCommandToRun, $sFuncName)
    ; On lance la commande !!!
    Local $iReturn = RunWait($sCommandToRun, '', @SW_HIDE)
    _YDLogger_Var("$iReturn", $iReturn, $sFuncName)
    ; Gestion des Return
    If (@error <> 0 And $iReturn > 0) Then
        _YDLogger_Error("Echec du lancement de la commande : " & $sCommandToRun, $sFuncName)
        $bReturn =  False
    Else
        _YDLogger_Log("Lancement de la commande : " & $sCommandToRun, $sFuncName)
        $bReturn =  True
    EndIf
    _YDLogger_Var("$bReturn", $bReturn, $sFuncName)
    Return $bReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_IsPulseLogSuccess
; Description ...: Verifier si une erreur est survenue dans le _pulse.log
; Syntax.........: _YDTool_IsPulseLogSuccess($_sPulseLogFullPath)
; Parameters ....: $_sPulseLogPath - Chemin du fichier de log
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_IsPulseLogSuccess($_sPulseLogFullPath)
    Local $sFuncName = "_YDTool_IsPulseLogSuccess"
    _YDLogger_Var("$_sPulseLogFullPath", $_sPulseLogFullPath, $sFuncName, 2)
    Local $bReturn = False
    ; On sort si fichier inaccessible
    If FileExists($_sPulseLogFullPath) = 0 Then
        _YDLogger_Error($_sPulseLogFullPath & " inaccessible !", $sFuncName)
        Return False
    EndIf
    ; On compte le nombre de lignes du fichier
    Local $iFileCountLines = _FileCountLines($_sPulseLogFullPath)
    _YDLogger_Var("$iFileCountLines", $iFileCountLines, $sFuncName)
    ; On ouvre le fichier en lecture
    Local $hFileOpen = FileOpen($_sPulseLogFullPath, $FO_READ)
    If $hFileOpen = -1 Then
        _YDLogger_Error("Impossible d'ouvrir le fichier : " & $_sPulseLogFullPath)
        Return False
    EndIf
    ; On lit l'avant-derniere ligne
    Local $sLine = FileReadLine($hFileOpen, _FileCountLines($_sPulseLogFullPath) -1)
     _YDLogger_Var("$sLine", $sLine, $sFuncName)
    ; On ferme le fichier
    FileClose($hFileOpen)
    ; Gestion des retours
    If StringInStr($sLine, "(0:0)") > 0 Then
        $bReturn = True
    Else
        $bReturn = False
    EndIf
    _YDLogger_Var("$bReturn", $bReturn, $sFuncName)
    Return $bReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_IsSingleton
; Description ...: Verifier si l application est lancee singleton
; Syntax.........: _YDTool_IsSingleton()
; Parameters ....:
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 21/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_IsSingleton()
    Local $sFuncName = "_YDTool_IsSingleton"
    Local $bReturn = False
    If _Singleton(_YDGVars_Get("sAppName"), 1) = 0 Then
        Local $sMsg = "L'application " & _YDGVars_Get("sAppName") & " est déjà en cours d'exécution !"
        _YDTool_SetMsgBoxError($sMsg, $sFuncName)
        $bReturn = False
    Else
        $bReturn = True
    EndIf
    _YDLogger_Var("$bReturn", $bReturn, $sFuncName)
    Return $bReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_ExitConfirm
; Description ...: Confirmation de fermeture via le tray
; Syntax.........: _YDTool_ExitConfirm()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 21/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_ExitConfirm()
    Local $sFuncName = "_YDTool_ExitConfirm"
 	Local $iResponse = MsgBox(4, _YDGVars_Get("sAppName"),"Etes-vous sûr de vouloir quitter " & _YDGVars_Get("sAppName") & " ?",30)
    If $iResponse = 6 Then
        _YDLogger_Log("Fermeture confirmee", $sFuncName)
        Exit
    Else
        _YDLogger_Log("Fermeture annulee", $sFuncName)
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_ExitApp
; Description ...: Fonction uniquement lancee a la fermture du programme
; Syntax.........: _YDTool_ExitApp()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 21/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_ExitApp()
    Local $sFuncName = "_YDTool_ExitApp"
    Local $sMsg
    Switch (@ExitMethod)
        Case 0
            $sMsg = "Fermeture normale"
        Case 1
            $sMsg = "Fermeture via appel a la fonction Exit"
        Case 2
            $sMsg = "Fermeture via la un clic sur Fermer dans la zone de notification"
        Case 3
            $sMsg = "Fermeture via deconnexion utilisateur : " & @UserName
        Case 4
            $sMsg = "Fermeture via arret de Windows"
        Case Else
            $sMsg = "Fermeture anormale !"
    EndSwitch
    If _YDGVars_Get("sLogPath") = "" Then
        ConsoleWriteError($sMsg)
    Else
        _YDLogger_Log($sMsg, $sFuncName)
        _YDLogger_Log("", $sFuncName)
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_ConvertPathToAdminShare
; Description ...: Convertit un chemin au format "partage administratif"
; Syntax.........: _YDTool_ConvertPathToAdminShare($_sLocalPath)
; Parameters ....: $_sLocalPath - Chemin local
; Return values .: $bPathAdminShareReturn
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 01/10/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_ConvertPathToAdminShare($_sLocalPath)
    Local $sFuncName = "_YDTool_ConvertPathToAdminShare"
    _YDLogger_Var("$_sLocalPath", $_sLocalPath, $sFuncName, 2)
    Local $bPathAdminShareReturn = StringReplace($_sLocalPath, ":", "$")
    _YDLogger_Var("$bPathAdminShareReturn", $bPathAdminShareReturn, $sFuncName)
    Return $bPathAdminShareReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetProgramFilesPath
; Description ...: Renvoi le Path du dossier Program Files (selon architecture du PMF)
; Syntax.........: _YDTool_GetProgramFilesPath([$_bAdminShare])
; Parameters ....: $_bAdminShare - Vrai s il s agit d un partage administratif
; Return values .: $bProgramFilesPathReturn
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetProgramFilesPath($_bAdminShare = False)
    Local $sFuncName = "_YDTool_GetProgramFilesPath"
    _YDLogger_Var("$_bAdminShare", $_bAdminShare, $sFuncName, 2)
    Local $bProgramFilesPathReturn = ($_bAdminShare = True) ? _YDTool_ConvertPathToAdminShare(_YDGVars_Get("sProgramFilesPath")) : _YDGVars_Get("sProgramFilesPath")
    _YDLogger_Var("$bProgramFilesPathReturn", $bProgramFilesPathReturn, $sFuncName)
    Return $bProgramFilesPathReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetAcrobatReaderLastRecentFile
; Description ...: Renvoi le dernier fichier ouvert dans Acrobat Reader
; Syntax.........: _YDTool_GetAcrobatReaderLastRecentFile()
; Parameters ....:
; Return values .: $sReturn
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetAcrobatReaderLastRecentFile()
    Local $sFuncName = "_YDTool_GetAcrobatReaderLastRecentFile"
    Local $sVersion = _YDTool_GetAcrobatReaderVersionName()
    Local $sLastRecentFile = RegRead("HKEY_CURRENT_USER\Software\Adobe\Acrobat Reader\" & $sVersion & "\AVGeneral\cRecentFiles\c1", "tDIText")
    Local $sReturn = ""
    If StringLeft($sLastRecentFile,1) = "/" Then
        $sReturn = StringReplace(StringMid($sLastRecentFile,2, 1) & ":" & StringMid($sLastRecentFile,3), "/", "\")
    EndIf
    _YDLogger_Var("$sReturn", $sReturn, $sFuncName)
    Return $sReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetAcrobatReaderVersionName
; Description ...: Renvoi la version actuellement installe de Acrobat Reader
; Syntax.........: _YDTool_GetAcrobatReaderVersionName()
; Parameters ....:
; Return values .: $sReturn
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetAcrobatReaderVersionName()
    Local $sFuncName = "_YDTool_GetAcrobatReaderVersionName"
    Local $sReturn = RegEnumKey("HKEY_LOCAL_MACHINE\SOFTWARE\Adobe\Acrobat Reader", 1)
    _YDLogger_Var("$sReturn" ,$sReturn, $sFuncName)
    Return $sReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_Get7zipExeFullPath
; Description ...: Renvoi le Path du programme 7zip.exe (selon architecture du PMF)
; Syntax.........: _YDTool_Get7zipExeFullPath([$_bAdminShare])
; Parameters ....: $_bAdminShare - Vrai s il s agit d un partage administratif
; Return values .: $sReturn
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_Get7zipExeFullPath($_bAdminShare = False)
    Local $sFuncName = "_YDTool_Get7zipExeFullPath"
    _YDLogger_Var("$_bAdminShare", $_bAdminShare, $sFuncName, 2)
    Local $sProgramFilesPath = _YDTool_GetProgramFilesPath($_bAdminShare)
    _YDLogger_Var("$sProgramFilesPath", $sProgramFilesPath, $sFuncName)
    Local $sReturn = $sProgramFilesPath & "\7-zip\7z.exe"
    _YDLogger_Var("$sReturn", $sReturn, $sFuncName)
    Return $sReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetOpenSshExeFullPath
; Description ...: Renvoi le Path du programme cygrunsrv.exe (selon architecture du PMF)
; Syntax.........: _YDTool_GetOpenSshExeFullPath([$_bAdminShare], [$_sOpenSshClient])
; Parameters ....: $_bAdminShare     - Vrai s il s agit d un partage administratif
;                  $_sOpenSshClient  - Nom du client OpenSSH (Nytrio | Mandriva)
; Return values .: $sReturn
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetOpenSshExeFullPath($_bAdminShare = False, $_sOpenSshClient = "Nytrio")
    Local $sFuncName = "_YDTool_GetOpenSshExeFullPath"
    _YDLogger_Var("$_bAdminShare", $_bAdminShare, $sFuncName, 2)
    Local $sReturn = _YDTool_GetProgramFilesPath($_bAdminShare) & "\" & $_sOpenSshClient & "\OpenSSH\bin\cygrunsrv.exe"
    _YDLogger_Var("$sReturn", $sReturn, $sFuncName)
    Return $sReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_CreateFolderIfNotExist
; Description ...: Cree un dossier s il n existe pas deja
; Syntax.........: _YDTool_CreateFolderIfNotExist($_sPath)
; Parameters ....: $_sPath - Chemin du dossier a creer
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_CreateFolderIfNotExist($_sPath)
    Local $sFuncName = "_YDTool_CreateFolderIfNotExist"
    _YDLogger_Var("$_sPath", $_sPath, $sFuncName, 2)
    If DirGetSize($_sPath) = -1 And DirCreate($_sPath) = 0 Then
        _YDLogger_Error("Creation dossier impossible : " & $_sPath, $sFuncName)
        Return False
    Else
        _YDLogger_Log("Creation du dossier OK : " & $_sPath, $sFuncName)
        Return True
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_DeleteFolderIfExist
; Description ...: Supprime un dossier s il existe
; Syntax.........: _YDTool_DeleteFolderIfExist($_sPath, [$_iOption])
; Parameters ....: $_sPath       - Chemin du dossier a supprimer
;                  $_iOption     - 0 : dossier vide | 1 : dossier rempli avec recursivite
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_DeleteFolderIfExist($_sPath, $_iOption = 1)
    Local $sFuncName = "_YDTool_DeleteFolderIfExist"
    _YDLogger_Var("$_sPath", $_sPath, $sFuncName, 2)
    If DirGetSize($_sPath) = -1 Then
         _YDLogger_Error($_sPath & " : dossier innexistant !", $sFuncName)
        Return False
    EndIf
    If DirRemove($_sPath, $_iOption) = 0 Then
        _YDLogger_Error("Suppression dossier impossible : " & $_sPath, $sFuncName)
        Return False
    Else
        _YDLogger_Log("Suppression du dossier OK : " & $_sPath, $sFuncName)
        Return True
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_DeleteFileIfExist
; Description ...: Supprime un fichier s il existe
; Syntax.........: _YDTool_DeleteFileIfExist($_sFullPath)
; Parameters ....: $_sFullPath       - Chemin du fichier a supprimer
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_DeleteFileIfExist($_sFullPath)
    Local $sFuncName = "_YDTool_DeleteFileIfExist"
    _YDLogger_Var("$_sPath", $_sFullPath, $sFuncName, 2)
    If FileDelete($_sFullPath) = 0 Then
        _YDLogger_Error("Suppression fichier impossible : " & $_sFullPath, $sFuncName)
        Return False
    Else
        _YDLogger_Log("Suppression du fichier OK : " & $_sFullPath, $sFuncName)
        Return True
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_MoveFile
; Description ...: Déplacement d'un fichier - surcouche de FileMove()
; Syntax.........: _YDTool_MoveFile($_sSourceFullPath, $_sDestinationFullPath, [$_iOption])
; Parameters ....: $_sSourcePath         - Chemin du fichier source
;                  $_sDestinationPath    - Chemin du fichier de destination
;                  $_iOption             - Options pour le déplacement du fichier
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_MoveFile($_sSourceFullPath, $_sDestinationFullPath, $_iOption = $FC_OVERWRITE + $FC_CREATEPATH)
    Local $sFuncName = "_YDTool_MoveFile"
    _YDLogger_Var("$_sSourcePath", $_sSourceFullPath, $sFuncName, 2)
    _YDLogger_Var("$_sDestinationPath", $_sDestinationFullPath, $sFuncName, 2)
    _YDLogger_Var("$_iOption", $_iOption, $sFuncName, 2)
    If FileMove($_sSourceFullPath, $_sDestinationFullPath, $_iOption) = 0 Then
        _YDLogger_Error("Déplacement du fichier impossible : " & $_sDestinationFullPath, $sFuncName)
        Return False
    Else
        _YDLogger_Log("Déplacement du fichier OK : " & $_sDestinationFullPath, $sFuncName)
        Return True
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_CopyFile
; Description ...: Copie d'un fichier - surcouche de FileCopy()
; Syntax.........: _YDTool_CopyFile($_sSourceFullPath, $_sDestinationFullPath, [$_iOption])
; Parameters ....: $_sSourcePath         - Chemin du fichier source
;                  $_sDestinationPath    - Chemin du fichier de destination
;                  $_iOption             - Options pour la copie du fichier
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_CopyFile($_sSourceFullPath, $_sDestinationFullPath, $_iOption = $FC_OVERWRITE + $FC_CREATEPATH)
    Local $sFuncName = "_YDTool_CopyFile"
    _YDLogger_Var("$_sSourcePath", $_sSourceFullPath, $sFuncName, 2)
    _YDLogger_Var("$_sDestinationPath", $_sDestinationFullPath, $sFuncName, 2)
    _YDLogger_Var("$_iOption", $_iOption, $sFuncName, 2)
    If FileCopy($_sSourceFullPath, $_sDestinationFullPath, $_iOption) = 0 Then
        _YDLogger_Error("Copie du fichier impossible : " & $_sDestinationFullPath, $sFuncName)
        Return False
    Else
        _YDLogger_Log("Copie du fichier OK : " & $_sDestinationFullPath, $sFuncName)
        Return True
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_CopyFileWithPsExec
; Description ...: Copie d'un fichier - surcouche de FileCopy()
; Syntax.........: _YDTool_CopyFileWithPsExec($_sHost, $_sSourceFullPath, $_sDestinationFullPath)
; Parameters ....: $_sHost				  - Nom du PMF ou IP
;				   $_sSourcePath         - Chemin du fichier source
;                  $_sDestinationPath    - Chemin du fichier de destination
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_CopyFileWithPsExec($_sHost, $_sSourceFullPath, $_sDestinationFullPath)
    Local $sFuncName = "_YDTool_CopyFileWithPsExec"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    _YDLogger_Var("$_sSourcePath", $_sSourceFullPath, $sFuncName, 2)
    _YDLogger_Var("$_sDestinationPath", $_sDestinationFullPath, $sFuncName, 2)
    Local $sCommandToRun = @ComSpec & ' /c "copy /y ' & $_sSourceFullPath & ' ' & $_sDestinationFullPath & '"'
    _YDLogger_Var("$sCommandToRun", $sCommandToRun, $sFuncName, 2)
    If Not _YDTool_ExecuteCommandWithPsExec($_sHost, $sCommandToRun, _YDGVars_Get("sPsExecExeFullPathServer")) Then
        _YDLogger_Error("Copie du fichier via PsExec impossible : " & $_sDestinationFullPath, $sFuncName)
        Return False
    Else
        _YDLogger_Log("Copie du fichier via PsExec OK : " & $_sDestinationFullPath, $sFuncName)
        Return True
    Endif
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_ExtractArchive
; Description ...: Extrait une archive en local ou a partir d'un partage administratif
; Syntax.........: _YDTool_ExtractArchive($_sSourceFullPath, $_sDestinationPath)
; Parameters ....: $_sSourceFullPath        - Chemin du fichier source à extraire
;                  $_sDestinationPath       - Chemin du fichier de destination
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_ExtractArchive($_sSourceFullPath, $_sDestinationPath)
    Local $sFuncName = "_YDTool_ExtractArchive"
    _YDLogger_Var("$_sSourceFullPath", $_sSourceFullPath, $sFuncName, 2)
    _YDLogger_Var("$_sDestinationPath", $_sDestinationPath, $sFuncName, 2)
    Local $Local7zipExeFullPath = _YDTool_Get7zipExeFullPath(False)
    _YDLogger_Var("$Local7zipExeFullPath", $Local7zipExeFullPath, $sFuncName)
    Local $sCommandToRun = @ComSpec & " /c " & ' "' & $Local7zipExeFullPath & '" x -y ' & $_sSourceFullPath & ' -o' & $_sDestinationPath & '\'
    _YDLogger_Var("$sCommandToRun", $sCommandToRun, $sFuncName)
    Local $iReturn = RunWait($sCommandToRun, '', @SW_HIDE)
    _YDLogger_Var("$iReturn", $iReturn)
    If (@error) Then
        _YDLogger_Error("Extraction du fichier impossible : " & $_sDestinationPath, $sFuncName)
        Return False
    Else
        _YDLogger_Log("Extraction du fichier OK : " & $_sDestinationPath, $sFuncName)
        Return True
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_ExtractArchiveWithPsExec
; Description ...: Extrait une archive en local ou a partir d'un partage administratif
; Syntax.........: _YDTool_ExtractArchiveWithPsExec($_sSourceFullPath, $_sDestinationPath, [$_sHost])
; Parameters ....: $_sHost					- Nom du PMF ou IP
; 				   $_sSourceFullPath        - Chemin du fichier source à extraire
;                  $_sDestinationPath       - Chemin du fichier de destination
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 24/06/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_ExtractArchiveWithPsExec($_sHost, $_sSourceFullPath, $_sDestinationPath)
    Local $sFuncName = "_YDTool_ExtractArchiveWithPsExec"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    _YDLogger_Var("$_sSourceFullPath", $_sSourceFullPath, $sFuncName, 2)
    _YDLogger_Var("$_sDestinationPath", $_sDestinationPath, $sFuncName, 2)
    Local $Local7zipExeFullPath = _YDTool_Get7zipExeFullPath(False)
    _YDLogger_Var("$Local7zipExeFullPath", $Local7zipExeFullPath, $sFuncName)
    Local $sCommandToRun = @ComSpec & " /c " & ' "' & $Local7zipExeFullPath & '" x -y ' & $_sSourceFullPath & ' -o' & $_sDestinationPath & '\'
    _YDLogger_Var("$sCommandToRun", $sCommandToRun, $sFuncName)
    If Not _YDTool_ExecuteCommandWithPsExec($_sHost, $sCommandToRun, _YDGVars_Get("sPsExecExeFullPathServer")) Then
        _YDLogger_Error("Extraction du fichier via PsExec impossible : " & $_sDestinationPath, $sFuncName)
        Return False
    Else
        _YDLogger_Log("Extraction du fichier via PsExec OK : " & $_sDestinationPath, $sFuncName)
        Return True
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_SetMsgBoxError
; Description ...: Affiche un MsgBox d'erreur (avec log)
; Syntax.........: _YDTool_SetMsgBoxError($_sMsg, [$_sFuncName])
; Parameters ....: $_sMsg       - Message a afficher
;                  $_sFuncName  - Nom de la fonctione appelante
; Return values .: False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_SetMsgBoxError($_sMsg, $_sFuncName = "")
    Local $sFuncName = "_YDTool_SetMsgBoxError"
    If $_sFuncName = "" Then
        $_sFuncName = $sFuncName
    EndIf
    MsgBox($MB_SYSTEMMODAL + $MB_ICONERROR + $MB_OK, "ERREUR", $_sMsg)
    _YDLogger_Error($_sMsg, $_sFuncName)
    Return False
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_SetTrayTip
; Description ...: Affiche un message dans le TrayTip
; Syntax.........: _YDTool_SetTrayTip($_sTitle, $_sMsg, [$_iTimeout], [$_iOption])
; Parameters ....: $_sTitle     - Titre a afficher dans le TrayTip
;                  $_sMsg       - Message a afficher dans le TrayTip
;                  $_iTimeout   - Duree du Traytip (en ms)
;                  $_iOption    - Icone
; Return values .: False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_SetTrayTip($_sTitle, $_sMsg, $_iTimeout = 3000, $_iOption = 0)
    Local $sFuncName = "_YDTool_SetTrayTip"
    TrayTip($_sTitle, $_sMsg, 0, $_iOption)
    _YDLogger_Var("$_sMsg", $_sMsg, $sFuncName)
    If ($_iTimeout > 0) Then
        Sleep($_iTimeout)
        TrayTip("", "", 0, $_iOption)
    Endif
    Return False
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetHostName
; Description ...: Recupere le nom du PMF pour une IP donnee
; Syntax.........: _YDTool_GetHostName($_sIP)
; Parameters ....: $_sIP     	- IP
; Return values .: Success      - $sHostNameReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 01/10/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetHostName($_sIP)
    Local $sFuncName = "_YDTool_GetHostName"
    _YDLogger_Var("$_sIP", $_sIP, $sFuncName, 2)
    Local $sHostNameReturn = ""
    If _YDTool_IsPing($_sIP) Then
        Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sIP & "\root\cimv2")
        If Not @error Then
            Local $colItems = $objWMIService.ExecQuery("SELECT SystemName FROM Win32_NetworkAdapter WHERE NetConnectionStatus=2", "WQL", 0x30)
            If IsObj($colItems) Then
                For $objItem In $colItems
                    If $objItem.SystemName <> "" Then
                        $sHostNameReturn = $objItem.SystemName
                    EndIf
                Next
            Endif
        EndIf
    EndIf
    _YDLogger_Var("$sHostNameReturn", $sHostNameReturn, $sFuncName)
    Return $sHostNameReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetHostLoggedUserName
; Description ...: Recupere le nom de l'utilisateur connecte au PMF pour un Host donne
; Syntax.........: _YDTool_GetHostLoggedUserName($_sHost)
; Parameters ....: $_sHost      - Nom du PMF ou IP
; Return values .: Success      - $sHostLoggedUserNameReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 01/10/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetHostLoggedUserName($_sHost)
    Local $sFuncName = "_YDTool_GetHostLoggedUserName"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    Local $sHostLoggedUserNameReturn = ""
    If _YDTool_IsPing($_sHost) Then
        Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sHost & "\root\cimv2")
        If Not @error Then
            Local $colItems = $objWMIService.ExecQuery("SELECT userName FROM Win32_ComputerSystem", "WQL", 0x30)
            If IsObj($colItems) Then
                For $objItem In $colItems
                    If $objItem.userName <> "" Then
                        $sHostLoggedUserNameReturn = $objItem.userName
                    EndIf
                Next
            Endif
        EndIf
    EndIf
    _YDLogger_Var("$sHostLoggedUserNameReturn", $sHostLoggedUserNameReturn, $sFuncName)
    Return $sHostLoggedUserNameReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetHostMacAddress
; Description ...: Recupere l'adresse MAC du Host donne
; Syntax.........: _YDTool_GetHostMacAddress($_sHost)
; Parameters ....: $_sHost      - Nom du PMF ou IP
; Return values .: Success      - $sHostMacReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 01/10/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetHostMacAddress($_sHost)
    Local $sFuncName = "_YDTool_GetHostMacAddress"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    Local $sHostMacReturn = ""
    If _YDTool_IsPing($_sHost) Then
        Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sHost & "\root\cimv2")
        If Not @error Then
            Local $colItems = $objWMIService.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapter WHERE NetConnectionStatus=2", "WQL", 0x30)
            If IsObj($colItems) Then
                For $objItem In $colItems
                    If $objItem.MACAddress <> "" Then
                        $sHostMacReturn = $objItem.MACAddress
                    EndIf
                Next
            Endif
        EndIf
    EndIf
    _YDLogger_Var("$sHostMacReturn", $sHostMacReturn, $sFuncName)
    Return $sHostMacReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetHostIpAddress
; Description ...: Recupere l'adresse IP du Host donne
; Syntax.........: _YDTool_GetHostIpAddress($_sHost)
; Parameters ....: $_sHost		- Nom du PMF ou IP
; Return values .: Success      - $sHostIpReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 01/10/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetHostIpAddress($_sHost)
    Local $sFuncName = "_YDTool_GetHostIpAddress"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    Local $sHostIpReturn = ""
    Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sHost & "\root\cimv2")
    If Not @error Then
        Local $colItems = $objWMIService.ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True", "WQL", 0x30)
        If IsObj($colItems) Then
            For $objItem In $colItems
                If $objItem.IPAddress(0) <> "" Then
                    $sHostIpReturn = $objItem.IPAddress(0)
                EndIf
            Next
        Endif
    Else
        _YDLogger_Log("Impossible d'initialiser l'object $objWMIService", $sFuncName)
        _YDLogger_Log("Tentative de recherche de l'IP par ping ...", $sFuncName)
;~ 		Local $iPID = Run(@ComSpec & " /c nslookup " & $_sHost, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
        Local $iPID = Run(@ComSpec & " /c ping -n 1 " & $_sHost, "", @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
        Local $sOutput=""
        Local $aIPs=""
        While 1
            $sOutput &= StdoutRead($iPID)
            If @error Then ; On sort de la boucle si le process se ferme ou si StdoutRead retourne une erreur
                ExitLoop
            EndIf
        WEnd
;~ 		$aIPs=StringRegExp($sOutput,'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b',3)
        $aIPs=StringRegExp($sOutput,'\[(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\]',3)
        $sHostIpReturn = $aIPs[UBound($aIPs) - 1]
    EndIf
    _YDLogger_Var("$sHostIpReturn", $sHostIpReturn, $sFuncName)
    If $sHostIpReturn = "" Then
        _YDLogger_Error("Impossible de determiner l'adresse IP du host : " & $_sHost, $sFuncName)
    EndIf
    Return $sHostIpReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetHostSite
; Description ...: Recupere le site du Host donne
; Syntax.........: _YDTool_GetHostSite($_sHost)
; Parameters ....: $_sHost      - Nom du PMF ou IP
; Return values .: Success      - $sHostSiteReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 01/10/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetHostSite($_sHost)
    Local $sFuncName = "_YDTool_GetHostSite"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    Local $sHostSiteReturn = ""
    Select
        Case StringMid($_sHost, 12, 1) = "A"
            $sHostSiteReturn = "ARRAS"		; 55.123.4.18 	=> \\W11620101AAF\
        Case StringMid($_sHost, 12, 1) = "E"
            $sHostSiteReturn = "BETHUNE" 	; 55.126.8.18 	=> \\W11620400ACF\
        Case StringMid($_sHost, 12, 1) = "R"
            $sHostSiteReturn = "BRUAY" 		; 55.126.36.10 	=> \\W11620101BRU\
        Case StringMid($_sHost, 12, 1) = "N"
            $sHostSiteReturn = "HENIN" 		; 55.126.16.12 	=> \\W11620400ADF\
        Case StringMid($_sHost, 12, 1) = "L"
            $sHostSiteReturn = "LENS"  		; 55.126.4.49 	=> \\W11620101ACF\
        Case StringMid($_sHost, 12, 1) = "V"
            $sHostSiteReturn = "LIEVIN" 	; 55.126.28.12 	=> \\W11620400AFF\
    EndSelect
    ;~ ; Ancienne methode avec les adresses IP
    ;~ Local $sIP = _YDTool_GetHostIpAddress($_sHost)
    ;~ _YDLogger_Var("$sIP", $sIP, $sFuncName, 2)
    ;~ Local $aIP = StringSplit($sIP, ".")
    ;~ Select
    ;~     Case Int($aIP[1]) = 55 And Int($aIP[2]) = 123 And Int($aIP[3]) >= 4 And Int($aIP[3]) <= 6		; 55.123.4.1 => 55.123.6.255
    ;~         $sHostSiteReturn = "ARRAS"		; 55.123.4.18 	=> \\W11620101AAF\
    ;~     Case Int($aIP[1]) = 55 And Int($aIP[2]) = 126 And Int($aIP[3]) >= 8 And Int($aIP[3]) <= 12		; 55.126.8.1 => 55.126.12.255
    ;~         $sHostSiteReturn = "BETHUNE" 	; 55.126.8.18 	=> \\W11620400ACF\
    ;~     Case Int($aIP[1]) = 55 And Int($aIP[2]) = 126 And Int($aIP[3]) >= 36 And Int($aIP[3]) <= 39		; 55.126.36.1 => 55.126.39.255
    ;~         $sHostSiteReturn = "BRUAY" 		; 55.126.36.10 	=> \\W11620101BRU\
    ;~     Case Int($aIP[1]) = 55 And Int($aIP[2]) = 126 And Int($aIP[3]) >= 16 And Int($aIP[3]) <= 19		; 55.126.16.1 => 55.126.19.255
    ;~         $sHostSiteReturn = "HENIN" 		; 55.126.16.12 	=> \\W11620400ADF\
    ;~     Case Int($aIP[1]) = 55 And Int($aIP[2]) = 126 And Int($aIP[3]) >= 4 And Int($aIP[3]) <= 7		; 55.126.4.1 => 55.126.7.255
    ;~         $sHostSiteReturn = "LENS"  		; 55.126.4.49 	=> \\W11620101ACF\
    ;~     Case Int($aIP[1]) = 55 And Int($aIP[2]) = 126 And Int($aIP[3]) >= 28 And Int($aIP[3]) <= 31		; 55.126.28.1 => 55.126.31.255
    ;~         $sHostSiteReturn = "LIEVIN" 	; 55.126.28.12 	=> \\W11620400AFF\
    ;~ EndSelect
    _YDLogger_Var("$sHostSiteReturn", $sHostSiteReturn, $sFuncName)
    Return $sHostSiteReturn
EndFunc


; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetHostIpSubnet
; Description ...: Recupere l'adresse IP de sous-reseau du Host donne
; Syntax.........: _YDTool_GetHostIpSubnet($_sHost)
; Parameters ....: $_sHost      - Nom du PMF ou IP
; Return values .: Success      - $sHostIpSubnetReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 01/10/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetHostIpSubnet($_sHost)
    Local $sFuncName = "_YDTool_GetHostIpSubnet"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    Local $sHostIpSubnetReturn = ""
    If _YDTool_IsPing($_sHost) Then
        Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sHost & "\root\cimv2")
        If Not @error Then
            Local $colItems = $objWMIService.ExecQuery("SELECT IPSubnet FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True", "WQL", 0x30)
            If IsObj($colItems) Then
                For $objItem In $colItems
                    If $objItem.IPSubnet(0) <> "" Then
                        $sHostIpSubnetReturn = $objItem.IPSubnet(0)
                    EndIf
                Next
            Endif
        EndIf
    EndIf
    _YDLogger_Var("$sHostIpSubnetReturn", $sHostIpSubnetReturn, $sFuncName)
    Return $sHostIpSubnetReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetHostRegValue
; Description ...: Recupere une valeur d'une clee de registre presente sur le Host
; Syntax.........: _YDTool_GetHostRegValue($_sHost, $_iHkey, $_sRegKey, $_sRegName)
; Parameters ....: $_sHost			- Nom du PMF ou IP
;				   $_iHkey			- HKEY principal
;				   $_sRegKey		- Clee ciblee
;				   $_sRegVal		- Nom de la cle
; Return values .: $sRegReturn      - Valeur de retour issue du registre
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 01/10/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetHostRegValue($_sHost, $_iHkey, $_sRegKey, $_sRegName)
    Local $sFuncName = "_YDTool_GetHostRegValue"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    _YDLogger_Var("$_iHkey", $_iHkey, $sFuncName, 2)
    _YDLogger_Var("$_sRegKey", $_sRegKey, $sFuncName, 2)
    _YDLogger_Var("$_sRegName", $_sRegName, $sFuncName, 2)
    Local $sHostRegReturn = ""
    If _YDTool_IsPing($_sHost) Then
        Local $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & $_sHost & "\root\default:StdRegProv")
        If Not @error Then
            $objWMIService.GetStringValue($_iHkey, $_sRegKey, $_sRegName, $sHostRegReturn)
            If $sHostRegReturn == "" Then
                _YDLogger_Error("Cle de registre non trouvée : \\" & $_sHost & "\" & $_iHkey & "\" & $_sRegKey & "\" & $_sRegName, $sFuncName)
            EndIf
        Endif
    EndIf
    _YDLogger_Var("$sHostRegReturn", $sHostRegReturn, $sFuncName)
    Return $sHostRegReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_IsHostAdminShare
; Description ...: Verifie si un partage administratif est accessible
; Syntax.........: _YDTool_IsHostAdminShare($_sHost, [$_sShareName])
; Parameters ....: $_sHost      - Nom du PMF ou IP
; 				   $_sShareName - Nom du partage administratif
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 01/10/2019
; Remarks .......: Remplace _YDTool_IsAdminShare
; Related .......:
; ===============================================================================================================================
Func _YDTool_IsHostAdminShare($_sHost, $_sShareName = "C$")
    Local $sFuncName = "_YDTool_IsHostAdminShare"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    _YDLogger_Var("$_sShareName", $_sShareName, $sFuncName, 2)
    Local $bHostAdminShareReturn = False
    Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sHost & "\root\cimv2")
    If Not @error Then
        Local $objWMIShare = $objWMIService.ExecQuery("SELECT * FROM Win32_Share")
        If Not @error Then
            For $objShare in $objWMIShare
                If $objShare.Name = $_sShareName Then
                    $bHostAdminShareReturn = True
                EndIf
            Next
        Else
            _YDLogger_Error("Impossible d'initialiser l'object $objWMIShare", $sFuncName)
        EndIf
    Else
        _YDLogger_Error("Impossible d'initialiser l'object $objWMIService", $sFuncName)
    EndIf
    _YDLogger_Var("$bHostAdminShareReturn", $bHostAdminShareReturn, $sFuncName)
    Return $bHostAdminShareReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_CreateHostAdminShare
; Description ...: Création d'un partage administratif
; Syntax.........: _YDTool_CreateHostAdminShare($_sHost, [$_sSharePath, $_sShareName, $_sShareDesc])
; Parameters ....: $_sHost      - Nom du PMF ou IP
; 				   $_sSharePath - Chemin du partage administratif
; 				   $_sShareName - Nom du partage administratif
; 				   $_sShareDesc - Description du partage administratif
; Return values .: Success      - $iWMIReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 01/10/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_CreateHostAdminShare($_sHost, $_sSharePath = "C:\", $_sShareName = "C$", $_sShareDesc = "Partage par défaut")
    Local $sFuncName = "_YDTool_CreateHostAdminShare"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    _YDLogger_Var("$_sSharePath", $_sSharePath, $sFuncName, 2)
    _YDLogger_Var("$_sShareName", $_sShareName, $sFuncName, 2)
    _YDLogger_Var("$_sShareDesc", $_sShareDesc, $sFuncName, 2)
    Local $iWMIReturn = -1
    Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sHost & "\root\cimv2")
    If Not @error Then
        Local $objWMIShare = $objWMIService.Get("Win32_Share")
        If Not @error Then
            $iWMIReturn = $objWMIShare.Create($_sSharePath, $_sShareName, 0, True, $_sShareDesc)
            If $iWMIReturn == 0 Then
                _YDLogger_Log("Partage administratif cree : " & $_sShareName, $sFuncName)
            Elseif $iWMIReturn == 22 Then
                _YDLogger_Log("Partage administratif " & $_sShareName & " déjà existant", $sFuncName)
            Else
                _YDLogger_Log("Erreur lors de la creation du partage administratif " & $_sShareName & " : " & $iWMIReturn, $sFuncName)
            EndIf
        Else
            _YDLogger_Error("Impossible d'initialiser l'object $objWMIShare", $sFuncName)
        EndIf
    Else
        _YDLogger_Error("Impossible d'initialiser l'object $objWMIService", $sFuncName)
    EndIf
    _YDLogger_Var("$iWMIReturn", $iWMIReturn, $sFuncName)
    Return $iWMIReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_DeleteHostAdminShare
; Description ...: Suppression d'un partage administratif
; Syntax.........: _YDTool_DeleteHostAdminShare($_sHost, [$_sShareName])
; Parameters ....: $_sHost      - Nom du PMF ou IP
; 				   $_sShareName - Nom du partage administratif
; Return values .: Success      - $iWMIReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 01/10/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_DeleteHostAdminShare($_sHost, $_sShareName = "C$")
    Local $sFuncName = "_YDTool_DeleteHostAdminShare"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    _YDLogger_Var("$_sShareName", $_sShareName, $sFuncName, 2)
    Local $iWMIReturn = -1
    Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sHost & "\root\cimv2")
    If Not @error Then
        Local $objWMIShare = $objWMIService.ExecQuery('SELECT * FROM Win32_Share Where Name="' & $_sShareName & '"')
        If Not @error Then
            For $objShare in $objWMIShare
                $iWMIReturn = $objShare.Delete
                If Not @error Then
                    _YDLogger_Log("Partage administratif supprime : " & $_sShareName, $sFuncName)
                Else
                    _YDLogger_Error("Impossible de supprimer le partage administratif : " & $objShare.Name, $sFuncName)
                EndIf
            Next
        Else
            _YDLogger_Error("Impossible d'initialiser l'object $objWMIShare", $sFuncName)
        EndIf
    Else
        _YDLogger_Error("Impossible d'initialiser l'object $objWMIService", $sFuncName)
    EndIf
    _YDLogger_Var("$iWMIReturn", $iWMIReturn, $sFuncName)
    Return $iWMIReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetAppWrapperRes
; Description ...: Renvoi les infos du Wrapper (que le script soit compile ou pas)
; Syntax.........: _YDTool_GetAppWrapperRes([$_sOption])
; Parameters ....: $_sOption    - ProductVersion | ProductName | Description
; Return values .: Success      - $sReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......: Inspired By http://www.autoitscript.com/forum/index.php?showtopic=96654
; ===============================================================================================================================
Func _YDTool_GetAppWrapperRes($_sOption = "ProductVersion")
    Local $sFuncName = "_YDTool_GetAppWrapperRes"
    Local $sReturn = "", $sOption
    If @Compiled Then
        If $_sOption = "Description" Then
            $sOption = "FileDescription"
        Else
            $sOption = $_sOption
        EndIf
        ;_YDLogger_Var("$sOption", $sOption, $sFuncName)
        $sReturn = FileGetVersion(@ScriptFullPath, $sOption)
    Else
        Local $hScriptFile, $sLine, $bDirectivesNotFound
        $hScriptFile = FileOpen(@ScriptFullPath, 0)
        If $hScriptFile = -1 Then
            _YDTool_SetMsgBoxError("Le fichier " & @ScriptFullPath & " ne peut pas etre lu !", $sFuncName)
        Else
            $bDirectivesNotFound = True
            While 1
                $sLine = FileReadLine($hScriptFile)
                If @error Then
                    _YDTool_SetMsgBoxError("Erreur innatendue lors de la lecture du fichier : " & @ScriptFullPath, $sFuncName)
                    ExitLoop ; may never be used, coz we dont loop till end of script!!!
                EndIf
                If $bDirectivesNotFound Then
                    If StringInStr(StringStripWS($sLine, 3), "#Region ;**** Directives created by AutoIt3Wrapper_GUI ****") Then
                        $bDirectivesNotFound = False
                    EndIf
                Else
                    If StringInStr($sLine, "#EndRegion") Then ExitLoop
                EndIf
                Local $iWrapperValuePos = StringInStr(StringStripWS($sLine, 3), "=")
                Local $sWrapperSchema = StringLeft(StringStripWS($sLine, 3), $iWrapperValuePos)
                Local $sWrapperOption = "#AutoIt3Wrapper_Res_" & $_sOption & "="
                ;_YDLogger_Var("$sWrapperSchema", $sWrapperSchema, $sFuncName)
                ;_YDLogger_Var("$sWrapperOption", $sWrapperOption, $sFuncName)
                If $sWrapperSchema = $sWrapperOption Then
                    ;_YDLogger_Log("Schema trouve : " & $sWrapperOption, $sFuncName)
                    $sReturn = StringTrimLeft($sLine, $iWrapperValuePos)
                    ExitLoop
                EndIf
            WEnd
        EndIf
        FileClose($hScriptFile)
    EndIf
    ;_YDLogger_Var("$sReturn", $sReturn, $sFuncName)
    Return $sReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GUIShowAbout
; Description ...: Renvoi une GUI "A propos"
; Syntax.........: _YDTool_GUIShowAbout()
; Parameters ....:
; Return values .: True
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 21/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GUIShowAbout()
    Local $sFuncName = "_YDTool_GUIShowAbout"
    Local $iAboutWidth = 550
    Local $iAboutHeight = 200
    Local $iAboutOkButtonWidth = 30
    Local $font = "Verdana"
    _YDLogger_Log("", $sFuncName)

    Local $hAboutGUI = GUICreate("A propos", $iAboutWidth, $iAboutHeight, -1, -1, BitOR($WS_POPUP,$WS_CAPTION))
    ; Titre
    GUISetFont(12, $iAboutWidth*2, 0, $font)
    GUICtrlCreateLabel(_YDGVars_Get("sAppName"), 0, 0, $iAboutWidth, -1, BitOr($SS_CENTER,$BS_CENTER))
    ; Description + version
    GUISetFont(9, $iAboutWidth, 0, $font)
    GUICtrlCreateLabel(_YDGVars_Get("sAppDesc"), 0, 40, $iAboutWidth, -1, BitOr($SS_CENTER,$BS_CENTER))
    GUICtrlCreateLabel(_YDGVars_Get("sAppVersionV"), 0, 80, $iAboutWidth, -1, BitOr($SS_CENTER,$BS_CENTER))
    Local $idLinkContact = GUICtrlCreateLabel(_YDGVars_Get("sAppContact"), 0, 120, $iAboutWidth, -1, BitOr($SS_CENTER,$BS_CENTER))
    GUICtrlSetColor(-1, 0x0000FF)
    GUICtrlSetCursor(-1, 0)
    ; Bouton OK
    Local $idOkButton = GUICtrlCreateButton("OK", $iAboutWidth/2-$iAboutOkButtonWidth/2, 160, $iAboutOkButtonWidth, 25, BitOr($BS_MULTILINE,$BS_CENTER))
    ; Affichage GUI
    GUISetState(@SW_SHOW, $hAboutGUI)
    ; Loop GUI
    While 1
        Local $iMsg = GUIGetMsg()
        Select
            Case $iMsg = $idOkButton
                GUIDelete($hAboutGUI)
                ExitLoop
            Case $iMsg = $idLinkContact
                ShellExecute("mailto:"&_YDGVars_Get("sAppContact"))
        EndSelect
        Sleep(50)
    WEnd
    Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetFileRow
; Description ...: Renvoi une ligne contenue dans un fichier a plat
; Syntax.........: _YDTool_GetFileRow($_sFilePath, $_iRow = 1)
; Parameters ....: $_sFilePath  - Chemin du fichier à vérifier
;                  $_iRow       - Ligne a ramener
; Return values .: $sRow        - Contenu de la ligne
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 13/01/2020
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetFileRow($_sFilePath, $_iRow = 1)
    Local $sFuncName = "_YDTool_GetFileRow"
    Local $hFile = FileOpen($_sFilePath, $FO_READ)
    If $hFile = -1 Then
        _YDTool_SetMsgBoxError("Le fichier " & $_sFilePath & " ne peut pas etre lu !", $sFuncName)
        Return
    EndIf
    Local $sRow = FileReadLine($hFile, $_iRow)
    FileClose($hFile)
    Return $sRow
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_WaitUntilFileIsReady
; Description ...: Verifie si une fenetre clignotante est présent pour le handle passe en parametre
; Syntax.........: _YDTool_WaitUntilFileIsReady($_sFilePath)
; Parameters ....: $_sFilePath  - Chemin du fichier à vérifier
;                  $_iSleepTime - Duree d'attente (en ms)
; Return values .: True
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 13/01/2020
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_WaitUntilFileIsReady($_sFilePath, $_iSleepTime = 5000)
    Local $sFuncName = "_YDTool_WaitUntilFileIsReady"
    Local $hFile
    Local $bFileOk = False
    For $i = 1 To $_iSleepTime/10
        $hFile = FileOpen($_sFilePath)
        If $hFile <> -1 Then
            FileClose($hFile)
            $bFileOk = True
            ExitLoop
        EndIf
        Sleep(10)
    Next
    _YDLogger_Var("bFileOk", $bFileOk, $sFuncName)
    Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_WinIsFlash
; Description ...: Verifie si une fenetre clignotante est présent pour le handle passe en parametre
; Syntax.........: _YDTool_WinIsFlash($_hWnd)
; Parameters ....: $_hWnd       - Handle a verifier
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 21/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_WinIsFlash($_hWnd)
    Local $sFuncName = "_YDTool_WinIsFlash"
    If WinActive($_hWnd) Then Return False
    Local $tFLASHWINFO = DllStructCreate("uint;hwnd;dword;uint;dword")
    DllStructSetData($tFLASHWINFO, 1, 20)
    DllStructSetData($tFLASHWINFO, 2, WinGetHandle($_hWnd))
    Local $a = DllCall("user32.dll", "int", "FlashWindowEx", "ptr", DllStructGetPtr($tFLASHWINFO))
    If $a[0] > 0 Then
        _YDLogger_Log("Flash detecte !", $sFuncName)
        Return True
    EndIf
    Return False
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_DeleteOldFiles
; Description ...: Supprime les fichiers d'un répertoire donné selon un nombre de jour
; Syntax.........: _YDTool_DeleteOldFiles($_sPath, $_nbdays[], $_sFilter])
; Parameters ....: $_sPath      - Chemin du dossier (au format D:\monchemin)
;                  $_nbdays     - Nombre de jours de retention
;                  $_sFilter    - Filtre sur les extensions par exemple (*.csv)
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 21/02/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_DeleteOldFiles($_sPath, $_nbdays, $_sFilter = "*" & $YDLOGGER_LOGEXT)
    Local $sFuncName = "_YDTool_DeleteOldFiles"
    Local $sTijd
    _YDLogger_Var("$_sPath", $_sPath, $sFuncName, 2)
    _YDLogger_Var("$_nbdays", $_nbdays, $sFuncName, 2)
    _YDLogger_Var("$_sFilter", $_sFilter, $sFuncName, 2)
    ; On verifie que le chemin est OK
    If FileExists($_sPath) = 0 Then
        _YDLogger_Error("Le répertoire " & $_sPath & " n'existe pas !", $sFuncName)
        Return False
    EndIf
    ; On verifie que le nombre de jour est OK
    If IsNumber($_nbdays) And $_nbdays <= 0 Then
        _YDLogger_Error("Le nombre de jours " & $_nbdays & " est incorrect !", $sFuncName)
        Return False
    EndIf
    ; On verifie que le filtre est OK
    If $_sFilter = "" Then
        _YDLogger_Error("Un filtre doit etre appliqué !", $sFuncName)
        Return False
    EndIf
    ; On recupere les infos des fichiers dans un tableau
    Local $aFiles = _FileListToArray($_sPath, $_sFilter)
    ; On sort si tableau vide ou
    If Not IsArray($aFiles) Or UBound($aFiles) = 0 Then
        _YDLogger_Error("Erreur lors de la creation du tableau !", $sFuncName)
        Return False
    EndIf
    _YDLogger_Var("UBound($aFiles)", UBound($aFiles), $sFuncName, 2)
    ; Si tout est OK, on boucle sur le tableau
    For $i = 1 To UBound($aFiles) - 1
        $sTijd = StringRegExpReplace(FileGetTime($_sPath & "\" & $aFiles[$i], 0, 1), "(.{4})(.{2})(.{2})(.{2})(.{2})(.{2})", "${1}/${2}/${3} ${4}:${5}:${6}") ; Last modified Date
        ;_YDLogger_Var("$sTijd", $sTijd, $sFuncName, 2)
        If _DateDiff('D', $sTijd, _NowCalc()) > $_nbdays Then ; 'D' = Difference in days between the given dates
            _YDTool_DeleteFileIfExist($_sPath & "\" & $aFiles[$i])
        EndIf
    Next
    Return True
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetDefaultPrinter
; Description ...: Recupere le nom de l'imprimante par defaut
; Syntax.........: _YDTool_GetDefaultPrinter($_sHost)
; Parameters ....: $_sHost      - Nom du PMF ou IP
; Return values .: Success      - $sDefaultPrinterReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 01/10/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetDefaultPrinter($_sHost)
    Local $sFuncName = "_YDTool_GetDefaultPrinter"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName, 2)
    Local $sDefaultPrinterReturn = ""
    Local $objWMIService = ObjGet("winmgmts:\\" & $_sHost & "\root\CIMV2")
    Local $colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_Printer", "WQL", 0x10 + 0x20)
    If IsObj($colItems) then
       For $objItem In $colItems
            ;_YDLogger_Var("$objItem.DeviceID", $objItem.DeviceID, $sFuncName)
            If  $objitem.Default <> 0 Then
                $sDefaultPrinterReturn = $objItem.DeviceID
            Endif
        Next
    Endif
    _YDLogger_Var("$sDefaultPrinterReturn", $sDefaultPrinterReturn, $sFuncName)
    Return $sDefaultPrinterReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_SetDefaultPrinter
; Description ...: Met l'imprimante passee en parametre comme imprimante par defaut
; Syntax.........: _YDTool_SetDefaultPrinter($_sPrinter)
; Parameters ....: $_sPrinter   - Nom de l'imprimante
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 30/04/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_SetDefaultPrinter($_sPrinter)
    Local $sFuncName = "_YDTool_SetDefaultPrinter"
    _YDLogger_Var("$_sPrinter", $_sPrinter, $sFuncName, 2)
    RunWait(@ComSpec & " /c RUNDLL32 PRINTUI.DLL,PrintUIEntry /q /y /n " & '"' & $_sPrinter & '"', "", @SW_HIDE)
    If @error Then
        _YDLogger_Error("Erreur lors de la mise a jour de l imprimante par defaut !")
        Return False
    Else
        Return True
    EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetAppConfValue
; Description ...: Recupere la valeur d une cle lue dans le fichier de config.ini
; Syntax.........: _YDTool_GetAppConfValue($_sIniSection, $_sIniKey)
; Parameters ....: $_sIniSection	- Nom de la section
; 				   $_sIniKey		- Nom de la cle
; Return values .: Success          - $sReturn
;                  Failure          - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 30/04/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetAppConfValue($_sIniSection, $_sIniKey)
    Local $sFuncName = "_YDTool_GetAppConfValue"
    Local $sAppConfValueReturn = ""
    If Not FileExists(_YDGVars_Get("sAppConfFile")) Then
        _YDLogger_Error("Fichier introuvable : " & _YDGVars_Get("sAppConfFile"))
    Else
        $sAppConfValueReturn = IniRead(_YDGVars_Get("sAppConfFile"), $_sIniSection, $_sIniKey, "")
        If @error Then
            _YDLogger_Error("Lecture impossible du fichier : " & _YDGVars_Get("sAppConfFile"))
        EndIf
    EndIf
    _YDLogger_Var("$sAppConfValueReturn", $sAppConfValueReturn, $sFuncName)
    Return $sAppConfValueReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetAppConfSection
; Description ...: Recupere un tableau contenant les valeurs d'une section du fichier de config.ini
; Syntax.........: _YDTool_GetAppConfSection($_sIniSection)
; Parameters ....: $_sIniSection	- Nom de la section
; Return values .: $aReturn         - Tableau
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 30/04/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetAppConfSection($_sIniSection)
    Local $sFuncName = "_YDTool_GetAppConfSection"
    Local $aReturn
    If Not FileExists(_YDGVars_Get("sAppConfFile")) Then
        _YDLogger_Error("Fichier introuvable : " & _YDGVars_Get("sAppConfFile"))
    Else
        $aReturn = IniReadSection(_YDGVars_Get("sAppConfFile"), $_sIniSection)
        If @error Then
            _YDLogger_Error("Lecture impossible du fichier : " & _YDGVars_Get("sAppConfFile"))
        EndIf
    EndIf
    _YDLogger_Var("$aReturn", $aReturn, $sFuncName)
    Return $aReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_SetAppConfValue
; Description ...: Modifie la valeur d une cle dans le fichier de config.ini
; Syntax.........: _YDTool_SetAppConfValue($_sIniSection, $_sIniKey, $_sValue)
; Parameters ....: $_sIniSection	- Nom de la section
; 				   $_sIniKey		- Nom de la cle
; 				   $_sValue		    - Nouvelle valeur
; Return values .: Success          - $sReturn
;                  Failure          - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 23/09/2020
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_SetAppConfValue($_sIniSection, $_sIniKey, $_sValue)
    Local $sFuncName = "_YDTool_SetAppConfValue"
    Local $iReturn = 0
    If Not FileExists(_YDGVars_Get("sAppConfFile")) Then
        _YDLogger_Error("Fichier introuvable : " & _YDGVars_Get("sAppConfFile"))
    ElseIf Not _YDTool_GetAppConfValue($_sIniSection, $_sIniKey) Then
        _YDLogger_Error("Cle introuvable : " & $_sIniKey & " dans la section " & $_sIniSection)
    Else
        $iReturn = IniWrite(_YDGVars_Get("sAppConfFile"), $_sIniSection, $_sIniKey, $_sValue)
        If @error Then
            _YDLogger_Error("Ecriture impossible de la cle " &  $_sIniKey & " du fichier : " & _YDGVars_Get("sAppConfFile"))
        EndIf
    EndIf
    _YDLogger_Var("$iReturn", $iReturn, $sFuncName)
    Return $iReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_SuspendProcessSwitch
; Description ...: Suspend/reactive l'execution d'un process
; Syntax.........: _YDTool_SuspendProcessSwitch("LNGP.exe", True)
; Parameters ....: $_iPIDOrName	    - PID ou nom du process
;                  $_bSuspend	    - True pour suspension / False pour reactivation
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 09/12/2019
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_SuspendProcessSwitch($_iPIDOrName, $_bSuspend = True)
    Local $sFuncName = "_YDTool_SuspendProcessSwitch"
    Local $bReturn = False
    Local $iSucess = 0
    Local $iPID = 0
    _YDLogger_Var("$_bSuspend", $_bSuspend, $sFuncName, 2)
    If IsString($_iPIDOrName) Then
        $iPID = ProcessExists($_iPIDOrName)
    EndIf
    If Not $iPID Then
        _YDLogger_Error("Process non trouve : " & $_iPIDOrName, $sFuncName)
        Return $bReturn
    EndIf
    Local $ai_Handle = DllCall("kernel32.dll", 'int', 'OpenProcess', 'int', 0x1f0fff, 'int', False, 'int', $iPID)
    If $_bSuspend Then
        _YDLogger_Log("Suspension du process : " & $iPID, $sFuncName)
        $iSucess = DllCall("ntdll.dll", "int", "NtSuspendProcess", "int", $ai_Handle[0])
    Else
        _YDLogger_Log("Reactivation du process : " & $iPID, $sFuncName)
        $iSucess = DllCall("ntdll.dll", "int", "NtResumeProcess", "int", $ai_Handle[0])
    EndIf
    DllCall('kernel32.dll', 'ptr', 'CloseHandle', 'ptr', $ai_Handle)
    If IsArray($iSucess) Then $bReturn = True
    _YDLogger_Var("$bReturn", $bReturn, $sFuncName, 2)
    Return $bReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_ConvertDateFrToYYYYMMDD
; Description ...: Convertit une date francaise (JJ/MM/AAAA) au format AAAAMMJJ
; Syntax.........: _YDTool_ConvertDateFrToYYYYMMDD('12/04/2020')
; Parameters ....: $_sDate	    - Date à convertir
; Return values .: Date au format AAAAMMJJ
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......: 21/01/2020
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_ConvertDateFrToYYYYMMDD($_sDate)
    Local $sFuncName = "_YDTool_ConvertDateFrToYYYYMMDD"
    Local $sReturn = "19000101"
    _YDLogger_Var("$_sDate", $_sDate, $sFuncName, 2)
    Local $aDate = StringSplit($_sDate, "/")
    If $adate[0] > 0 Then
        $sReturn = $aDate[3] & $aDate[2] & $aDate[1]
    EndIf
    _YDLogger_Var("$sReturn", $sReturn, $sFuncName, 2)
    Return $sReturn
EndFunc
