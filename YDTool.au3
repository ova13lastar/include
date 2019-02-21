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
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
_YDGVars_Set("sProgramFilesPath", RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir"))
_YDGVars_Set("sPulseRootPathServer", "\\55.126.5.250\bei\PULSE2")
_YDGVars_Set("sPulseAntoolsPathServer", _YDGVars_Get("sPulseRootPathServer") & "\_antools")
_YDGVars_Set("sPsExecExeFullPathServer", _YDGVars_Get("sPulseAntoolsPathServer") & "\" & "PsExec.exe")
_YDGVars_Set("sAntoolsPathLocal", "C:\APPLILOC\_antools")
_YDGVars_Set("sAntoolsPathLocalAdminShare", StringReplace(_YDGVars_Get("sAntoolsPathLocal"), ":", "$"))
_YDGVars_Set("sPulseLogPathLocal", "C:\PMF\RAPPINST\_pulse.log")
_YDGVars_Set("sPulseLogPathLocalAdminShare", StringReplace(_YDGVars_Get("sPulseLogPathLocal"), ":", "$"))
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
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName)
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
; Description ...: Verifie si l'utilisateur a acces au partage administrtaif sur un Host donne
; Syntax.........: _YDTool_IsAdminShare($_sHost)
; Parameters ....: $_sHost       - Nom du PMF ou IP
; Return values .: Success      - True
;                  Failure      - False
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_IsAdminShare($_sHost)
    Local $sFuncName = "_YDTool_CheckAdministrativeShare"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName)
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
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName)
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
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName)
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
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName)
    _YDLogger_Var("$_sServiceName", $_sServiceName, $sFuncName)
    _YDLogger_Var("$_sFilterKey", $_sFilterKey, $sFuncName)
    _YDLogger_Var("$_sFilterValue", $_sFilterValue, $sFuncName)
    Local $sReturn = False
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
                        $sReturn = True
                    EndIf
                Next
            Endif
        EndIf
    EndIf
    _YDLogger_Var("$sReturn", $sReturn, $sFuncName)
    Return $sReturn
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
    _YDLogger_Var("$_sPsExecFullPath", $_sPsExecFullPath, $sFuncName)
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName)
    _YDLogger_Var("$_sCommand", $_sCommand, $sFuncName)
    ; On verifie que le le chemin du PsExec est accessible
    If FileExists($_sPsExecFullPath) = 0 Then
        _YDLogger_Error($_sPsExecFullPath & " innacessible !", $sFuncName)
        Return False
    Endif
    ; On verifie que le Host est accessible
    If Not _YDTool_IsPing($_sHost) Then Return False
    ; Si la commande est un chemin, on verifie qu il est bien accessible en mode partage administratif
    If StringInStr($_sCommand, "\") > 0 Then
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
        Return False
    Else
        _YDLogger_Log("Lancement de la commande : " & $sCommandToRun, $sFuncName)
        Return True
    EndIf
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
    _YDLogger_Var("$_sPulseLogFullPath", $_sPulseLogFullPath, $sFuncName)
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
     _YDLogger_Var("$sLine", $sLine)
    ; On ferme le fichier
    FileClose($hFileOpen)
    ; Gestion des retours
    If StringInStr($sLine, "(0:0)") > 0 Then
        Return True
    Else
        Return False
    EndIf
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
    If _Singleton(_YDGVars_Get("sAppName"), 1) = 0 Then
        Local $sMsg = "L'application " & _YDGVars_Get("sAppName") & " est déjà en cours d'exécution !"
        _YDTool_SetMsgBoxError($sMsg, $sFuncName)
        Return False
    Else 
        Return True
    EndIf
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
; Return values .: $sReturn
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_ConvertPathToAdminShare($_sLocalPath)
    Local $sFuncName = "_YDTool_ConvertPathToAdminShare"
    _YDLogger_Var("$_sLocalPath", $_sLocalPath, $sFuncName)
    Local $sReturn = StringReplace($_sLocalPath, ":", "$")
    _YDLogger_Var("$sReturn", $sReturn, $sFuncName)
    Return $sReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetProgramFilesPath
; Description ...: Renvoi le Path du dossier Program Files (selon architecture du PMF)
; Syntax.........: _YDTool_GetProgramFilesPath([$_bAdminShare])
; Parameters ....: $_bAdminShare - Vrai s il s agit d un partage administratif
; Return values .: $sReturn
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetProgramFilesPath($_bAdminShare = False)
    Local $sFuncName = "_YDTool_GetProgramFilesPath"
    _YDLogger_Var("$_bAdminShare", $_bAdminShare, $sFuncName)
    Local $sReturn = ($_bAdminShare = True) ? _YDTool_ConvertPathToAdminShare(_YDGVars_Get("sProgramFilesPath")) : _YDGVars_Get("sProgramFilesPath")
    _YDLogger_Var("$sReturn",$sReturn,  $sFuncName)
    Return $sReturn
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
    _YDLogger_Var("$_bAdminShare", $_bAdminShare, $sFuncName)
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
    _YDLogger_Var("$_bAdminShare", $_bAdminShare, $sFuncName)
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
    _YDLogger_Var("$_sPath", $_sPath, $sFuncName)
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
    _YDLogger_Var("$_sPath", $_sPath, $sFuncName)
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
    _YDLogger_Var("$_sPath", $_sFullPath, $sFuncName)
    If FileDelete($_sFullPath) = 0 Then
        _YDLogger_Error("Suppression fichier impossible : " & $_sFullPath, $sFuncName)
        Return False
    Else
        _YDLogger_Log("Suppression du fichier OK : " & $_sFullPath, $sFuncName)
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
    _YDLogger_Var("$_sSourcePath", $_sSourceFullPath, $sFuncName)
    _YDLogger_Var("$_sDestinationPath", $_sDestinationFullPath, $sFuncName)
    _YDLogger_Var("$_iOption", $_iOption, $sFuncName)
    If FileCopy($_sSourceFullPath, $_sDestinationFullPath, $_iOption) = 0 Then
        _YDLogger_Error("Copie du fichier impossible : " & $_sDestinationFullPath, $sFuncName)
        Return False
    Else
        _YDLogger_Log("Copie du fichier OK : " & $_sDestinationFullPath, $sFuncName)
        Return True
    EndIf
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
    _YDLogger_Var("$_sSourceFullPath", $_sSourceFullPath, $sFuncName)
    _YDLogger_Var("$_sDestinationPath", $_sDestinationPath, $sFuncName)
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
    ;TrayTip($_sTitle, $_sMsg, $_iTimeout, $_iOption)
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
; Parameters ....: $_sIP     - Adresse IP
; Return values .: Success      - $sReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetHostName($_sIP)
    Local $sFuncName = "_YDTool_GetHostName"
    _YDLogger_Var("$_sIP", $_sIP, $sFuncName)
    Local $sReturn = ""
    If _YDTool_IsPing($_sIP) Then
        Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sIP & "\root\cimv2")
        If Not @error Then
            Local $colItems = $objWMIService.ExecQuery("SELECT SystemName FROM Win32_NetworkAdapter WHERE NetConnectionStatus=2", "WQL", 0x30)
            If IsObj($colItems) Then
                For $objItem In $colItems
                    If $objItem.SystemName <> "" Then
                        $sReturn = $objItem.SystemName
                    EndIf
                Next
            Endif
        EndIf
    EndIf
    _YDLogger_Var("$sReturn", $sReturn, $sFuncName)
    Return $sReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetHostLoggedUserName
; Description ...: Recupere le nom de l'utilisateur connecte au PMF pour un Host donne
; Syntax.........: _YDTool_GetHostLoggedUserName($_sHost)
; Parameters ....: $_sHost       - Nom du PMF ou IP
; Return values .: Success      - $sReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetHostLoggedUserName($_sHost)
    Local $sFuncName = "_YDTool_GetHostLoggedUserName"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName)
    Local $sReturn = ""
    If _YDTool_IsPing($_sHost) Then
        Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sHost & "\root\cimv2")
        If Not @error Then
            Local $colItems = $objWMIService.ExecQuery("SELECT userName FROM Win32_ComputerSystem", "WQL", 0x30)
            If IsObj($colItems) Then
                For $objItem In $colItems
                    If $objItem.userName <> "" Then
                        $sReturn = $objItem.userName
                    EndIf
                Next
            Endif
        EndIf
    EndIf
    _YDLogger_Var("$sReturn", $sReturn, $sFuncName)
    Return $sReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetHostMacAddress
; Description ...: Recupere l'adresse MAC du Host donne
; Syntax.........: _YDTool_GetHostMacAddress($_sHost)
; Parameters ....: $_sHost       - Nom du PMF ou IP
; Return values .: Success      - $sReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetHostMacAddress($_sHost)
    Local $sFuncName = "_YDTool_GetHostMacAddress"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName)
    Local $sReturn = ""
    If _YDTool_IsPing($_sHost) Then
        Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sHost & "\root\cimv2")
        If Not @error Then
            Local $colItems = $objWMIService.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapter WHERE NetConnectionStatus=2", "WQL", 0x30)
            If IsObj($colItems) Then
                For $objItem In $colItems
                    If $objItem.MACAddress <> "" Then
                        $sReturn = $objItem.MACAddress
                    EndIf
                Next
            Endif
        EndIf
    EndIf
    _YDLogger_Var("$sReturn", $sReturn, $sFuncName)
    Return $sReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetHostIpAddress
; Description ...: Recupere l'adresse IP du Host donne
; Syntax.........: _YDTool_GetHostIpAddress($_sHost)
; Parameters ....: $_sHost       - Nom du PMF ou IP
; Return values .: Success      - $sReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetHostIpAddress($_sHost)
    Local $sFuncName = "_YDTool_GetHostIpAddress"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName)
    Local $sReturn = ""
    If _YDTool_IsPing($_sHost) Then
        Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sHost & "\root\cimv2")
        If Not @error Then
            Local $colItems = $objWMIService.ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True", "WQL", 0x30)
            If IsObj($colItems) Then
                For $objItem In $colItems
                    If $objItem.IPAddress(0) <> "" Then
                        $sReturn = $objItem.IPAddress(0)
                    EndIf
                Next
            Endif
        EndIf
    EndIf
    _YDLogger_Var("$sReturn", $sReturn, $sFuncName)
    Return $sReturn
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _YDTool_GetHostIpSubnet
; Description ...: Recupere l'adresse IP de sous-reseau du Host donne
; Syntax.........: _YDTool_GetHostIpSubnet($_sHost)
; Parameters ....: $_sHost       - Nom du PMF ou IP
; Return values .: Success      - $sReturn
;                  Failure      - ""
; Author ........: yann.daniel@assurance-maladie.fr
; Modified.......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func _YDTool_GetHostIpSubnet($_sHost)
    Local $sFuncName = "_YDTool_GetHostIpSubnet"
    _YDLogger_Var("$_sHost", $_sHost, $sFuncName)
    Local $sReturn = ""
    If _YDTool_IsPing($_sHost) Then
        Local $objWMIService = ObjGet("winmgmts:{impersonationLevel = impersonate}!\\" & $_sHost & "\root\cimv2")
        If Not @error Then
            Local $colItems = $objWMIService.ExecQuery("SELECT IPSubnet FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True", "WQL", 0x30)
            If IsObj($colItems) Then
                For $objItem In $colItems
                    If $objItem.IPSubnet(0) <> "" Then
                        $sReturn = $objItem.IPSubnet(0)
                    EndIf
                Next
            Endif
        EndIf
    EndIf
    _YDLogger_Var("$sReturn", $sReturn, $sFuncName)
    Return $sReturn
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
    GUICtrlCreateLabel(_YDGVars_Get("sAppName"), 0, 40, $iAboutWidth, -1, BitOr($SS_CENTER,$BS_CENTER))
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







