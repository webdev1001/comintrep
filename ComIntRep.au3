#NoTrayIcon
#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Version=Beta
#AutoIt3Wrapper_Icon=Resources\ComIntRep.ico
#AutoIt3Wrapper_Outfile=ComIntRep.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Fileversion=2.1.0.2104
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2014 Rizonesoft
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Res_Icon_Add=Resources\ResIP.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\Wins.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\RenICon.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\FlDNS.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\InEx.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\UpHist.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\WinUp.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\RepSSL.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\ResFir.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\hosts.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\RepWG.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\RunFixB.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\RunFix.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\Complete.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\Gear.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\Facebook.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\Twitter.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\LinkedIn.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\Google.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\PP.ico
#AutoIt3Wrapper_Res_Icon_Add=Resources\Info.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


Opt("TrayMenuMode", 1) ;~ Default tray menu items (Script Paused/Exit) will not be shown.
Opt("MustDeclareVars", 1)
Opt("GUIOnEventMode", 1)

#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <ButtonConstants.au3>
#include <GuiConstantsEx.au3>
#include <Constants.au3>
#include <GuiEdit.au3>
#include <Misc.au3>


#include <UDF\Configuration.au3>
#include <UDF\Functions.au3>
#include <UDF\Services.au3>


;~ Application Settings
Global Const $APPSET_TITLE = "Complete Internet Repair"

Global Const $RESOURCE_PROCESS = @ScriptDir & "\themes\101.ani"
Global Const $LOGGING_COMINTREP = @ScriptDir & "\logging\ComIntRepair.log"
Global Const $CIRResLSP = @ScriptDir & "\logging\CIRLSPs.log"
Global Const $CIRResLog = @ScriptDir & "\logging\CIRReset.log"
Global Const $Note2EXE = @ScriptDir & "\Bin\Notepad2\Notepad2.exe"
Global Const $GDataStoreDir = @WindowsDir & "\SoftwareDistribution\DataStore"
Global Const $GDownLoadDir = @WindowsDir & "\SoftwareDistribution\Download"
Global Const $GCATRootDir = @SystemDir & "\CatRoot2"
Global Const $CIRVer = FileGetVersion(@ScriptFullPath)

Global $LOGGING_MAXSIZE = 2048, $LOGGING_ENABLE = 1
Global $GOHPage
Global $IEXPLORE_VERSION

Global $mForm, $AppIcon, $PrAni, $lblWelc, $BtnGo, $eStatus
Global $RepCount = 10, $FileMenu, $CommMenu, $OpMenu, $HelpMenu, $HlpAbout, $AboutDlg
Global $ChkRep[$RepCount + 1], $IcoRep[$RepCount + 1], $BtnRep[$RepCount + 1], $HoverIcon[$RepCount + 1]

Global $ClearWinUpHist = False
Global $EventLogConfigured = False
Global $ResetWinsock = False
Global $Cancel = 0
Global $GoMode = 0



If @OSVersion = "WIN_2000" Or @OSVersion = "WIN_XPe" Then
	MsgBox(64, $APPSET_TITLE, 	$APPSET_TITLE & " is not compatable with your version of windows. If you believe this to be an error, " & _
								"please feel free to visit https://www.rizonesoft.com and send me a message.", 30)
	ShellExecute("https://www.rizonesoft.com")
	Exit
Else

	If Not @AutoItX64 And @OSArch = "X64" Then

		If FileExists(@ScriptDir & "\ComIntRep_x64.exe") Then
			ShellExecute(@ScriptDir & "\ComIntRep_x64.exe")
			Exit
		Else

			If Not IsDeclared("iMsgBox") Then Local $iMsgBox
			$iMsgBox = MsgBox(	$MB_YESNO + $MB_ICONEXCLAMATION + 262144, "Warning", _
								$APPSET_TITLE & " 32 Bit is not compatible with your Windows version. " & _
								"Please download " & $APPSET_TITLE & " 64 Bit. Would you like to visit the Download page " & _
								"now to download the 64 Bit version?", 60)
			Switch $iMsgBox
				Case  $IDYES
					ShellExecute("http://www.rizonesoft.com")
					Exit
				Case -1, $IDNO
					Exit
			EndSwitch

		EndIf

	Else

		If _Singleton(@ScriptName, 1) = 0 Then
			MsgBox(262192, "Warning!", "An occurence of " & $APPSET_TITLE & " is already running.", 30)
			Exit
		Else

			If Not FileExists(@ScriptDir & "\logging") Then DirCreate(@ScriptDir & "\logging")
			If Not FileExists(@ScriptDir & "\Themes") Then DirCreate(@ScriptDir & "\Themes")
			FileInstall("Resources\101.ani", $RESOURCE_PROCESS, 0)

			_LoadSettings()

			If FileExists($LOGGING_COMINTREP) Then
				FileSetAttrib($LOGGING_COMINTREP, "-RASHOT")
				If FileGetSize($LOGGING_COMINTREP) > ($LOGGING_MAXSIZE * 1024) Then
					FileDelete($LOGGING_COMINTREP)
				EndIf
			EndIf

			_LogWrite("", False)
			_LogWrite("", False)
			_LogWrite("                                            ./", False)
			_LogWrite("                                          (o o)", False)
			_LogWrite("--------------------------------------oOOo-(_)-oOOo--------------------------------------", False)

			If FileExists(@ProgramFilesDir & "\Internet Explorer\iexplore.exe") Then
				Local $sSpltString =  StringSplit(FileGetVersion(@ProgramFilesDir & "\Internet Explorer\iexplore.exe"), ".")
				$IEXPLORE_VERSION = StringStripWS($sSpltString[1] & "." & $sSpltString[2] & "." & $sSpltString[3], 8)
			EndIf

			_Main()

		EndIf

	EndIf

EndIf

Func _LoadSettings()

	$LOGGING_MAXSIZE = IniRead(@ScriptDir & "\CIntRep.ini", "Logging", "LogMaxSize", 2048)
	$LOGGING_ENABLE = IniRead(@ScriptDir & "\CIntRep.ini", "Logging", "LogEnabled", 1)

EndFunc


Func _Main()

	Local $FiEventView, $FiLoggMenu, $FiOLogDir, $FiOpenLog, $FiTcpResLog, $FiReboot, $FiClose
	Local $CommSysRes, $CommNetDiagWeb, $CommIPAll, $ComShowLSP, $ComInsIP6, $ComUnInsIP6, $ComRDP, $ComIEProperties
	Local $OpPre, $HlpHome, $HlpSpeedTest, $HlPasswords, $lblFreeMsg, $lnkPurch
	Local $lblNetDiagWeb, $lblSysRestore

	Local $BtnOptions, $BtnLSP

	$mForm = GUICreate($APPSET_TITLE & " " & _GetExecVersioning(@ScriptFullPath, 5), 420, 580, -1, -1)
	GUISetFont(8.5, 400, 0, "Verdana")
	$AppIcon = GUICtrlCreateIcon(@ScriptFullPath, 99, 5, 5, 72, 72)
	$PrAni = GUICtrlCreateIcon($RESOURCE_PROCESS, -1, 10, 10, 64, 64)
	GuiCtrlSetState($PrAni, $GUI_HIDE)
	$lblWelc = GUICtrlCreateLabel(	"Hold your mouse over an option's icon to view its description. " & _
									"Select your repair options and press 'Go!' to start. " & _
									"Do not select something unless your computer has the described problem. " & _
									"Skip any option you do not understand.", 90, 10, 310, 70)
	GuiCtrlSetColor($lblWelc, 0x555555)
	Switch @OSVersion
		Case "WIN_7", "WIN_8", "WIN_81", "WIN_2008", "WIN_2008R2", "WIN_2012", "WIN_2012R2"
			$lblNetDiagWeb = GUICtrlCreateLabel("Start Microsoft Internet Connection Troubleshooter.", 20, 90, 400, 20)
			;~ GuiCtrlSetFont($lblNetDiagWeb, 9, -1, 4) ;Underlined
			GuiCtrlSetColor($lblNetDiagWeb, 0x295496)
			GuiCtrlSetCursor($lblNetDiagWeb, 0)
		Case "WIN_XP", "WIN_XPe", "WIN_VISTA", "WIN_2003"
			$lblSysRestore = GUICtrlCreateLabel("Click here to create a System Restore Point.", 20, 90, 400, 20)
			;~ GuiCtrlSetFont($lblSysRestore, 9, -1, 4) ;Underlined
			GuiCtrlSetColor($lblSysRestore, 0x295496)
			GuiCtrlSetCursor($lblSysRestore, 0)
	EndSwitch

	$FileMenu = GUICtrlCreateMenu("&File")
	$FiEventView = GuiCtrlCreateMenuItem("&Event Viewer...", $FileMenu)
	GuiCtrlCreateMenuItem("", $FileMenu)
	$FiLoggMenu = GUICtrlCreateMenu("&Logging", $FileMenu)
	$FiOLogDir = GuiCtrlCreateMenuItem("Open &logging Directory...", $FiLoggMenu)
	GUICtrlCreateMenuItem("", $FiLoggMenu)
	$FiOpenLog = GuiCtrlCreateMenuItem("&Open [ComIntRepair.log]...", $FiLoggMenu)
	If @OSVersion = "WIN_XP" Or @OSVersion = "WIN_2003" Then
		GUICtrlCreateMenuItem("", $FiLoggMenu)
		$FiTcpResLog = GUICtrlCreateMenuItem("Open [CIRReset.log]...", $FiLoggMenu)
	EndIf
	GuiCtrlCreateMenuItem("", $FileMenu)
	$FiReboot = GuiCtrlCreateMenuItem("&Reboot Windows", $FileMenu)
	$FiClose = GuiCtrlCreateMenuItem("&Close " & @TAB & " Esc", $FileMenu)
	$CommMenu = GUICtrlCreateMenu("&Commands")

	$CommSysRes = GUICtrlCreateMenuItem("Create a System Restore Point", $CommMenu)
	GUICtrlCreateMenuItem("", $CommMenu)
	Switch @OSVersion
		Case "WIN_7", "WIN_8", "WIN_81", "WIN_2008", "WIN_2008R2", "WIN_2012", "WIN_2012R2"
			$CommNetDiagWeb = GUICtrlCreateMenuItem("Start Microsoft internet connection troubleshooter", $CommMenu)
			GUICtrlCreateMenuItem("", $CommMenu)
	EndSwitch
	$CommIPAll = GUICtrlCreateMenuItem("Show &TCP/IP configuration", $CommMenu)
	$ComShowLSP = GUICtrlCreateMenuItem("Show Winsock &LSPs", $CommMenu)
	If @OSVersion = "WIN_XP" Or @OSVersion = "WIN_2003" Then
		GUICtrlCreateMenuItem("", $CommMenu)
		$ComInsIP6 = GUICtrlCreateMenuItem("&Install IP6 protocol", $CommMenu)
		$ComUnInsIP6 = GUICtrlCreateMenuItem("&Uninstall IP6 protocol", $CommMenu)
	EndIf
	GUICtrlCreateMenuItem("", $CommMenu)
	$ComRDP = GUICtrlCreateMenuItem("Open Remote Desktop (RDP)", $CommMenu)
	$ComIEProperties = GUICtrlCreateMenuItem("Open Internet Explorer properties", $CommMenu)

	$OpMenu = GUICtrlCreateMenu("&Options")
	$OpPre = GuiCtrlCreateMenuItem("&Preferences...", $OpMenu)

	$HelpMenu = GUICtrlCreateMenu("&Help")
	$HlpHome = GUICtrlCreateMenuItem("&Rizonesoft Home", $HelpMenu)
	GuiCtrlCreateMenuItem("", $HelpMenu)
	$HlpSpeedTest = GUICtrlCreateMenuItem("&Internet Speed Test", $HelpMenu)
	$HlPasswords = GUICtrlCreateMenuItem("Get &Router Passwords", $HelpMenu)
	GuiCtrlCreateMenuItem("", $HelpMenu)
	$HlpAbout = GUICtrlCreateMenuItem("&About...", $HelpMenu)

	GUICtrlSetOnEvent($FiEventView, "_OpenEventViewer")
	GUICtrlSetOnEvent($FiOLogDir, "_OpenLoggingDirectory")
	GUICtrlSetOnEvent($FiOpenLog, "_OpenCIRLog")
	GUICtrlSetOnEvent($FiReboot, "_Reboot")
	GUICtrlSetOnEvent($FiClose, "_CloseClicked")
	GUICtrlSetOnEvent($CommSysRes, "_OpenWindowsSystemRestore")
	GUICtrlSetOnEvent($CommNetDiagWeb, "_OpenNetworkDiagnosticsWeb")
	GUICtrlSetOnEvent($CommIPAll, "_GetTCPIPFullConfig")
	GUICtrlSetOnEvent($ComShowLSP, "_ShowWinsockLSPs")
	GUICtrlSetOnEvent($ComInsIP6, "_InstallIP6")
	GUICtrlSetOnEvent($ComUnInsIP6, "_UnInstallIP6")
	GUICtrlSetOnEvent($ComRDP, "_OpenRDP")
	GUICtrlSetOnEvent($ComIEProperties, "_OpenIEProperties")

	GUICtrlSetOnEvent($OpPre, "_Options")

	GUICtrlSetOnEvent($HlpHome, "_HomePageClicked")
	GUICtrlSetOnEvent($HlpSpeedTest, "_SpeedTest")
	GUICtrlSetOnEvent($HlPasswords, "_GetRouterPasswords")
	GUICtrlSetOnEvent($HlpAbout, "_AboutDlg")

	GUICtrlSetOnEvent($lblSysRestore, "_OpenWindowsSystemRestore")
	GUICtrlSetOnEvent($lblNetDiagWeb, "_OpenNetworkDiagnosticsWeb")

	Local Const $Gap = 20
	GUICtrlCreateGroup("", 10, 115, 400, 325)
	For $i = 0 To $RepCount
		$IcoRep[$i] = GUICtrlCreateIcon(@ScriptFullPath, 201 + $i, 20, 135 + ($i * $Gap), 16, 16)
		GUICtrlSetCursor($IcoRep[$i], 0)
		$ChkRep[$i] = GUICtrlCreateCheckbox("", 50, 135 + ($i * $Gap), 280, 16)
		$HoverIcon[$i] = 1
		$BtnRep[$i] = GUICtrlCreateIcon(@ScriptFullPath, 212, 370, 135 + ($i * $Gap), 16, 16)
		GUICtrlSetCursor($BtnRep[$i], 0)
		GuiCtrlSetOnEvent($BtnRep[$i], "_RunRepair")
	Next
	GUICtrlCreateGroup("", -99, -99, 1, 1)  ;close group

	GuiCtrlSetTip($IcoRep[0], 	"This option rewrites important registry keys that are used by " & @CRLF & _
								"the Internet Protocol (TCP/IP) stack. This has the same result " & @CRLF & _
								"as removing and reinstalling the protocol.")
	GuiCtrlSetTip($IcoRep[1], 	"This can be used to recover from Winsock corruption result in lost " & @CRLF & _
								"of network connectivity. This option should be used with care "  & @CRLF & _
								"becuase any pre-installed LSPs will need to be reinstalled.")
	GuiCtrlSetTip($IcoRep[2], 	"Release and renew all Interent (TCP/IP) connections.")
	GuiCtrlSetTip($IcoRep[3], 	"Flush DNS Resolver Cache, refresh all DHCP leases and " & @CRLF & _
								"re-register DNS names.")
	GuiCtrlSetTip($IcoRep[4], 	"Re-registers all the concerned dll and ocx files required for " & @CRLF & _
								"the smooth operation of Internet Explorer " & $IEXPLORE_VERSION & ".")
	GuiCtrlSetTip($IcoRep[5], 	"This option will clear the Windows Update History. "  & @CRLF & _
								"It will do this by emptying the " & @CRLF & _
								"[" & $GDataStoreDir & "]" & @CRLF & "[" & $GDownLoadDir & "] directories.")
	GuiCtrlSetTip($IcoRep[6], 	"This option will try and fix Windows Update / Automatic Updates. " & @CRLF & _
								"Try this when you are unable to download or install updates.")
	GuiCtrlSetTip($IcoRep[7], 	"If you are having trouble connecting to SSL / Secured websites " & @CRLF & _
								"(Ex. Banking) then this option could help.")
	GuiCtrlSetTip($IcoRep[8], 	"Reset the Windows Firewall configuration to its default state.")
	GuiCtrlSetTip($IcoRep[9], 	"Reset the Windows hosts file to its default state.")
	GuiCtrlSetTip($IcoRep[10], 	"Select this option if you cannot view other "  & @CRLF & _
								"workgroup computers on the network.")

	GuiCtrlSetData($ChkRep[0], " Reset Internet Protocol (TCP/IP)")
	GuiCtrlSetData($ChkRep[1], " Repair Winsock (Reset Catalog)")
	GuiCtrlSetData($ChkRep[2], " Renew Internet Connections")
	GuiCtrlSetData($ChkRep[3], " Flush DNS Resolver Cache")
	GuiCtrlSetData($ChkRep[4], " Repair Internet Explorer " & $IEXPLORE_VERSION)
	GuiCtrlSetData($ChkRep[5], " Clear Windows Update History")
	GuiCtrlSetData($ChkRep[6], " Repair Windows / Automatic Updates")
	GuiCtrlSetData($ChkRep[7], " Repair SSL / HTTPS / Cryptography")
	GuiCtrlSetData($ChkRep[8], " Reset Windows Firewall Configuration")
	GuiCtrlSetData($ChkRep[9], " Restore the default hosts file")
	GuiCtrlSetData($ChkRep[10], " Repair Workgroup Computers view")

	If @OSVersion = "WIN_VISTA" Or @OSVersion = "WIN_2008" Or @OSVersion = "WIN_2008R2" Or @OSVersion = "WIN_7" Then
		;GuiCtrlSetState($IcoRep[10], $GUI_DISABLE)
		;GuiCtrlSetState($ChkRep[10], $GUI_DISABLE)
		;GuiCtrlSetState($BtnRep[10], $GUI_DISABLE)
	EndIf

	$BtnGo = GUICtrlCreateButton("Go!", 200, 380, 190, 50)
	GuiCtrlSetState($BtnGo, $GUI_FOCUS)
	GUICtrlSetFont($BtnGo, 11, 400, 0, "Verdana")
	GuiCtrlSetOnEvent($BtnGo, "_Go")

	;$iTip = GuiCtrlCreateIcon(@ScriptFullPath, 212, 10, 210, 16,16)

	;GUICtrlSetFont($lblInfo, 8.5, 400, 0, "Verdana", $GFSm)
	;$BtnOptions = GUICtrlCreateButton("Options...", 205, 360, 100, 30)
	;GUICtrlSetFont($BtnOptions, 8.5, 400, 0, "Verdana", $GFSm)
	;$BtnLSP = GUICtrlCreateButton("LSPs...", 310, 360, 60, 30)
	;GUICtrlSetFont($BtnLSP, 8.5, 400, 0, "Verdana", $GFSm)

	$eStatus = GUICtrlCreateEdit("", 10, 453, 400, 90, BitOR($WS_VSCROLL, $WS_HSCROLL, $ES_READONLY, $ES_AUTOVSCROLL))
	GuiCtrlSetFont(-1, 9, 400, -1, "Courier New")

	;SoundPlay(@ScriptDir & "\Sounds\welcome.wav")

	GUISetState(@SW_SHOW, $mForm)

	GUIRegisterMsg($WM_COMMAND, "MY_WM_COMMAND")
	GUISetOnEvent($GUI_EVENT_CLOSE, "_CloseClicked")
	GUICtrlSetOnEvent($BtnLSP, "_ShowWinsockLSPs")

	While 1
		_GUIControl()
		_ReduceMemory()
		Sleep(35)
	WEnd

EndFunc


Func _OpenNetworkDiagnosticsWeb()
	Run("msdt.exe /id NetworkDiagnosticsWeb")
EndFunc


Func _OpenWindowsSystemRestore()
	Run("systempropertiesprotection")
EndFunc


Func _OpenRizonesoftDownloads()
	ShellExecute("http://www.rizonesoft.com/freeware-downloads/")
EndFunc


Func _GUIControl()
	Local $cursor = GUIGetCursorInfo()
	If Not @error Then
		For $a = 0 To $RepCount
			If $cursor[4] = $BtnRep[$a] And $HoverIcon[$a] = 1 Then
				$HoverIcon[$a] = 0
				GUICtrlSetImage($BtnRep[$a], @ScriptFullPath, 213)
			ElseIf $cursor[4] <> $BtnRep[$a] And $HoverIcon[$a] = 0 Then
				$HoverIcon[$a] = 1
				GUICtrlSetImage($BtnRep[$a], @ScriptFullPath, 212)
			EndIf
		Next
	EndIf
EndFunc


Func MY_WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)

    Switch BitAND($wParam, 0xFFFF) ;LoWord = IDFrom
        Case $BtnGo
            Switch BitShift($wParam, 16) ;HiWord = Code
				Case $BN_CLICKED
					If GUICtrlRead($BtnGo) = "Stop!" Then
						$Cancel = 1
					EndIf
            EndSwitch
    EndSwitch

    Return $GUI_RUNDEFMSG
EndFunc;==>WM_COMMAND


Func _CloseClicked()

	Local $PID = ProcessExists(@ScriptName) ; Will return the PID or 0 if the process isn't found.
	If $PID Then ProcessClose(@ScriptName)
	Exit

EndFunc

Func _Reboot()
	Shutdown(18)
EndFunc

Func _RunRepair()
	_StartProcess()
	Switch @GUI_CtrlId
		Case $BtnRep[0]
			_ResetTCPIP()
		Case $BtnRep[1]
			_RepairWinsock()
		Case $BtnRep[2]
			_ReleaseRenewIP()
		Case $BtnRep[3]
			_FlushReDNS()
		Case $BtnRep[4]
			_RepairIE()
		Case $BtnRep[5]
			_ClearUpdateHistory()
		Case $BtnRep[6]
			_RepairWUAU()
		Case $BtnRep[7]
			_RepairSHC()
		Case $BtnRep[8]
			_ResetFirewall()
		Case $BtnRep[9]
			_RestoreHosts()
		Case $BtnRep[10]
			_RepairWorkGroups()
	EndSwitch
	_EndProcess()
EndFunc

Func _InstallIP6()
	;_StartProcess()
	_MemoLogWrite("Installing the TCP/IP v6 protocol, Please wait.....")
	Local $eCode = ShellExecuteWait("netsh", "int ipv6 install", "", "", @SW_HIDE)
	If Not $eCode Then
		_MemoLogWrite("The TCP/IP v6 protocol Installation was successful or it's already installed.", 1)
	Else
		_MemoLogWrite("Could not install TCP/IP v6.", 2)
		_MemoLogWrite("You may need to restart your computer for the settings to take effect.")
	EndIf
	;_EndProcess()
EndFunc

Func _UnInstallIP6()
	;_StartProcess()
	_MemoLogWrite("Uninstalling the TCP/IP v6 protocol, Please wait.....")
	Local $eCode = ShellExecuteWait("netsh", "int ipv6 uninstall", "", "", @SW_HIDE)
	If Not $eCode Then
		_MemoLogWrite("The TCP/IP v6 protocol was successfully uninstalled or it's not installed.", 1)
	Else
		_MemoLogWrite("Could not uninstall TCP/IP v6.", 2)
		_MemoLogWrite("You may need to restart your computer for the settings to take effect.")
	EndIf
	;_EndProcess()
EndFunc

Func _GetTCPIPFullConfig()
	_MemoLogWrite("Getting full TCP/IP configuration, Please wait.....", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	RunWait(@ComSpec & " /k ipconfig /all", "", @SW_SHOW)
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)
EndFunc


Func _OpenRDP()
	_MemoLogWrite("Opening the Remote Desktop (RDP) tool, Please wait.....", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	ShellExecute("mstsc")
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)
EndFunc


Func _OpenIEProperties()
	_MemoLogWrite("Opening the Internet Explorer Properties dialog box, Please wait.....", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	ShellExecute("inetcpl.cpl")
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)
EndFunc


Func _Go()

	$GoMode = 1

;~ 	Local $cs, $cc = 0, $iText

	For $c = 0 To $RepCount
		If GUICtrlRead($ChkRep[$c]) = $GUI_CHECKED Then
			_StartProcess()
			ExitLoop
		EndIf
	Next

;~ 	If $cc > 0 Then
;~
;~ 	EndIf

	For $i = 0 To $RepCount

		If GUICtrlRead($ChkRep[$i]) = $GUI_CHECKED Then
			_StartProcess()
			If GUICtrlRead($ChkRep[$i], 1) = " Reset Internet Protocol (TCP/IP)" Then
					If Not $Cancel Then _ResetTCPIP()
					GUICtrlSetState($ChkRep[$i], $GUI_UNCHECKED)
					GUICtrlSetImage($IcoRep[$i], @ScriptFullPath, 214)
			ElseIf GUICtrlRead($ChkRep[$i], 1) = " Repair Winsock (Reset Catalog)" Then
					If Not $Cancel Then _RepairWinsock()
					GUICtrlSetState($ChkRep[$i], $GUI_UNCHECKED)
					GUICtrlSetImage($IcoRep[$i], @ScriptFullPath, 214)
			ElseIf GUICtrlRead($ChkRep[$i], 1) = " Renew Internet Connections" Then
					If Not $Cancel Then _ReleaseRenewIP()
					GUICtrlSetState($ChkRep[$i], $GUI_UNCHECKED)
					GUICtrlSetImage($IcoRep[$i], @ScriptFullPath, 214)
			ElseIf GUICtrlRead($ChkRep[$i], 1) = " Flush DNS Resolver Cache" Then
					If Not $Cancel Then _FlushReDNS()
					GUICtrlSetState($ChkRep[$i], $GUI_UNCHECKED)
					GUICtrlSetImage($IcoRep[$i], @ScriptFullPath, 214)
			ElseIf GUICtrlRead($ChkRep[$i], 1) = " Repair Internet Explorer " & $IEXPLORE_VERSION Then
					If Not $Cancel Then _RepairIE()
					GUICtrlSetState($ChkRep[$i], $GUI_UNCHECKED)
					GUICtrlSetImage($IcoRep[$i], @ScriptFullPath, 214)
			ElseIf GUICtrlRead($ChkRep[$i], 1) = " Clear Windows Update History" Then
					If Not $Cancel Then $ClearWinUpHist = True
					GUICtrlSetState($ChkRep[$i], $GUI_UNCHECKED)
					GUICtrlSetImage($IcoRep[$i], @ScriptFullPath, 214)
			ElseIf GUICtrlRead($ChkRep[$i], 1) = " Repair Windows / Automatic Updates" Then
					If GUICtrlRead($ChkRep[$i]) = $GUI_CHECKED Then
						GUICtrlSetState($ChkRep[$i + 1], $GUI_CHECKED)
					EndIf
					If Not $Cancel Then _RepairWUAU()
					GUICtrlSetState($ChkRep[$i], $GUI_UNCHECKED)
					GUICtrlSetImage($IcoRep[$i], @ScriptFullPath, 214)
			ElseIf GUICtrlRead($ChkRep[$i], 1) = " Repair SSL / HTTPS / Cryptography" Then
					If Not $Cancel Then _RepairSHC()
					GUICtrlSetState($ChkRep[$i], $GUI_UNCHECKED)
					GUICtrlSetImage($IcoRep[$i], @ScriptFullPath, 214)
			ElseIf GUICtrlRead($ChkRep[$i], 1) = " Reset Windows Firewall Configuration" Then
					If Not $Cancel Then _ResetFirewall()
					GUICtrlSetState($ChkRep[$i], $GUI_UNCHECKED)
					GUICtrlSetImage($IcoRep[$i], @ScriptFullPath, 214)
			ElseIf GUICtrlRead($ChkRep[$i], 1) = " Restore the default hosts file" Then
					If Not $Cancel Then _RestoreHosts()
					GUICtrlSetState($ChkRep[$i], $GUI_UNCHECKED)
					GUICtrlSetImage($IcoRep[$i], @ScriptFullPath, 214)
			ElseIf GUICtrlRead($ChkRep[$i], 1) = " Repair Workgroup Computers view" Then
					If Not $Cancel Then _RepairWorkGroups()
					GUICtrlSetState($ChkRep[$i], $GUI_UNCHECKED)
					GUICtrlSetImage($IcoRep[$i], @ScriptFullPath, 214)
			EndIf
		EndIf
	Next
	_EndProcess()
	_BootMessage()
	$EventLogConfigured = False
	$ResetWinsock = False
	$Cancel = 0

EndFunc


Func _ResetTCPIP()

	_MemoLogWrite("Resetting all TCP/IP Interfaces, Please wait.....", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	Switch @OSVersion

		Case "WIN_XP", "WIN_2003"

			Local $eCode = ShellExecuteWait("netsh", "interface ip reset """ & $CIRResLog & """", "", "", @SW_HIDE)
			If Not $eCode Then
				_MemoLogWrite("TCP/IP Stack reset successful.", 1)
				_MemoLogWrite("TCP/IP Reset log located @ [" & $CIRResLog & "]", 1)
			EndIf
			Local $eCodeEx = ShellExecuteWait("netsh", "interface reset all", "", "", @SW_HIDE)
			If Not $eCodeEx Then
				_MemoLogWrite("TCP/IP interfaces reset successful.", 1)
			EndIf
			If $eCode <> 0 Or $eCodeEx <> 0 Then
				_MemoLogWrite("TCP/IP interfaces reset failed or no user specific settings found.", 2)
			EndIf
			$eCode = ShellExecuteWait("netsh", "interface ipv6 reset all", "", "", @SW_HIDE)
			If Not $eCode Then
				_MemoLogWrite("TCP/IP v6 interfaces reset successful.", 1)
			Else
				_MemoLogWrite("The TCP/IP v6 protocol might not be installed.", 3)
				_MemoLogWrite("Click on 'Commands' then 'Install IP6 protocol' to install TCP/IP v6.")
			EndIf

		Case "WIN_VISTA", "WIN_2008", "WIN_2008R2", "WIN_7", "WIN_8", "WIN_81"

			ShellExecuteWait("netsh", "interface ipv4 reset all", "", "", @SW_HIDE)
			_MemoLogWrite("TCP/IP interfaces reset successful.", 1)
			ShellExecuteWait("netsh", "interface ipv6 reset all", "", "", @SW_HIDE)
			_MemoLogWrite("TCP/IP v6 interfaces reset successful.", 1)

	EndSwitch
	_MemoLogWrite("You may need to restart your computer for the settings to take effect.")
	_MemoLogWrite("Finished resetting the Internet Protocol (TCP/IP).")
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)

EndFunc



Func _RepairWinsock()

	$ResetWinsock = True
	_MemoLogWrite("Attempting to reset Winsock catalog, Please wait.....")
	_LogWrite("-----------------------------------------------------------------------------------------", False)

	Switch @OSVersion
		Case "WIN_XP"
			Switch @OSServicePack
				Case "Service Pack 1"
					_MemoLogWrite("It is recommended that you install Windows XP Service Pack 2 or later.", 3)
					RegDelete("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\Winsock")
					RegDelete("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\Winsock2")
				Case Else
					_ResetWinsock()
			EndSwitch
		Case "WIN_2003", "WIN_VISTA", "WIN_2008", "WIN_2008R2", "WIN_7", "WIN_8", "WIN_81"
			_ResetWinsock()
	EndSwitch
	_MemoLogWrite("Finished repairing Winsock")
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)

EndFunc


Func _ResetWinsock()

	Local $eCode = ShellExecuteWait("netsh", "winsock reset catalog", "", "", @SW_HIDE)
	Local $eCodeEx = ShellExecuteWait("netsh", "winsock reset", "", "", @SW_HIDE)
	If $eCode <> 0 Or $eCodeEx <> 0 Then
		_MemoLogWrite("Could not reset the Winsock Catalog.", 2)
	Else
		_MemoLogWrite("Successfully reset the Winsock Catalog.", 1)
	EndIf

EndFunc


Func _ReleaseRenewIP()

	_MemoLogWrite("Releasing TCP/IP connections, Please wait.....")
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	Local $eCode = ShellExecuteWait("ipconfig", "/release", "", "", @SW_HIDE)
	If Not $eCode Then
		_MemoLogWrite("Successfully released TCP/IP connections.", 1)
	Else
		_MemoLogWrite("For some reason the TCP/IP connections could not be released.", 2)
	EndIf
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	_MemoLogWrite("Renewing TCP/IP connections, Please wait.....")
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	Local $eCode = ShellExecuteWait("ipconfig", "/renew", "", "", @SW_HIDE)
	If Not $eCode Then
		_MemoLogWrite("Successfully renewed TCP/IP adapters.", 1)
	Else
		_MemoLogWrite("For some reason the TCP/IP connections could not be renewed.", 2)
	EndIf
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)

EndFunc


Func _FlushReDNS()

	If Not $EventLogConfigured Then _ConfigureEventLog()
	_MemoLogWrite("Flushing DNS Resolver Cache, Please wait.....")
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	ShellExecuteWait("ipconfig", "/flushdns", "", "", @SW_HIDE)
	_MemoLogWrite("Successfully flushed DNS Resolver Cache.", 1)
	_MemoLogWrite("Refreshing all DHCP leases and re-registering DNS names, Please wait.....")
	ShellExecuteWait("ipconfig", "/registerdns", "", "", @SW_HIDE)
	_MemoLogWrite("Registration of the DNS resource records has been initiated.", 1)
	_MemoLogWrite("Note: Any errors will be reported in the 'Event Viewer' in about 15 minutes.")
	_MemoLogWrite("Note: Click on 'File' and then 'Event Viewer...' to open the Event Viewer.", 3)
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)

EndFunc


Func _RepairIE()

	_MemoLogWrite("Repairing Internet Explorer " & $IEXPLORE_VERSION & ", Please wait.....")
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	If ProcessExists("iexplore.exe") Then
		MsgBox(48, "Internet Explorer", "Closing Internet Explorer.  Save your work before you press OK")
		ProcessClose("iexplore.exe")
	EndIf

	_ReregisterDLL(Chr(34) & @ProgramFilesDir & "\Internet Explorer\DiagnosticsHub_is.dll" & Chr(34))
	_ReregisterDLL(Chr(34) & @ProgramFilesDir & "\Internet Explorer\DiagnosticsTap.dll" & Chr(34))
	_ReregisterDLL(Chr(34) & @ProgramFilesDir & "\Internet Explorer\F12.dll" & Chr(34))
	_ReregisterDLL(Chr(34) & @ProgramFilesDir & "\Internet Explorer\F12Tools.dll" & Chr(34))
	_ReregisterDLL(Chr(34) & @ProgramFilesDir & "\Internet Explorer\hmmapi.dll" & Chr(34))
	_ReregisterDLL(Chr(34) & @ProgramFilesDir & "\Internet Explorer\iedvtool.dll" & Chr(34))
	_ReregisterDLL(Chr(34) & @ProgramFilesDir & "\Internet Explorer\ieproxy.dll" & Chr(34))
	_ReregisterDLL(Chr(34) & @ProgramFilesDir & "\Internet Explorer\msdbg2.dll" & Chr(34))
	_ReregisterDLL(Chr(34) & @ProgramFilesDir & "\Internet Explorer\pdm.dll" & Chr(34))
	_ReregisterDLL(Chr(34) & @ProgramFilesDir & "\Internet Explorer\pdmproxy100.dll" & Chr(34))
	_ReregisterDLL(Chr(34) & @ProgramFilesDir & "\Internet Explorer\perf_nt.dll" & Chr(34))
	_ReregisterDLL(Chr(34) & @ProgramFilesDir & "\Internet Explorer\perfcore.dll" & Chr(34))
	_ReregisterDLL(Chr(34) & @ProgramFilesDir & "\Internet Explorer\Timeline_is.dll" & Chr(34))

	;~ Symptom: open in new tab/window not working
	_ReregisterDLL("actxprxy.dll")
	_ReregisterDLL("asctrls.ocx")
	_ReregisterDLL("browseui.dll", "/s /i")
	;~ regsvr32 /s /i browseui.dll,NI (unnecessary)
	_ReregisterDLL("cdfview.dll")
	_ReregisterDLL("comcat.dll")
	_ReregisterDLL("comctl32.dll", "/s /i /n")
	_ReregisterDLL("corpol.dll")
	_ReregisterDLL("cryptdlg.dll")
	_ReregisterDLL("""" & @ProgramFilesDir & "\Internet Explorer\custsat.dll""")

	_ReregisterDLL("digest.dll", "/s /i /n")
	_ReregisterDLL("dispex.dll")
	_ReregisterDLL("dxtmsft.dll")
	_ReregisterDLL("dxtrans.dll")
	;~ Symptom: Add-Ons-Manager menu entry is present but nothing happens
	_ReregisterDLL("extmgr.dll")
	;~ Simple HTML Mail API
	_ReregisterDLL("""" & @ProgramFilesDir & "\Internet Explorer\hmmapi.dll""")
	_ReregisterDLL("hlink.dll")
	;~ Group policy snap-in
	_ReregisterDLL("ieaksie.dll")
	;~ Smart Screen
	_ReregisterDLL("ieapfltr.dll")
	;~ IEAK Branding
	_ReregisterDLL("iedkcs32.dll")
	;~ Dev Tools
	_ReregisterDLL("""" & @ProgramFilesDir & "\Internet Explorer\iedvtool.dll""")
	_ReregisterDLL("iedvtool.dll")
	;~ IE7 tabbed browser
	_ReregisterDLL("ieframe.dll", "/s /i /n")
	;~ _ReregisterDLL("ieframe.dll", "/s /i")
	_ReregisterDLL("iepeers.dll")
	;~ Symptom: IE8 closes immediately on launch, missing from IE7
	_ReregisterDLL("""" & @ProgramFilesDir & "\Internet Explorer\ieproxy.dll""")
	_ReregisterDLL("ieproxy.dll")
	;~ iesetup.dll has DllINstall for WinXP,NT4Only,NTx86
	_ReregisterDLL("iesetup.dll", "/s /i")
	_ReregisterDLL("imgutil.dll")
	_ReregisterDLL("inetcpl.cpl", "/s /i")
	_ReregisterDLL("inetcpl.cpl", "/s /i /n")
	_ReregisterDLL("initpki.dll", "/s /i:A")
	_ReregisterDLL("inseng.dll", "/s /i")
	_ReregisterDLL("jscript.dll")
	;~ License Manager
	_ReregisterDLL("licmgr10.dll")
	_ReregisterDLL("mlang.dll")
	_ReregisterDLL("mobsync.dll")
	_ReregisterDLL("msapsspc.dll")
	;~ Symptom: Javascript links don't work (Robin Walker) .NET hub file
	_ReregisterDLL("mscoree.dll")
	_ReregisterDLL("mscorier.dll")
	_ReregisterDLL("mscories.dll")
	;~ VS Debugger
	_ReregisterDLL("msdbg2.dll")
	_ReregisterDLL("mshta.exe")
	_ReregisterDLL("mshtml.dll", "/s /i")
	_ReregisterDLL("mshtmled.dll")
	_ReregisterDLL("msident.dll")
	_ReregisterDLL("msieftp.dll", "/s /i")
	_ReregisterDLL("msnsspc.dll")
	_ReregisterDLL("msr2c.dll")
	_ReregisterDLL("msrating.dll")
	_ReregisterDLL("mstime.dll")
	_ReregisterDLL("msxml.dll")
	;~ Symptom: Printing problems, open in new window
	_ReregisterDLL("ole32.dll")
	;~ Symptom: Find on this page is blank
	_ReregisterDLL("oleacc.dll")
	_ReregisterDLL("occache.dll", "/s /i")
	_ReregisterDLL("oleaut32.dll")
	;~ Process debug manager
	_ReregisterDLL("""" & @ProgramFilesDir & "\Internet Explorer\pdm.dll""")
	_ReregisterDLL("plugin.ocx")
	_ReregisterDLL("pngfilt.dll")
	_ReregisterDLL("proctexe.ocx")
	_ReregisterDLL("scrobj.dll", "/s /i")
	_ReregisterDLL("sendmail.dll")
	_ReregisterDLL("setupwbv.dll", "/s /i")
	_ReregisterDLL("shdocvw.dll", "/s /i")
	;~ regsvr32 /s /i shdocvw.dll,NI
	_ReregisterDLL("tdc.ocx")
	_ReregisterDLL("url.dll")
	_ReregisterDLL("urlmon.dll", "/s /i")
	;~ regsvr32 /s /i urlmon.dll,NI,HKLM
	_ReregisterDLL("urlmon.dll,NI,HKLM", "/s /i")
	_ReregisterDLL("vbscript.dll")
	;~ VML Renderer
	_ReregisterDLL("""" & @ProgramFilesDir & "\microsoft shared\vgx\vgx.dll""")
	_ReregisterDLL("webcheck.dll", "/s /i")
	;_ReregisterDLL("wininet.dll", "/s /i /n")
	If @OSVersion = "WIN_XP" Or @OSVersion = "WIN_2003" Then
		;~ Symptom: new tabs page cannot display content because it cannot access the controls (added 27. 3.2009)
		;~ This is a result of a bug in shdocvw.dll (see above), probably only on Windows XP
		_MemoLogWrite("Fixing 'New tabs page cannot display content because it cannot access the controls'.")
		_MemoLogWrite("This is a result of a bug in shdocvw.dll.")
		Local $RegReturn = RegWrite("HKEY_CLASSES_ROOT\TypeLib\{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}\1.1\0\win32", "", "REG_SZ", "%SystemRoot%\system32\ieframe.dll")
		If $RegReturn Then
		Else
			Switch @error
				Case 1
					_MemoLogWrite("Unable to open requested registry key.", 2)
				Case 2
					_MemoLogWrite("Unable to open requested main registry key.", 2)
				Case 3
					_MemoLogWrite("Unable to remote connect to the registry.", 2)
				Case -1
					_MemoLogWrite("Unable to open requested registry value.", 2)
				Case -2
					_MemoLogWrite("Registry value type not supported.", 2)
			EndSwitch
		EndIf
		_MemoLogWrite("Registering Outlook Express files.....")
		_ReregisterDLL("""" & @ProgramFilesDir & "\Outlook Express\msoe.dll""")
		_ReregisterDLL("""" & @ProgramFilesDir & "\Outlook Express\oeimport.dll""")
		_ReregisterDLL("""" & @ProgramFilesDir & "\Outlook Express\oemiglib.dll""")
		_ReregisterDLL("""" & @ProgramFilesDir & "\Outlook Express\wabfind.dll""")
		_ReregisterDLL("""" & @ProgramFilesDir & "\Outlook Express\wabimp.dll""")
		;~ _MemoLogWrite("Registering Connection Wizard files.....")
		;~ _ReregisterDLL("""" & @ProgramFilesDir & "\Internet Explorer\Connection Wizard\icwconn.dll""")
		;~ _ReregisterDLL("""" & @ProgramFilesDir & "\Internet Explorer\Connection Wizard\icwdl.dll""")
		;~ _ReregisterDLL("""" & @ProgramFilesDir & "\Internet Explorer\Connection Wizard\icwutil.dll""")
		;~ _ReregisterDLL("""" & @ProgramFilesDir & "\Internet Explorer\Connection Wizard\trialoc.dll""")
	EndIf
	_MemoLogWrite("Finished repairing Internet Explorer " & $IEXPLORE_VERSION)
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)

EndFunc


Func _RepairWUAU()

	If Not $EventLogConfigured Then _ConfigureEventLog()
	_MemoLogWrite("Repairing Windows Update / Automatic Updates, Please wait.....")
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	_MemoLogWrite("Stopping the BITS Service.....")
	If Not _SvcStop("bits") Then
		_MemoLogWrite("BITS was not started in the first place.", 3)
	Else
		_MemoLogWrite("BITS Stopped Successfully.", 1)
	EndIf
	_MemoLogWrite("Stopping the Automatic Updates (wuauserv) Service.....")
	If Not _SvcStop("wuauserv") Then
		_MemoLogWrite("Automatic Updates (wuauserv) Service was not started in the first place.", 3)
	Else
		_MemoLogWrite("Automatic Updates (wuauserv) Service Stopped Successfully.", 1)
	EndIf
	If $ClearWinUpHist Then _ClearUpdateHistory()
	_MemoLogWrite("Setting BITS Security Descriptor.....")
	ShellExecuteWait("sc", 'sdset bits "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)"', "", "", @SW_HIDE)
	_MemoLogWrite("BITS Security Descriptor Set.", 1)
	_MemoLogWrite("Setting Automatic Updates (wuauserv) Service Security Descriptor.....")
	ShellExecuteWait("sc", 'sdset wuauserv "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)"', "", "", @SW_HIDE)
	_MemoLogWrite("Automatic Updates (wuauserv) Security Descriptor Set.", 1)
	_MemoLogWrite("Configuring the Automatic Updates (wuauserv) Service.....")
	_SvcSetStartMode("wuauserv","Automatic")
	_MemoLogWrite("Automatic Updates (wuauserv) Service Configured.", 1)
	_MemoLogWrite("Configuring BITS.....")
	_SvcSetStartMode("bits","Automatic")
	_MemoLogWrite("BITS Configured.", 1)
	_MemoLogWrite("Registering WUAU DLLs.....")
	_ReregisterDLL("actxprxy.dll")
	_ReregisterDLL("atl.dll")
	_ReregisterDLL("browseui.dll")
	_ReregisterDLL("corpol.dll")
	_ReregisterDLL("cryptdlg.dll")
	_ReregisterDLL("dispex.dll")
	_ReregisterDLL("dssenh.dll")
	_ReregisterDLL("gpkcsp.dll")
	_ReregisterDLL("initpki.dll")
	_ReregisterDLL("jscript.dll")
	_ReregisterDLL("mshtml.dll")
	_ReregisterDLL("msscript.ocx")
	_ReregisterDLL("msxml.dll")
	_ReregisterDLL("msxml2.dll")
	_ReregisterDLL("msxml3.dll")
	_ReregisterDLL("msxml4.dll")
	_ReregisterDLL("msxml6.dll")
	_ReregisterDLL("muweb.dll")
	_ReregisterDLL("ole.dll")
	_ReregisterDLL("ole32.dll")
	_ReregisterDLL("oleaut.dll")
	_ReregisterDLL("oleaut32.dll")
	_ReregisterDLL("qmgr.dll")
	_ReregisterDLL("qmgrprxy.dll")
	_ReregisterDLL("gpkcsp.dll")
	_ReregisterDLL("rsaenh.dll")
	_ReregisterDLL("sccbase.dll")
	_ReregisterDLL("scrobj.dll")
	_ReregisterDLL("scrrun.dll")
	_ReregisterDLL("shdocvw.dll")
	_ReregisterDLL("shell.dll")
	_ReregisterDLL("shell32.dll")
	_ReregisterDLL("slbcsp.dll")
	_ReregisterDLL("softpub.dll")
	_ReregisterDLL("urlmon.dll")
	_ReregisterDLL("vbscript.dll")
	_ReregisterDLL("winhttp.dll")
	_ReregisterDLL("wintrust.dll")
	_ReregisterDLL("wshext.dll")
	_ReregisterDLL("wuapi.dll")
	_ReregisterDLL("wuaueng.dll")
	_ReregisterDLL("wuaueng1.dll")
	_ReregisterDLL("wucltui.dll")
	_ReregisterDLL("wucltux.dll")
	_ReregisterDLL("wups.dll")
	_ReregisterDLL("wups2.dll")
	_ReregisterDLL("wuweb.dll")
	_ReregisterDLL("wuwebv.dll")
	_MemoLogWrite("WUAU DLLs Reregistered.", 1)
	If Not $ResetWinsock Then _RepairWinsock()
	Switch @OSVersion
		Case "WIN_2000", "WIN_XP", "WIN_XPe", "WIN_2003"
			_MemoLogWrite("Setting proxy to direct access.....")
			ShellExecuteWait("proxycfg.exe", "-d", "", "", @SW_HIDE)
			_MemoLogWrite("Proxy set to direct access.", 1)
		Case "WIN_VISTA", "WIN_2008", "WIN_7", "WIN_2008R2"
			_MemoLogWrite("Resetting proxy settings.....")
			ShellExecuteWait("netsh", "winhttp reset proxy", "", "", @SW_HIDE)
			_MemoLogWrite("Proxy settings reset successfully.", 1)
	EndSwitch
	_MemoLogWrite("Restarting the Automatic Updates (wuauserv) Service.....")
	If Not _SvcStart("wuauserv") Then
		_MemoLogWrite("The wuauserv Service could not be started.", 2)
	Else
		_MemoLogWrite("Automatic Updates (wuauserv) Service Restarted.", 1)
	EndIf
	_MemoLogWrite("Restarting the BITS Service.....")
	If Not _SvcStart("bits") Then
		_MemoLogWrite("The BITS Service could not be started.", 2)
	Else
		_MemoLogWrite("BITS Service Restarted.", 1)
	EndIf
	ShellExecuteWait("fsutil","resource setautoreset true "&@HomeDrive&":\", @SystemDir,Default,@SW_HIDE)
	If @OSVersion = "WIN_VISTA" Or @OSVersion = "WIN_2008" Or @OSVersion = "WIN_2008R2" Or @OSVersion = "WIN_7" Then
		_MemoLogWrite("Clearing the BITS queue.....")
		ShellExecuteWait("bitsadmin.exe", "/reset /allusers", "", "", @SW_HIDE)
		_MemoLogWrite("BITS queue cleared.", 1)
	EndIf
	RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Group Policy Objects\LocalUser\Software\Microsoft\Windows\CurrentVersion\Policies\WindowsUpdate\DisableWindowsUpdateAccess")
	RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWindowsUpdate")
	RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDevMgrUpdate")
	RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WindowsUpdate", "DisableWindowsUpdateAccess")
	RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\WindowsUpdate")
	RegDelete("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWindowsUpdate")
	RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", "NoAutoUpdate")
	RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", "AUOptions")
	RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", "ScheduledInstallDay")
	RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", "ScheduledInstallTime")
	RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", "NoAutoRebootWithLoggedOnUsers")
	RegDelete("HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WindowsUpdate")
	RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update", "LastWaitTimeout")
	RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update", "DetectionStartTime")
	RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update", "NextDetectionTime")
	RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update", "ScheduledInstallDate")
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update", "NoAutoUpdate", "REG_DWORD", 0)
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update", "AUOptions", "REG_DWORD", 4)
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update", "ScheduledInstallDay", "REG_DWORD", 0)
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update", "ScheduledInstallTime", "REG_DWORD", 3)
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update", "NoAutoRebootWithLoggedOnUsers", "REG_DWORD", 1)
	RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main", "NoUpdateCheck", "REG_DWORD", 0)
	_MemoLogWrite("Initiating Windows Updates detection right away.....")
	RunWait("wuauclt /detectnow", @SystemDir)
	_MemoLogWrite("Finished repairing Windows Update / Automatic Updates.")
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)

EndFunc

Func _ClearUpdateHistory()

	_MemoLogWrite("Clearing File Stores (Update History).....")
	FileDelete(@AppDataCommonDir & "\Microsoft\Network\Downloader\qmgr*.dat")
	_MemoLogWrite("Clearing [" & $GDownLoadDir & "].....")
	If DirRemove($GDownLoadDir, 1) Then
		_MemoLogWrite("[" & $GDownLoadDir & "] Cleared.", 1)
	EndIf
	_MemoLogWrite("Clearing [" & $GDataStoreDir & "].....")
	If DirRemove($GDataStoreDir, 1) Then
		_MemoLogWrite("[" & $GDataStoreDir & "] Cleared.", 1)
	EndIf
	_MemoLogWrite("Clearing [" & $GCATRootDir & "].....")
	DirRemove($GCATRootDir, 1)
	_MemoLogWrite("[" & $GCATRootDir & "] Cleared.", 1)
	$ClearWinUpHist = False

EndFunc

Func _RepairSHC()

	If Not $EventLogConfigured Then _ConfigureEventLog()
	_MemoLogWrite("Repairing SSL / HTTPS / Cryptography service, Please wait.....")
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	_MemoLogWrite("Configuring the Cryptographic Service.....")
	_SvcSetStartMode("CryptSvc","Automatic")
	_MemoLogWrite("Cryptographic Service Configured.")
	_MemoLogWrite("Stopping the Cryptographic Service.....")
	If Not _SvcStop("CryptSvc") Then
		_MemoLogWrite("Cryptographic service was not started in the first place.", 3)
	Else
		_MemoLogWrite("Cryptographic service Stopped Successfully.", 1)
	EndIf
	_MemoLogWrite("Clearing [" & @WindowsDir & "\system32\CatRoot].....", 1)
	;DirRemove(@WindowsDir&"\system32\CatRoot, 1)
	FileDelete(@WindowsDir & "\system32\CatRoot\{F750E6C3-38EE-11D1-85E5-00C04FC295EE}\tmp*.CAT")
	FileDelete(@WindowsDir & "\system32\CatRoot\{127D0A1D-4EF2-11D1-8608-00C04FC295EE}\tmp*.CAT")
	FileDelete(@WindowsDir & "\system32\CatRoot\{F750E6C3-38EE-11D1-85E5-00C04FC295EE}\KB*.CAT")
	FileDelete(@WindowsDir & "\system32\CatRoot\{127D0A1D-4EF2-11D1-8608-00C04FC295EE}\KB*.CAT")
	FileDelete(@WindowsDir & "\inf\oem*.*")
	_MemoLogWrite("[" & @WindowsDir & "\system32\CatRoot] cleared." , 1)
	_MemoLogWrite("Re-registering SSL / HTTPS / Cryptography DLLs.....")
	_ReregisterDLL("cryptdlg.dll")
	_ReregisterDLL("cryptext.dll")
	_ReregisterDLL("cryptui.dll")
	_ReregisterDLL("dssenh.dll")
	_ReregisterDLL("gpkcsp.dll")
	_ReregisterDLL("initpki.dll")
	_ReregisterDLL("licdll.dll")
	_ReregisterDLL("mssign32.dll")
	_ReregisterDLL("mssip32.dll")
	_ReregisterDLL("regwizc.dll")
	_ReregisterDLL("rsaenh.dll")
	_ReregisterDLL("scardssp.dll")
	_ReregisterDLL("sccbase.dll")
	_ReregisterDLL("scecli.dll")
	_ReregisterDLL("slbcsp.dll")
	_ReregisterDLL("softpub.dll")
	_ReregisterDLL("winhttp.dll")
	_ReregisterDLL("wintrust.dll")
	_MemoLogWrite("SSL / HTTPS / Cryptography DLLs re-registered.")
	FileSetAttrib(@WindowsDir, "-RSH")
	FileSetAttrib(@SystemDir, "-RSH")
	FileSetAttrib(@WindowsDir & "\system32\CatRoot", "-RSH", 1)
	_MemoLogWrite("Restarting the Cryptographic Service.....")
	If Not _SvcStart("CryptSvc") Then
		_MemoLogWrite("The Cryptographic Service could not be started.", 2)
	Else
		_MemoLogWrite("Cryptographic Service restarted.", 1)
	EndIf
	_MemoLogWrite("Finished repairing SSL / HTTPS / Cryptography service.")
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)

EndFunc


Func _ResetFirewall()

	Local $ERRORCode

	_MemoLogWrite("Resetting the Windows Firewall configuraton, Please wait.....")
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	Switch @OSVersion
		Case "WIN_2000", "WIN_XP", "WIN_XPe", "WIN_2003"
			$ERRORCode = ShellExecuteWait("netsh", "firewall reset", "", "", @SW_HIDE)
		Case "WIN_VISTA", "WIN_2008", "WIN_7", "WIN_2008R2", "WIN_8", "WIN_81"
			$ERRORCode = ShellExecuteWait("netsh", "advfirewall reset", "", "", @SW_SHOW)
	EndSwitch
	If $ERRORCode = 0 Then
		_MemoLogWrite("Windows Firewall configuration reset successful.", 1)
	Else
		_MemoLogWrite("Could not reset Windows Firewall configuration.", 2)
	EndIf
	_MemoLogWrite("Finished resetting the Windows Firewall configuraton.")
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)

EndFunc


Func _ConfigureEventLog()

	_MemoLogWrite("Configuring the Windows Event Log Service, Please wait.....")
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	_SvcSetStartMode("eventlog","Automatic")
	_MemoLogWrite("Windows Event Log Service Configured.", 1)
	_MemoLogWrite("Starting the Windows Event Log Service.....")
	If Not _SvcStart("eventlog") Then
		_MemoLogWrite("The Windows Event Log Service could not be started.", 2)
		_MemoLogWrite("Attempting to repair the Windows Event Log Service.....")
		Switch @OSVersion
			Case "WIN_XP", "WIN_2003"
				RegWrite(	"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Eventlog", "Description", "REG_SZ", "Enables event log messages " & _
							"issued by Windows-based programs and components to be viewed in Event Viewer. This service cannot be stopped.")
				RegWrite(	"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Eventlog", "DisplayName", "REG_SZ", "Event Log")
				RegWrite(	"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Eventlog", "ErrorControl", "REG_DWORD", 0x00000001)
				RegWrite(	"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Eventlog", "Group", "REG_SZ", "Event log")
				RegWrite(	"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Eventlog", "ImagePath", "REG_EXPAND_SZ", _
							"%SystemRoot%\system32\services.exe")
				RegWrite(	"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Eventlog", "ObjectName", "REG_SZ", "LocalSystem")
				RegWrite(	"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Eventlog", "PlugPlayServiceType", "REG_DWORD", 0x00000003)
				RegWrite(	"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Eventlog", "Start", "REG_DWORD", 0x00000002)
				RegWrite(	"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Eventlog", "Type", "REG_DWORD", 0x00000020)
			Case "WIN_VISTA", "WIN_2008", "WIN_7", "WIN_2008R2"
				_MemoLogWrite("Repairing the Windows Event Log Service.....")
				RegWrite(	"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\eventlog", _
							"ObjectName", "REG_SZ", "NT AUTHORITY\LocalService")
				Sleep(250)
				If Not _SvcStart("eventlog") Then
					_MemoLogWrite("The Windows Event Log Service could not be repaired and started.", 2)
				Else
					_MemoLogWrite("Windows Event Log Service repaired Successfully.", 1)
				EndIf
		EndSwitch
	Else
		_MemoLogWrite("Windows Event Log Service Started Successfully.", 1)
	EndIf
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	$EventLogConfigured = True

EndFunc

Func _RepairWorkGroups()

	_MemoLogWrite("Repairing Workgroup Computers view, Please wait.....")
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	RegDelete("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\NetBt\Parameters","NodeType")
	RegDelete("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\NetBt\Parameters","DhcpNodeType")
	_MemoLogWrite("Finished repairing Workgroup Computers view.")
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)

EndFunc

Func _RestoreHosts()

	_MemoLogWrite("Restoring the default Windows HOSTS file, Please wait.....")
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	Local $lHOSTS = @WindowsDir & "\System32\drivers\etc\hosts"

	FileSetAttrib($lHOSTS, "-RASHNOT")
	FileMove($lHOSTS, @SystemDir & "\drivers\etc\hosts.bak")
	FileDelete($lHOSTS)

	Local $oHOSTS = FileOpen($lHOSTS, 1)
	If $lHOSTS = -1 Then
	EndIf
	_MemoLogWrite("Writing data to the HOSTS file.....")
	FileWrite($oHOSTS, "# Copyright (c) 1993-1999 Microsoft Corp." & @CRLF)
	FileWrite($oHOSTS, "#" & @CRLF)
	FileWrite($oHOSTS, "# This is a sample HOSTS file used by Microsoft TCP/IP for Windows." & @CRLF)
	FileWrite($oHOSTS, "#" & @CRLF)
	FileWrite($oHOSTS, "# This file contains the mappings of IP addresses to host names. Each" & @CRLF)
	FileWrite($oHOSTS, "# entry should be kept on an individual line. The IP address should" & @CRLF)
	FileWrite($oHOSTS, "# be placed in the first column followed by the corresponding host name." & @CRLF)
	FileWrite($oHOSTS, "# The IP address and the host name should be separated by at least one" & @CRLF)
	FileWrite($oHOSTS, "# space." & @CRLF)
	FileWrite($oHOSTS, "#" & @CRLF)
	FileWrite($oHOSTS, "# Additionally, comments (such as these) may be inserted on individual" & @CRLF)
	FileWrite($oHOSTS, "# lines or following the machine name denoted by a '#' symbol." & @CRLF)
	FileWrite($oHOSTS, "#" & @CRLF)
	FileWrite($oHOSTS, "# For example:" & @CRLF)
	FileWrite($oHOSTS, "#" & @CRLF)
	FileWrite($oHOSTS, "#      102.54.94.97     rhino.acme.com          # source server" & @CRLF)
	FileWrite($oHOSTS, "#       38.25.63.10     x.acme.com              # x client host" & @CRLF)
	FileWrite($oHOSTS, "" & @CRLF)
	Switch @OSVersion
		Case "WIN_XP", "WIN_2003"
			FileWrite($oHOSTS, "127.0.0.1       localhost" & @CRLF)
		Case "WIN_VISTA", "WIN_2008"
			FileWrite($oHOSTS, "127.0.0.1       localhost" & @CRLF)
			FileWrite($oHOSTS, "::1             localhost" & @CRLF)
		Case "WIN_7", "WIN_2008R2", "WIN_8", "WIN_81"
			FileWrite($oHOSTS, "# localhost name resolution is handle within DNS itself." & @CRLF)
			FileWrite($oHOSTS, "#       127.0.0.1       localhost" & @CRLF)
			FileWrite($oHOSTS, "#       ::1             localhost" & @CRLF)
	EndSwitch
	FileClose($oHOSTS)
	_MemoLogWrite("HOSTS file created successfully.")
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)

EndFunc

;~ Func _RESETHOSTS()
;~ 	MEMOLOGWRITE("Creating new " & $OSVERSION & " HOSTS file...")
;~
;~ 	If $LOCALHOSTFILE3454 = -1 Then
;~ 		MsgBox(0, "Error", "Unable to open file.")
;~ 		Exit
;~ 	EndIf

;~ EndFunc

Func _ShowWinsockLSPs()

	_StartProcess()
	_LogWrite("Generating List of Installed Winsock LSPs, Please wait.....", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	FileSetAttrib($CIRResLSP, "-RASHNOT")
	RunWait(@ComSpec & ' /c netsh winsock show catalog >"' & $CIRResLSP & '"', "", @SW_HIDE)
	If FileExists($CIRResLSP) Then
		_MemoLogWrite("Winsock LSPs List Saved to '" & $CIRResLSP & "'", 1)
		_OpenTextFile($CIRResLSP)
	Else
		_MemoLogWrite("Could not save LSP list.", 2)
	EndIf
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	_EndProcess()

EndFunc


Func _OpenEventViewer()
	_StartProcess()
	If Not $EventLogConfigured Then _ConfigureEventLog()
	ShellExecute("eventvwr")
	_EndProcess()
EndFunc


Func _OpenTextFile($TXTFileName)
	If FileExists($TXTFileName) Then
		_MemoLogWrite("Opening [" & $TXTFileName & "]")
		If FileExists($Note2EXE) Then
			ShellExecute($Note2EXE, $TXTFileName)
		Else
			ShellExecute($TXTFileName)
		EndIf
	Else
		_MemoLogWrite("Could not find the [" & $TXTFileName & "] file.", 2)
	EndIf
EndFunc

Func _Options()
	_OptionsDlg($mForm)
EndFunc

Func _SpeedTest()
	ShellExecute("http://www.speedtest.net")
EndFunc

Func _GetRouterPasswords()
	ShellExecute("http://www.routerpasswords.com")
EndFunc


Func _OptionsDlg($hParent = 0)

	_LoadSettings()

	Local $ODlg, $ChkCREvent, $ChkEnLogging, $inLogSize, $lblOpLSize, $BtnDelLog, $BtnOpSave, $BtnOpCancel, $nMsg

	Opt("GUIOnEventMode", 0)
	WinSetOnTop($hParent, "", 0)
	GUISetState(@SW_DISABLE, $hParent)

	$ODlg = GUICreate("Preferences", 450, 300, -1, -1)
	GUISetFont(8.5, 400, 0, "Verdana")
	GUISetIcon(@ScriptFullPath, 215)

	GUICtrlCreateLabel("Logging", 10, 53, 60, 20)
	GUICtrlCreateLabel("", 70, 60, 350, 2, $SS_ETCHEDHORZ)
	$ChkEnLogging = GUICtrlCreateCheckbox("Enable logging", 20, 80, 360, 20)
	GUICtrlCreateLabel("Log size must not exceed :", 20, 107, 160, 20)
	$inLogSize = GUICtrlCreateInput($LOGGING_MAXSIZE, 180, 105, 100, 20, $ES_RIGHT)
	GUICtrlSetFont(-1, 9, 400, 0, "Verdana")
	GUICtrlCreateLabel("KB", 290, 107, 50, 20)
	$lblOpLSize = GUICtrlCreateLabel(	"Log size: " & Round(FileGetSize($LOGGING_COMINTREP) / 1024, 2) & _
										" KB", 20, 130, 250, 20)
	GUICtrlSetColor(-1, 0x066186)
	$BtnDelLog = GUICtrlCreateButton("Delete", 340, 110, 100, 30, $WS_GROUP)
	$BtnOpSave = GUICtrlCreateButton("Save", 230, 250, 100, 30, $WS_GROUP)
	$BtnOpCancel = GUICtrlCreateButton("Cancel", 340, 250, 100, 30, $WS_GROUP)

	If $LOGGING_ENABLE = 1 Then GuiCtrlSetState($ChkEnLogging, $GUI_CHECKED)
	If GUICtrlRead($ChkEnLogging) = $GUI_UNCHECKED Then GUICtrlSetState($inLogSize, $GUI_DISABLE)

	GUISetState(@SW_SHOW)

	GuiCtrlSetState($BtnOpSave, $GUI_FOCUS)

	While 1
		Local $nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE, $BtnOpCancel
				ExitLoop
			Case $ChkEnLogging
				If GUICtrlRead($ChkEnLogging) = $GUI_CHECKED Then
					GUICtrlSetState($inLogSize, $GUI_ENABLE)
				Else
					GUICtrlSetState($inLogSize, $GUI_DISABLE)
				EndIf
			Case $BtnOpSave
				Local $lMS = GUICtrlRead($inLogSize)
				If GUICtrlRead($ChkEnLogging) = $GUI_CHECKED Then
					IniWrite(@ScriptDir & "\CIntRep.ini", "Logging", "LogEnabled", 1)
					If StringIsInt($lMS) = 1 Then
						IniWrite(@ScriptDir & "\CIntRep.ini", "Logging", "LogMaxSize", $lMS)
					EndIf
				Else
					IniWrite(@ScriptDir & "\CIntRep.ini", "Logging", "LogEnabled", 0)
				EndIf
				_LoadSettings()
			Case $BtnDelLog
				FileDelete($LOGGING_COMINTREP)
				GuiCtrlSetData($lblOpLSize, "Log size: " & Round(FileGetSize($LOGGING_COMINTREP) / 1024, 2) & " KB")

		EndSwitch
	WEnd

	Opt("GUIOnEventMode", 1)
	GUISetState(@SW_ENABLE, $hParent)
	GUIDelete($ODlg)

EndFunc


Func _RegistryWriter($KeyName, $ValueName, $Type, $Value)

	Local $RegReturn = RegWrite($KeyName, $ValueName, $Type, $Value)
	If Not $RegReturn Then
		Switch @error
			Case 1
				_MemoLogWrite("Unable to open requested registry key.", 2)
			Case 2
				_MemoLogWrite("Unable to open requested main registry key.", 2)
			Case 3
				_MemoLogWrite("Unable to remote connect to the registry.", 2)
			Case -1
				_MemoLogWrite("Unable to open requested registry value.", 2)
			Case -2
				_MemoLogWrite("Registry value type not supported.", 2)
		EndSwitch
	EndIf

EndFunc

Func _ReregisterDLL($FilePath, $Param = "/s")

	Local $RSVR32Error
	If Not $Cancel Then
		;~ _MemoLogWrite("RegSvr32.exe: Registering '" & $FilePath & "'.....")
		$RSVR32Error = ShellExecuteWait("regsvr32.exe", " " & $Param & " " & $FilePath, "")
		Switch $RSVR32Error
			Case 0
				_MemoLogWrite("RegSvr32.exe: " & $FilePath & "' registration succeeded.", 1)
			Case 1
				_MemoLogWrite("RegSvr32.exe: " & $FilePath & "' To register a module, you must provide a binary name.", 2)
			Case 3
				_MemoLogWrite("RegSvr32.exe: " & $FilePath & "' Specified module not found", 2)
			Case 4
				_MemoLogWrite("RegSvr32.exe: " & $FilePath & "' Module loaded but entry-point DllRegisterServer was not found.")
			Case 5
				_MemoLogWrite("RegSvr32.exe: " & $FilePath & "' Error number: 0x80070005", 2)
		EndSwitch
	EndIf
	If $RSVR32Error >= 1 Then
		Return 0
	Else
		Return 1
	EndIf

EndFunc   ;==>_ReregisterDLL


Func _MemoLogWrite($Message = "", $iWarning = 0, $bTStamp = True)

	Local $sPrefix = ""

	Select
		Case $iWarning = 1
			GuiCtrlSetColor($eStatus, 0x006EC3)
		Case $iWarning = 2
			GuiCtrlSetColor($eStatus, 0xA23538)
		Case $iWarning = 3
			GuiCtrlSetColor($eStatus, 0xD14424)
	EndSelect
	Sleep(10)

	_GUICtrlEdit_AppendText($eStatus, $sPrefix & "--> " & $Message & @CRLF)
	_LogWrite($sPrefix & $Message, $bTStamp)

EndFunc


Func _LogWrite($Message = "", $bTStamp = True)

	Local $OpenLog, $sTStamp = ""

	If $LOGGING_ENABLE = 1 Then

		$OpenLog = FileOpen($LOGGING_COMINTREP, 1)
		If $OpenLog = -1 Then
		EndIf

		If $bTStamp Then $sTStamp = "[" & @MDAY & "/" & @MON & "/" & @YEAR & _
									" " & @HOUR & ":" & @MIN & ":" & @SEC & "] "
		FileWrite($OpenLog, $sTStamp & $Message & @CRLF)
		FileClose($OpenLog)

	EndIf

EndFunc

Func _OpenLoggingDirectory()

	ShellExecute(@ScriptDir & "\Logging")
	If @error Then
		_MemoLogWrite("Could not open [" & @ScriptDir & "\Logging].", 2)
	Else
		_MemoLogWrite("The 'logging' directory should now be open.", 1)
	EndIf

EndFunc

Func _OpenCIRLog()
	_OpenTextFile($LOGGING_COMINTREP)
EndFunc


Func _OpenCIRResetLog()
	_OpenTextFile($CIRResLog)
EndFunc

Func _StartProcess()

	GuiCtrlSetState($FileMenu, $GUI_DISABLE)
	GuiCtrlSetState($CommMenu, $GUI_DISABLE)
	GuiCtrlSetState($OpMenu, $GUI_DISABLE)
	GuiCtrlSetState($HelpMenu, $GUI_DISABLE)
	GUICtrlSetState($AppIcon, $GUI_HIDE)
	GUICtrlSetState($PrAni, $GUI_SHOW)

	GUICtrlSetData($eStatus, "")
	GUICtrlSetData($BtnGo, "Stop!")
	GuiCtrlSetState($BtnGo, $GUI_FOCUS)
	GUISetCursor(15)
	For $i = 0 To $RepCount
		GuiCtrlSetState($BtnRep[$i], $GUI_DISABLE)
	Next

EndFunc

Func _EndProcess()

	GuiCtrlSetState($FileMenu, $GUI_ENABLE)
	GuiCtrlSetState($CommMenu, $GUI_ENABLE)
	GuiCtrlSetState($OpMenu, $GUI_ENABLE)
	GuiCtrlSetState($HelpMenu, $GUI_ENABLE)
	GUICtrlSetState($PrAni, $GUI_HIDE)
	GUICtrlSetState($AppIcon, $GUI_SHOW)

	GUICtrlSetData($BtnGo, "Go!")
	GuiCtrlSetData($lblWelc,	"All tasks completed. You can ignore most of the errors, because not all " & _
								"the files being re-registered are the same on every computer. The errors are " & _
								"displayed to keep track with what is happening to your computer.")
	GuiCtrlSetColor($lblWelc, 0x066186)

	For $i = 0 To $RepCount - 7
		GuiCtrlSetState($ChkRep[$i], $GUI_ENABLE)
	Next

	For $i = 0 To $RepCount
		GUICtrlSetImage($IcoRep[$i], @ScriptFullPath, 201 + $i)
		GuiCtrlSetState($IcoRep[$i], $GUI_ENABLE)
		GuiCtrlSetState($ChkRep[$i], $GUI_ENABLE)
		GuiCtrlSetState($BtnRep[$i], $GUI_ENABLE)
	Next

	GUISetCursor(-1)

	GuiCtrlSetState($BtnGo, $GUI_FOCUS)
	SoundPlay(@ScriptDir & "\Sounds\complete.wav")

EndFunc

Func _BootMessage()

	Local $MBox
	_MemoLogWrite("You will need to reboot your computer before the settings will take effect.", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)
	$MBox = MsgBox(65,	"Reboot required!","You will need to reboot your computer before the settings will take effect. " & _
						"Answer 'OK' to reboot your computer or 'Cancel' if you would like to reboot later. " & _
						"Note that some settings might not take effect or some components might not function correctly until you reboot." & @CRLF & @CRLF & _
						"Your computer will reboot automatically in 60 seconds.", 60)
						;_MemoLogWrite("You will need to reboot your computer before the settings will take effect.", 3)
	Switch $MBox
		Case 1, -1
			_MemoLogWrite("Your computer is restarting now.....", 1)
			_Reboot()
		Case 2
			_MemoLogWrite("Reboot Canceled.", 3)
	EndSwitch
	_LogWrite("", False)
	_LogWrite("-----------------------------------------------------------------------------------------", False)

EndFunc


Func _AboutDlg()

	GuiCtrlSetState($HlpAbout, $GUI_DISABLE)

	Local $abTitle, $abVersion, $abCopyright
	Local $abHome, $abGNU
	Local $abSpaceLabel, $abSpaceProg, $abBtnOK
	Local $abPayPal, $abFacebook, $abTwittter
	Local $abLinkedIn, $abGoogle

	$aboutDlg = GUICreate("About " & $APPSET_TITLE, 400, 500, -1, -1, BitOr($WS_CAPTION, $WS_POPUPWINDOW), $WS_EX_TOPMOST)
	GUISetFont(8.5, 400, 0, "Verdana", $AboutDlg, 5)
	GUISetIcon(@ScriptFullPath, 221)

	GUISetOnEvent($GUI_EVENT_CLOSE, "_CloseAboutDlg", $AboutDlg)

	GUICtrlCreateIcon(@ScriptFullPath, 99, 10, 10, 64, 64)
	$abPayPal = GUICtrlCreateIcon(@ScriptFullPath, 220, 320, 0, 64, 64)
	GUICtrlSetTip($abPayPal, "Help us keep our software free.")
	GUICtrlSetCursor($abPayPal, 0)
	$abTitle = GUICtrlCreateLabel($APPSET_TITLE, 88, 16, 220, 18)
	GuiCtrlSetFont($abTitle, 10)
	$abVersion = GUICtrlCreateLabel("Version " & FileGetVersion(@ScriptFullPath), 88, 40, 220, 20)
	$abCopyright = GUICtrlCreateLabel("Copyright © 2014 Rizonesoft", 88, 55, 220, 20)
	GuiCtrlSetColor($abCopyright, 0x555555)

	GUICtrlCreateLabel("Rizonesoft Home: ", 20, 90, 130, 15, $SS_RIGHT)
	$abHome = GUICtrlCreateLabel("www.rizonesoft.com", 155, 90, 200, 15)
	GuiCtrlSetFont($abHome, -1, -1, 4) ;Underlined
	GuiCtrlSetColor($abHome, 0x0000FF)
	GuiCtrlSetCursor($abHome, 0)
	$abGNU = GUICtrlCreateLabel("This program is free software: you can redistribute it and/or modify " & _
								"it under the terms of the GNU General Public License as published by " & _
								"the Free Software Foundation, either version 3 of the License, or " & _
								"(at your option) any later version." & @CRLF & @CRLF & _
								"This program is distributed in the hope that it will be useful, " & _
								"but WITHOUT ANY WARRANTY; without even the implied warranty of " & _
								"MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the " & _
								"GNU General Public License for more details.", 20, 125, 350, 180)
	GuiCtrlSetColor($abGNU, 0x555555)
	GUICtrlCreateLabel(	"Contributors: Derick Payne (Rizonesoft), Matthew McMullan (NerdFencer), Saunders", 20, 280, 350, 100)

	Local $ScriptDirSplt = StringSplit(@ScriptDir, "\")
	Local $ScriptDrive  = $ScriptDirSplt[1]
	Local $drvSpaceUsed = DriveSpaceTotal($ScriptDrive) - DriveSpaceFree($ScriptDrive)

	$abSpaceLabel = GUICtrlCreateLabel("(" & $ScriptDrive & ") " & Round(DriveSpaceFree($ScriptDrive) / 1024, 1) & " GB free of " & _
					Round(DriveSpaceTotal($ScriptDrive) / 1024, 1) & " GB", 15, 380, 300, 15)
	$abSpaceProg = GUICtrlCreateProgress(15, 400, 350, 15)
	GUICtrlSetData($abSpaceProg, ($drvSpaceUsed / DriveSpaceTotal($ScriptDrive)) * 100)
	$abBtnOK = GUICtrlCreateButton("OK", 250, 450, 123, 33, $BS_DEFPUSHBUTTON)

	$abFacebook = GUICtrlCreateIcon(@ScriptFullPath, 216, 20, 450, 32, 32)
	GUICtrlSetTip($abFacebook, "Like us on Facebook and stay updated.")
	GUICtrlSetCursor($abFacebook, 0)
	$abTwittter = GUICtrlCreateIcon(@ScriptFullPath, 217, 55, 450, 32, 32)
	GUICtrlSetTip($abTwittter, "Follow us on Twitter for the latest updates.")
	GUICtrlSetCursor($abTwittter, 0)
	$abLinkedIn = GUICtrlCreateIcon(@ScriptFullPath, 218, 90, 450, 32, 32)
	GUICtrlSetTip($abLinkedIn, "Find us on LinkedIn.")
	GUICtrlSetCursor($abLinkedIn, 0)
	$abGoogle = GUICtrlCreateIcon(@ScriptFullPath, 219, 125, 450, 32, 32)
	GUICtrlSetTip($abGoogle, "Find us on Google.")
	GUICtrlSetCursor($abGoogle, 0)

	GUICtrlSetOnEvent($abHome, "_HomePageClicked")
	GUICtrlSetOnEvent($abFacebook, "_OpenFacebook")
	GUICtrlSetOnEvent($abTwittter, "_FollowOnTwitter")
	GUICtrlSetOnEvent($abLinkedIn, "_OpenLinkedIn")
	GUICtrlSetOnEvent($abGoogle, "_OpenGoogle")
	GUICtrlSetOnEvent($abBtnOK, "_CloseAboutDlg")
	GUICtrlSetOnEvent($abPayPal, "_DonateSomething")

	GUISetState(@SW_SHOW, $AboutDlg)


EndFunc

Func _CloseAboutDlg()

	GuiCtrlSetState($HlpAbout, $GUI_ENABLE)
	GUIDelete($aboutDlg)

EndFunc

Func _HomePageClicked()
	ShellExecute("http://www.rizonesoft.com")
EndFunc


Func _OpenFacebook()
	ShellExecute("https://www.facebook.com/rizonesoft")
EndFunc


Func _FollowOnTwitter()
	ShellExecute("https://twitter.com/rizonesoft")
EndFunc


Func _OpenLinkedIn()
	ShellExecute("http://www.linkedin.com/in/rizonesoft")
EndFunc


Func _OpenGoogle()
	ShellExecute("https://plus.google.com/+Rizonesoftsa/posts")
EndFunc


Func _DonateSomething()
	ShellExecute("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=7UGGCSDUZJPFE")
EndFunc