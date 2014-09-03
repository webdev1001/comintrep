#include-once

; #INDEX# =======================================================================================================================
; Title .........: Configuration
; AutoIt Version : 3.3.8.1
; Language ......: English
; Description ...: Stores and retrieves global configuration or settings.
; Author(s) .....: Derick Payne (Rizonesoft)
; Dll(s) ........:
; ===============================================================================================================================



; #CURRENT# =====================================================================================================================
;_GetGlobalSettings
; ===============================================================================================================================

; #DEFAULT SETTINGS# ===================================================================================================================
Global $DOUBLETAP_EXIT = False
Global $DOORSPROC_SLEEP
Global $DOORSTHEME_WINCOLOR
Global $DOORSTHEME_FONTSIZE
Global $DOORSTHEME_FONTWEIGHT = 400
Global $DOORSTHEME_FONTATTRIBUTE = Default
Global $DOORSTHEME_FONTNAME
Global $DOORSTHEME_FONTQUALITY
Global $DOORSTHEME_TOOLTIPSTYLE
; ===============================================================================================================================

_GetGlobalSettings()

Func _GetGlobalSettings()

	$DOORSPROC_SLEEP = IniRead(@ScriptDir & "\doors.ini", "Doors System", "ProcessSleep", 55)
	$DOORSTHEME_WINCOLOR = IniRead(@ScriptDir & "\doors.ini", "Doors Theme", "WinColor", 0xEBEBEB)
	$DOORSTHEME_FONTSIZE = IniRead(@ScriptDir & "\doors.ini", "Doors Theme", "FontSize", 8.5)
	$DOORSTHEME_FONTNAME = IniRead(@ScriptDir & "\doors.ini", "Doors Theme", "FontName", "Verdana")
	$DOORSTHEME_FONTQUALITY = IniRead(@ScriptDir & "\doors.ini", "Doors Theme", "FontSmoothing", 5)
	$DOORSTHEME_TOOLTIPSTYLE = IniRead(@ScriptDir & "\doors.ini", "Doors Theme", "ToolTipStyle", 2)

EndFunc