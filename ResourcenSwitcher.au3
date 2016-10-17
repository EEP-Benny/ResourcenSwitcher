#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=ResSwitchmultiple.ico
#AutoIt3Wrapper_Outfile=ResourcenSwitcher2.exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=2.2.0.2
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=P
#AutoIt3Wrapper_Res_Language=1031 ;German (Germany)
#AutoIt3Wrapper_Res_Comment=
#AutoIt3Wrapper_Res_Description="Resourcen-Switcher"
#AutoIt3Wrapper_Res_Icon_Add=folder_add.ico
#AutoIt3Wrapper_Res_Icon_Add=folder_edit.ico
#AutoIt3Wrapper_Res_Icon_Add=folder_delete.ico
#AutoIt3Wrapper_Res_Icon_Add=update.ico
#AutoIt3Wrapper_Res_Icon_Add=database.ico
#AutoIt3Wrapper_Res_Icon_Add=database_go.ico
#AutoIt3Wrapper_Res_Icon_Add=database_link.ico
#AutoIt3Wrapper_Res_Icon_Add=database_link_go.ico
#AutoIt3Wrapper_Res_Icon_Add=database_error.ico
#AutoIt3Wrapper_Res_Icon_Add=link.ico
#AutoIt3Wrapper_Res_Icon_Add=link_go.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;Todo: 	- Nach Fehlermeldung (Resourcen ist Ordner) Icons aktualisieren/entfernen
;		- Fensterposition auf Bildschirm zurückholen
;		- Umbenennen: Neuer Ordner-Button

#Region - Include Parameters
#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <GuiListView.au3>
#include <GuiImageList.au3>
#include <GuiTab.au3>
#include <StaticConstants.au3>
#include <ButtonConstants.au3>
#include <GuiButton.au3>
#include <File.au3>
#include "StringCompareVersions.au3"
#EndRegion - Include Parameters

#Region - Options
Opt("GUIOnEventMode", 1)
Opt("GUICoordMode", 0)
Opt("TrayIconHide", 1)
Opt("GUICloseOnEsc", 0)
Opt("MustDeclareVars", 1)
#EndRegion - Options


Global $ToolName = "ResourcenSwitcher 2.2"
Global $EEPRegPath

Global $IniFileName = StringReplace(StringReplace(@ScriptFullPath,".au3",".ini",-1),".exe",".ini",-1)
;~MsgBox(0,"Ini-File",$IniFileName)
Global $IniSectionSettings = "Settings"
Global $IniSectionsVersions = "EEPVersion"

Global $TxtFilesName = "ResourcenSwitcher 2.txt"

Global $Restart = False ; Flag is set to True if a Restart is wanted

;Schreibrechte prüfen
Local $tempFileName=_TempFile(@ScriptDir,"~")
Local $tempFile=FileOpen($tempFileName,2)
If $tempFile=-1 Then
	MsgBox(48, $ToolName, "Dieses Programm braucht Schreibrechte in seinem eigenen Programmverzeichnis."&@CRLF&"Bitte starte es als Administrator oder gewähre die Schreibrechte manuell."&@CRLF&"Das Programm wird beendet.")
	Exit
EndIf
FileClose($tempFile)
FileDelete($tempFileName)

Global $EEPVersionsCount = IniRead($IniFileName, $IniSectionSettings, "EEPVersionsCount", 0)
If ($EEPVersionsCount < 1) Then
	If 6 = MsgBox(4+32,$ToolName, "Das Programm kennt noch keine EEP-Versionen, für die Resourcenordner gewechselt werden könnten."&@CRLF&"Soll im Internet danach gesucht werden?") Then
		Update()
	EndIf
	Exit ;ohne EEP-Versionen kann das Programm nicht arbeiten. Wenn beim Update neue dazugekommen sind, wird vorher neu gestartet.

;~ 	$EEPVersionsCount = 5
;~ 	If Not IniWrite($IniFileName, $IniSectionSettings, "EEPVersionsCount", $EEPVersionsCount) Then
;~ 		MsgBox(48, $ToolName, "Dieses Programm braucht Schreibrechte in seinem eigenen Programmverzeichnis."&@CRLF&"Bitte starte es als Administrator oder gewähre die Schreibrechte manuell."&@CRLF&"Das Programm wird beendet.")
;~ 		Exit
;~ 	EndIf
;~
;~ 	Local $SectionEEP6[3][2] = [["Name", '"bis EEP6"'], _
;~ 			["RegPath", '"HKCR\Software\Software Untergrund\EEXP"'], _
;~ 			["exeName", '"EEP.exe"']]
;~ 	IniWriteSection($IniFileName, $IniSectionsVersions & "1", $SectionEEP6, 0)
;~ 	Local $SectionEEP7[3][2] = [["Name", '"EEP7"'], _
;~ 			["RegPath", '"HKLM\SOFTWARE\Trend\EEP 7.00\EEXP"'], _
;~ 			["exeName", '"EEP7.exe"']]
;~ 	IniWriteSection($IniFileName, $IniSectionsVersions & "2", $SectionEEP7, 0)
;~ 	Local $SectionEEP8[3][2] = [["Name", '"EEP8"'], _
;~ 			["RegPath", '"HKLM\SOFTWARE\Trend\EEP 8.00\EEXP"'], _
;~ 			["exeName", '"EEP8.exe"']]
;~ 	IniWriteSection($IniFileName, $IniSectionsVersions & "3", $SectionEEP8, 0)
;~ 	Local $SectionEEP9[3][2] = [["Name", '"EEP9 (32bit)"'], _
;~ 			["RegPath", '"HKLM\SOFTWARE\Trend\EEP 9.00\EEXP"'], _
;~ 			["exeName", '"EEP9.exe"']]
;~ 	IniWriteSection($IniFileName, $IniSectionsVersions & "4", $SectionEEP9, 0)
;~ 	Local $SectionEEP964[3][2] = [["Name", '"EEP9 (64bit)"'], _
;~ 			["RegPath", '"HKLM64\SOFTWARE\Trend\EEP 9.00\EEXP"'], _
;~ 			["exeName", '"EEP9x64.exe"']]
;~ 	IniWriteSection($IniFileName, $IniSectionsVersions & "5", $SectionEEP964, 0)
EndIf

Global $EEPVersions_Name[1]
Global $EEPVersions_RegPath[1]
Global $EEPVersions_HasRegResBase[1]
Global $EEPVersions_Nr[1]
For $i = 1 To $EEPVersionsCount
	;Prüfen, ob die EEP-Version (in der Registry) existiert
	RegRead(IniRead($IniFileName, $IniSectionsVersions & $i, "RegPath", ""), "Directory")
	If @error Then
;~  		MsgBox(0,$ToolName,IniRead($IniFileName,$IniSectionsVersions&$i,"Name","???")&" gibt es nicht")
	Else
		_ArrayAdd($EEPVersions_Name, IniRead($IniFileName, $IniSectionsVersions & $i, "Name", ""))
		_ArrayAdd($EEPVersions_RegPath, IniRead($IniFileName, $IniSectionsVersions & $i, "RegPath", ""))

		;Prüfen, ob "ResBase" in der Registry steht
		RegRead(IniRead($IniFileName, $IniSectionsVersions & $i, "RegPath", ""), "ResBase")
		If @error Then
			_ArrayAdd($EEPVersions_HasRegResBase, False)
		Else
			_ArrayAdd($EEPVersions_HasRegResBase, True)
		EndIf

		_ArrayAdd($EEPVersions_Nr, $i)
	EndIf
Next
;~ _ArrayDisplay($EEPVersions_HasRegResBase)
$EEPVersionsCount = UBound($EEPVersions_Name) - 1

Global $EEPVersionAkt = 1 ;wird später per SetVersion neu gesetzt

;~ Global $IniSectionResourcenFolders = $IniSectionsVersions & $EEPVersionAkt

Global $ResourcenFolder_Descriptions[1]
Global $ResourcenFolder_Paths[1]

Global $EntryToEdit = 0
Global $CurrResFolderLink[$EEPVersionsCount + 1]
Global $CurrResFolderReg[$EEPVersionsCount + 1]

Global $LinkColumnsIcon[$EEPVersionsCount+1]
Global $LinkColumnIcon
Global $RegColumnsIcon[$EEPVersionsCount+1]
Global $RegColumnIcon
Global $LastHoveredItem[2] = [0, 0]

;Unterschied zwischen Fenstergröße und Clientsize rausfinden
Local $TestGUI = GUICreate("Test-GUI", 500, 300, Default, Default, BitOR($WS_SIZEBOX, $WS_MINIMIZEBOX))
Local $Size = WinGetClientSize($TestGUI)
Global $ClientDiff[2] = [500 - $Size[0], 300 - $Size[1]]
GUIDelete($TestGUI)



#Region - Main GUI
;GUI
Global $GUI = GUICreate($ToolName, 280 + $ClientDiff[0], 145 + $ClientDiff[1], Default, Default, BitOR($WS_SIZEBOX, $WS_MINIMIZEBOX))
GUISetOnEvent($GUI_EVENT_CLOSE, "Close")
;~ ;Local $SizeBorder=WinGetPos($GUI)
;~ ;Local $SizeClient=WinGetClientSize($GUI)
;~ ;WinMove($GUI,Default,Default,Default,500,400)
;~ ;MsgBox(0,"Fenstergröße","Angelegt mit 285x175px"&@CRLF&"Aktueller Clientbereich: "&$SizeClient[0]&"x"&$SizeClient[1]&"px"&@CRLF&"Aktuelle Fenstergröße: "&$SizeBorder[2]&"x"&$SizeBorder[3]&"px")

;ImageList mit Pfeil zum Anzeigen des aktuellen Resourcen-Ordners
Global $ImageList = _GUIImageList_Create(16, 16, 5, 1, 4)
;~ ;If @Compiled Then
;~ ;	_GUIImageList_AddIcon($ImageList,@ScriptFullPath,4)
;~ ;Else
;~ ;	_GUIImageList_AddIcon($ImageList,@ScriptDir&"\Pfeil.ico")
;~ ;EndIf
Global $IconKeins=-1
Global $IconRegAktiv=_GUIImageList_AddIcon($ImageList, @ScriptFullPath, 8)		;0 - Reg aktiv
Global $IconRegHover=_GUIImageList_AddIcon($ImageList, @ScriptFullPath, 9)		;1 - Reg Hover
Global $IconRegLinkAktiv=_GUIImageList_AddIcon($ImageList, @ScriptFullPath, 10)	;2 - RegLink aktiv
Global $IconRegLinkHover=_GUIImageList_AddIcon($ImageList, @ScriptFullPath, 11)	;3 - RegLink Hover
Global $IconRegError=_GUIImageList_AddIcon($ImageList, @ScriptFullPath, 12)		;4 - Reg Error
Global $IconLinkAktiv=_GUIImageList_AddIcon($ImageList, @ScriptFullPath, 13)	;5 - Link aktiv
Global $IconLinkHover=_GUIImageList_AddIcon($ImageList, @ScriptFullPath, 14)	;6 - Link Hover

;ImageList für die EEP-Icons
Global $ImageList_EEPIcons = _GUIImageList_Create(16, 16, 4, 1)

;Tabs zum Auswählen der EEP-Version
Global $TabView_EEPVersions = GUICtrlCreateTab(10, 10, 265, 130)
GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)
;~ GUISetOnEvent(-1, "Close")
;~ ;_GUICtrlTab_SetPadding($TabView_EEPVersions,3,3)
_GUICtrlTab_SetImageList($TabView_EEPVersions, $ImageList_EEPIcons)

Global $TabItems_EEPVersions[1]
Global $ListViews_ResourcenFolders[1]
Global $ListView_ResourcenFolders
Global $Buttons_Add[1]
Global $Buttons_Edit[1]
Global $Buttons_Delete[1]
For $i = 1 To $EEPVersionsCount
;~ 	;SetVersion($i)
	_ArrayAdd($TabItems_EEPVersions, GUICtrlCreateTabItem($EEPVersions_Name[$i]))
;~ 	;GUICtrlSetState($TabItems_EEPVersions[$i],$GUI_HIDE)
;~ 	;MsgBox(0,"AppPath",RegRead($EEPVersions_RegPath[$i],"Directory")&"\"&IniRead($IniFileName,$IniSectionsVersions&$EEPVersions_Nr[$i],"exeName","muell.exe"))
	_GUIImageList_AddIcon($ImageList_EEPIcons, RegRead($EEPVersions_RegPath[$i], "Directory") & "\" & IniRead($IniFileName, $IniSectionsVersions & $EEPVersions_Nr[$i], "exeName", "muell.exe"), 0)
;~ 	;MsgBox(0,"Icon-Nr",_GUIImageList_AddIcon($ImageList_EEPIcons,@SystemDir & "\shell32.dll",$i))
;~ 	;MsgBox(0,"Gesamt-Zahl an Icons",_GUIImageList_GetImageCount($ImageList_EEPIcons))
	_GUICtrlTab_SetItemImage($TabView_EEPVersions, $i - 1, $i - 1)

	GUISetCoord(0, 0)

	;Liste zum Auswählen des Resourcenordners
	_ArrayAdd($ListViews_ResourcenFolders, GUICtrlCreateListView("||Beschreibung|Pfad", 20, 40, 210, 85, BitOR($GUI_SS_DEFAULT_LISTVIEW, $LVS_NOSORTHEADER)))
	_GUICtrlListView_SetExtendedListViewStyle($ListViews_ResourcenFolders[$i], BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES))
	GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)
	_GUICtrlListView_SetImageList($ListViews_ResourcenFolders[$i], $ImageList, 1)

	;Button Hinzufügen
	_ArrayAdd($Buttons_Add, GUICtrlCreateButton("&Hinzufügen", 220, 0, 25, 25, $BS_ICON))
	GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)
	GUICtrlSetOnEvent($Buttons_Add[$i], "ShowGUIResourcenFolderAdd")
	GUICtrlSetTip(-1, "Neuen Resourcen-Ordner zur Liste hinzufügen")
	GUICtrlSetImage(-1, @ScriptFullPath, -5, 0)
;~ 	;_GUICtrlButton_SetImageList($Buttons_Add[$i],$ImageList,4)

	;Button Ändern
	_ArrayAdd($Buttons_Edit, GUICtrlCreateButton("Ä&ndern", 0, 30, 25, 25, $BS_ICON))
	GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)
	GUICtrlSetOnEvent($Buttons_Edit[$i], "ShowGUIResourcenFolderEdit")
	GUICtrlSetTip(-1, "Ausgewählten Resourcen-Ordner bearbeiten")
	GUICtrlSetImage(-1, @ScriptFullPath, -6, 0)

	;Button Löschen
	_ArrayAdd($Buttons_Delete, GUICtrlCreateButton("&Löschen", 0, 30, 25, 25, $BS_ICON))
	GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)
	GUICtrlSetOnEvent($Buttons_Delete[$i], "DeleteResourcenFolder")
	GUICtrlSetTip(-1, "Ausgewählten Resourcen-Ordner aus der Liste entfernen")
	GUICtrlSetImage(-1, @ScriptFullPath, -7, 0)

Next

GUICtrlCreateTabItem("") ;Erstellung der Tabs abschließen

;Label Hilfe
Global $Label_Help = GUICtrlCreateLabel("Hilfe", -20, -92, 25, 15, $GUI_SS_DEFAULT_LABEL + $SS_RIGHT)
GUICtrlSetFont(-1, 8.5, 400, 4)
GUICtrlSetColor(-1, 0x0000FF)
GUICtrlSetCursor(-1, 0)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)
GUICtrlSetOnEvent($Label_Help, "ShowHelp")

;Button Update
Global $Button_Update = GUICtrlCreateButton("&Update", 33, -5, 24, 24, $BS_ICON)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)
GUICtrlSetOnEvent(-1, "Update")
GUICtrlSetTip(-1, "Im Internet nach Informationen zu neuen EEP- und Programm-Versionen suchen")
GUICtrlSetImage(-1, @ScriptFullPath, -8, 0)

#cs
	;Button Switch
	Global $Button_Switch=GUICtrlCreateButton("&Switch",-225,30,100,30)
	GUICtrlSetResizing(-1,$GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent($Button_Switch,"SwitchResourcenFolder")

	;Button Hilfe
	Global $Button_Help=GUICtrlCreateButton("&?",110,0,25,30)
	GUICtrlSetResizing(-1,$GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent($Button_Help,"ShowHelp")

	;Button Beenden
	Global $Button_Quit=GUICtrlCreateButton("&Beenden",35,0,100,30)
	GUICtrlSetResizing(-1,$GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT)
	GUICtrlSetOnEvent($Button_Quit,"Close")
#ce

#EndRegion - Main GUI

#Region - Resourcen-Folder GUI
;GUI
Global $ResourcenFolder_GUI = GUICreate("Resourcen-Ordner", 260, 100, Default, Default, BitOR($WS_CAPTION, $WS_POPUP, $WS_SYSMENU, $DS_SETFOREGROUND), Default, $GUI)
GUISetOnEvent($GUI_EVENT_CLOSE, "Cancel")

;Button Pfad durchsuchen
Global $ResourcenFolder_Button_BrowsePath = GUICtrlCreateButton("&Pfad", 10, 10, 70, 20)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)
GUICtrlSetOnEvent($ResourcenFolder_Button_BrowsePath, "BrowsePath")

;Input Pfad
Global $ResourcenFolder_Input_Path = GUICtrlCreateInput("", 75, 0, 165, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT)

;Label Beschreibung
GUICtrlCreateLabel("Beschreibung", -75, 28)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)

;Input Beschreibung
Global $ResourcenFolder_Input_Description = GUICtrlCreateInput("", 75, -3, 165, 20)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKHEIGHT)

;Button OK
Global $ResourcenFolder_Button_OK = GUICtrlCreateButton("&OK", -70, 30, 110, 30)
GUICtrlSetState($ResourcenFolder_Button_OK, $GUI_DEFBUTTON)

;Button Abbrechen
Global $ResourcenFolder_Button_Cancel = GUICtrlCreateButton("&Abbrechen", 120, 0, 110, 30)
GUICtrlSetOnEvent($ResourcenFolder_Button_Cancel, "Cancel")
#EndRegion - Resourcen-Folder GUI

ReadResourcenFolders()
SetVersion(IniRead($IniFileName, $IniSectionSettings, "SelectedEEPVersion", 1)) ;Zuletzt aktive Version einstellen
GUICtrlSetState($TabItems_EEPVersions[$EEPVersionAkt],$GUI_SHOW) ;und den entsprechenden Tab vorwählen
GUISetState(@SW_SHOW, $GUI)
ResizeWindow()
GUIRegisterMsg($WM_GETMINMAXINFO, "WM_NOTIFY") ; Minimale Größe festlegen
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

#Region Functions
;~ Func SwitchResourcenFolder()
;~ 	Local $Index = _GUICtrlListView_GetSelectedIndices($ListView_ResourcenFolders, False)
;~ 	If $Index = "" Then
;~ 		MsgBox(48, $ToolName, "Du musst einen Resourcen-Ordner zum Wechseln auswählen", Default, $GUI)
;~ 		Return
;~ 	EndIf
;~ 	SwitchRegistryTo($Index)
;~ EndFunc   ;==>SwitchResourcenFolder

Func SwitchRegistryTo($Index)
	;Abbrechen, wenn es keinen Registry-Eintrag gibt
	If $EEPVersions_HasRegResBase[$EEPVersionAkt]=False Then
;~ 		MsgBox(0,$ToolName,"Diese Version hat keinen Registry-Eintrag")
		DisplayResourcenFolders($EEPVersionAkt)
		Return
	EndIf
	If $Index >= 0 Then
		Local $path = _GUICtrlListView_GetItemText($ListView_ResourcenFolders, $Index, 3)
	Else
		Local $path = RegRead($EEPRegPath, "Directory") & "\Resourcen"
		$Index = $CurrResFolderLink[$EEPVersionAkt]
	EndIf
	If RegRead($EEPRegPath, "ResBase") <> $path Then
		RegWrite($EEPRegPath, "ResBase", "REG_SZ", $path)
		If @error Then
			MsgBox(48, $ToolName, "Bitte starte dieses Programm mit Schreibrechten in der Registry, z.B. als Administrator." & @CRLF & _
					"Die Zugriffsrechte können sich auf folgenden Pfad/Schlüssel beschränken:" & @CRLF & _
					$EEPRegPath)
			DisplayResourcenFolders($EEPVersionAkt)
			Return
		EndIf
	EndIf
	_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $CurrResFolderReg[$EEPVersionAkt], $IconKeins, 0)
;~ 	_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $Index, 0, 0)
	$CurrResFolderReg[$EEPVersionAkt] = $Index
	DisplayResourcenFolders($EEPVersionAkt)
;~ 	MsgBox(64,$ToolName,"Resourcen-Ordner wurde gewechselt nach"&@CRLF&$path,Default,$GUI)
EndFunc   ;==>SwitchRegistryTo

Func SwitchLinkTo($Index)
	If $LastHoveredItem[1]=1 Then _GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $LastHoveredItem[0], $LinkColumnIcon[$LastHoveredItem[0]], 1)
	_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $LastHoveredItem[0], $RegColumnIcon[$LastHoveredItem[0]], 0)
	_GUICtrlListView_RedrawItems($ListView_ResourcenFolders, $LastHoveredItem[0], $LastHoveredItem[0])
	Local $path = _GUICtrlListView_GetItemText($ListView_ResourcenFolders, $Index, 3)
	Local $homeRes = RegRead($EEPRegPath, "Directory") & "\Resourcen"
	If FileExists($homeRes) And Not DirIsLink(RegRead($EEPRegPath, "Directory"), "Resourcen") Then
		If MsgBox(20, $ToolName, $homeRes &@CRLF& "ist noch ein vollwertiger Resourcen-Ordner, sodass an " & @CRLF & _
		  "seiner Stelle keine Verknüpfung erstellt werden kann."&@CRLF&@CRLF& _
		  "Soll der Resourcen-Ordner umbenannt werden?"&@CRLF&"Den neuen Pfad kannst du gleich festlegen.")=6 Then
;~ 			$EntryToEdit = _ArraySearch($ResourcenFolder_Paths[$EEPVersionAkt],$homeRes)
			$EntryToEdit = _GUICtrlListView_FindInText($ListViews_ResourcenFolders[$EEPVersionAkt],$homeRes)+1
			WinSetTitle($ResourcenFolder_GUI, "", "Resourcen-Ordner umbenennen")
			GUICtrlSetData($ResourcenFolder_Input_Path, $homeRes)
			GUICtrlSetData($ResourcenFolder_Input_Description, _GUICtrlListView_GetItemText($ListViews_ResourcenFolders[$EEPVersionAkt],$EntryToEdit-1,2))
			GUICtrlSetOnEvent($ResourcenFolder_Button_OK, "RenameResourcenFolder")
			GUISetState(@SW_SHOW, $ResourcenFolder_GUI)
;~ 			MsgBox(0,_GUICtrlListView_GetItemText($ListViews_ResourcenFolders[$EEPVersionAkt],$EntryToEdit-1,2),$EntryToEdit-1)
		EndIf
		Return
	EndIf
	If Not FileExists($path) Then ;Prüfen, ob der Zielordner existiert
		MsgBox(48, $ToolName, "Der Ordner"&@CRLF&$path&@CRLF&"existiert nicht und kann daher nicht als Linkziel festgelegt werden.")
		Return
	ElseIf StringInStr(FileGetAttrib($path), "D") = 0 Then ;Prüfen, ob der Zielordner auch wirklich ein Ordner ist.
		MsgBox(48, $ToolName, $path&@CRLF&"ist kein Ordner und kann daher nicht als Linkziel festgelegt werden.")
		Return
	EndIf
	If FileCreateNTFSLink($path, $homeRes, 1) Then
		_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $CurrResFolderLink[$EEPVersionAkt], $IconKeins, 1)
;~ 		_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $Index, 2, 1)
		$CurrResFolderLink[$EEPVersionAkt] = $Index
		If LinkGetRealDir($homeRes) <> $path Then ;Falls der Link von LinkGetRealDir nicht richtig aufgelöst werden kann, den Pfad in eine versteckte Textdatei schreiben
			If FileExists($homeRes & "\" & $TxtFilesName) Then FileSetAttrib($homeRes & "\" & $TxtFilesName, "-H")
			Local $file = FileOpen($homeRes & "\" & $TxtFilesName, 2)
			FileWrite($homeRes & "\" & $TxtFilesName, $path)
			FileClose($file)
			FileSetAttrib($homeRes & "\" & $TxtFilesName, "+H")
		EndIf
		SwitchRegistryTo(-1)
	Else
		MsgBox(48, $ToolName, "Um die Verknüpfung zu erstellen, sind Schreibrechte im EEP-Verzeichnis nötig."&@CRLF&"Bitte starte dieses Programm als Administrator oder gewähre die Schreibrechte manuell.")
	EndIf
EndFunc   ;==>SwitchLinkTo

Func RenameResourcenFolder()
	Local $homeRes = RegRead($EEPRegPath, "Directory") & "\Resourcen"
	Local $path = GUICtrlRead($ResourcenFolder_Input_Path)
	If FileExists($path) Then
		MsgBox(48,$ToolName,"Der gewählte Ordner existiert bereits.")
		Return
	EndIf
	DirMove($homeRes,$path)
	EditResourcenFolder()
EndFunc


Func ShowHelp()
;~ 	Local $PDFtype = RegRead("HKEY_CLASSES_ROOT\.pdf", "")
;~ 	Local $PDFapplication = RegRead("HKEY_CLASSES_ROOT\" & $PDFtype & "\Shell\Open\Command", "")
;~ 	Run(StringReplace($PDFapplication, "%1", StringReplace(StringReplace(@ScriptFullPath,".au3",".pdf",-1),".exe",".pdf",-1)))
	Local $pdfFileName=StringReplace(StringReplace(@ScriptFullPath,".au3",".pdf",-1),".exe",".pdf",-1)
	ShellExecute($pdfFileName)
EndFunc   ;==>ShowHelp

Func Update()
	Local $url=IniRead($IniFileName,$IniSectionSettings,"UpdateURL","")
	If $url="" Then
		$url="http://emaps-eep.de/files?ResourcenSwitcherUpdate.ini"
		If Not IniWrite($IniFileName,$IniSectionSettings,"UpdateURL",$url) Then
			MsgBox(48, $ToolName, "Dieses Programm braucht Schreibrechte in seinem eigenen Programmverzeichnis."&@CRLF&"Bitte starte es als Administrator oder gewähre die Schreibrechte manuell."&@CRLF&"Das Programm wird beendet.")
			Exit
		EndIf
	ElseIf StringInStr($url,"emaps.de.vu") Then
		$url=StringReplace($url,"emaps.de.vu","emaps-eep.de")
		IniWrite($IniFileName,$IniSectionSettings,"UpdateURL",$url)
	EndIf
	Local $tempFileName=_TempFile(@ScriptDir,"~",".ini")
	If InetGet($url,$tempFileName,1)<=0 Then
		MsgBox(48,$ToolName,"Beim Update ist ein Fehler aufgetreten."&@CRLF&"Die Datei konnte nicht heruntergeladen werden.")
		FileDelete($tempFileName)
		Return
	EndIf
	Local $UpdateVersionsCount = IniRead($tempFileName,"General","EEPVersionsCount",-1)
	Local $UpdateMessage=""
	If $UpdateVersionsCount<=0 Then
		$UpdateMessage="Die heruntergeladene Datei enthält keine Informationen zu EEP-Versionen"
		;MsgBox(48,$ToolName,"Beim Update ist ein Fehler aufgetreten."&@CRLF&"Die heruntergeladene Datei enthält keine Informationen zu EEP-Versionen")
	Else
		Local $NewVersions[1] = [0]
		For $i=1 To $UpdateVersionsCount
			Local $UpdateRegPath = IniRead($tempFileName,"EEPVersion"&$i,"RegPath","")
			If $UpdateRegPath="" Then ContinueLoop		;Leere Pfade gelten nicht

			;In der aktuellen ini suchen, ob es schon einen Eintrag mit diesem Registry-Pfad gibt
			Local $CurrentVersionsCount = IniRead($IniFileName,$IniSectionSettings,"EEPVersionsCount",0)
			For $j=1 To $CurrentVersionsCount
				Local $CurrentRegPath = IniRead($IniFileName,$IniSectionsVersions&$j,"RegPath","")
				If $CurrentRegPath = $UpdateRegPath Then
					ContinueLoop 2
				EndIf
			Next

			;Wenn kein bestehender Eintrag gefunden wurde, den neuen reinkopieren
			$CurrentVersionsCount+=1
			IniWrite($IniFileName,$IniSectionSettings,"EEPVersionsCount",$CurrentVersionsCount)

			Local $UpdateName = IniRead($tempFileName,"EEPVersion"&$i,"Name","")
			Local $Section = IniReadSection($tempFileName,"EEPVersion"&$i)
			_ArrayDelete($Section,0) ; IniReadSection gibt noch die Anzahl zurück, die braucht IniWriteSection nicht
			IniWriteSection($IniFileName, $IniSectionsVersions & $CurrentVersionsCount, $Section, 0)

			$NewVersions[0]+=1
			_ArrayAdd($NewVersions,$UpdateName)
		Next
		If $NewVersions[0] = 0 Then
			$UpdateMessage="Es sind keine neuen EEP-Versionen bekannt."
			;MsgBox(64,$ToolName,"Es sind keine neuen EEP-Versionen bekannt")
		Else
			$UpdateMessage="Es wurden "&$NewVersions[0]&" neue EEP-Versionen zur Liste hinzugefügt:"&@CRLF&_ArrayToString($NewVersions,@CRLF,1)
			;MsgBox(64,$ToolName,"Es wurden "&$NewVersions[0]&" neue EEP-Versionen zur Liste hinzugefügt:"&@CRLF&_ArrayToString($NewVersions,@CRLF,1))

			;Neu starten, damit die neuen Einträge erkannt werden
			$Restart=True
			;Close()
		EndIf
	EndIf
	;Installierte Version herausfinden
	Local $CurVersion=FileGetVersion(@ScriptFullPath)
	Local $NewVersion=IniRead($tempFileName,"General","CurrentVersion","0.0.0.0")
	If _StringCompareVersions($CurVersion,$NewVersion)<0 Then
		Local $url=IniRead($tempFileName,"General","DownloadURL","")
		If 6 = MsgBox(4+64+256,$ToolName,$UpdateMessage&@CRLF&@CRLF& _
				"Es gibt eine neuere Version des ResourcenSwitchers:"&@CRLF& _
				$CurVersion&"  Installierte Version"&@CRLF& _
				$NewVersion&"  Neue Version"&@CRLF&@CRLF& _
				"Die neue Version kann unter"&@CRLF& _
				$url&@CRLF& _
				"heruntergeladen werden."&@CRLF&@CRLF& _
				"Soll die Seite jetzt geöffnet werden?") Then
			ShellExecute($url)
		EndIf
	Else
		MsgBox(64,$ToolName,$UpdateMessage&@CRLF&@CRLF&"Es wurde keine neuere Programmversion gefunden.")
	EndIf
	FileDelete($tempFileName)
	If $Restart Then Close() ;Wenn ein Neustart erforderlich ist, dann tue das jetzt
EndFunc

Func ShowGUIResourcenFolderAdd()
	WinSetTitle($ResourcenFolder_GUI, "", "Resourcen-Ordner hinzufügen")
	GUICtrlSetData($ResourcenFolder_Input_Path, "")
	GUICtrlSetData($ResourcenFolder_Input_Description, "")
	GUICtrlSetOnEvent($ResourcenFolder_Button_OK, "AddResourcenFolder")
	GUISetState(@SW_SHOW, $ResourcenFolder_GUI)
EndFunc   ;==>ShowGUIResourcenFolderAdd

Func ShowGUIResourcenFolderEdit()
	Local $SelectedIndex = _GUICtrlListView_GetSelectedIndices($ListViews_ResourcenFolders[$EEPVersionAkt], False)
	If $SelectedIndex = "" Then
		MsgBox(48, $ToolName, "Du musst einen Eintrag zum Ändern auswählen", Default, $GUI)
		Return
	EndIf
	Local $ItemTextArray = _GUICtrlListView_GetItemTextArray($ListViews_ResourcenFolders[$EEPVersionAkt])
	WinSetTitle($ResourcenFolder_GUI, "", "Resourcen-Ordner bearbeiten")
	GUICtrlSetData($ResourcenFolder_Input_Path, $ItemTextArray[4])
	GUICtrlSetData($ResourcenFolder_Input_Description, $ItemTextArray[3])
	GUICtrlSetOnEvent($ResourcenFolder_Button_OK, "EditResourcenFolder")
	GUISetState(@SW_SHOW, $ResourcenFolder_GUI)
	$EntryToEdit = $SelectedIndex + 1
EndFunc   ;==>ShowGUIResourcenFolderEdit

Func AddResourcenFolder()
	Cancel()
	Local $path = GUICtrlRead($ResourcenFolder_Input_Path)
	Local $Description = GUICtrlRead($ResourcenFolder_Input_Description)
	_ArrayAdd($ResourcenFolder_Paths[$EEPVersionAkt], $path)
	_ArrayAdd($ResourcenFolder_Descriptions[$EEPVersionAkt], $Description)
	DisplayResourcenFolders($EEPVersionAkt)
EndFunc   ;==>AddResourcenFolder

Func EditResourcenFolder()
	If $EntryToEdit > UBound($ResourcenFolder_Descriptions[$EEPVersionAkt]) - 1 Or $EntryToEdit <1 Then
		AddResourcenFolder()
		Return
	EndIf
	Cancel()
	Local $path = GUICtrlRead($ResourcenFolder_Input_Path)
	Local $Description = GUICtrlRead($ResourcenFolder_Input_Description)
	Local $Paths = $ResourcenFolder_Paths[$EEPVersionAkt]
	Local $Descriptions = $ResourcenFolder_Descriptions[$EEPVersionAkt]
	$Paths[$EntryToEdit] = $path
	$Descriptions[$EntryToEdit] = $Description
	$ResourcenFolder_Paths[$EEPVersionAkt] = $Paths
	$ResourcenFolder_Descriptions[$EEPVersionAkt] = $Descriptions
	DisplayResourcenFolders($EEPVersionAkt)
EndFunc   ;==>EditResourcenFolder

Func DeleteResourcenFolder()
	Local $SelectedIndex = _GUICtrlListView_GetSelectedIndices($ListViews_ResourcenFolders[$EEPVersionAkt], False)
	If $SelectedIndex = "" Then
		MsgBox(48, $ToolName, "Du musst einen Eintrag zum Löschen auswählen", Default, $GUI)
		Return
	EndIf
	Local $ItemTextArray = _GUICtrlListView_GetItemTextArray($ListViews_ResourcenFolders[$EEPVersionAkt])
	If $ItemTextArray[4] = RegRead($EEPVersions_RegPath[$EEPVersionAkt], "ResBase") Then
		MsgBox(48, $ToolName, "Du kannst nicht den aktuellen Resourcen-Ordner löschen", Default, $GUI)
		Return
	EndIf
	If MsgBox(292, $ToolName, "Soll dieser Resourcen-Ordner wirklich aus der Liste gelöscht werden?" & @CRLF & $ItemTextArray[3] & @TAB & $ItemTextArray[4], Default, $GUI) = 6 Then
		Local $Paths = $ResourcenFolder_Paths[$EEPVersionAkt]
		Local $Descriptions = $ResourcenFolder_Descriptions[$EEPVersionAkt]
		_ArrayDelete($Paths, $SelectedIndex + 1)
		_ArrayDelete($Descriptions, $SelectedIndex + 1)
		$ResourcenFolder_Paths[$EEPVersionAkt] = $Paths
		$ResourcenFolder_Descriptions[$EEPVersionAkt] = $Descriptions
		DisplayResourcenFolders($EEPVersionAkt)
	EndIf
EndFunc   ;==>DeleteResourcenFolder

Func BrowsePath()
	Local $Folder = FileSelectFolder("Resourcen-Ordner wählen", "", 4, RegRead($EEPRegPath, "Directory"), $ResourcenFolder_GUI)
	If $Folder Then
		GUICtrlSetData($ResourcenFolder_Input_Path, $Folder)
	EndIf
EndFunc   ;==>BrowsePath

Func ReadResourcenFolders()
	For $v = 1 To $EEPVersionsCount
		Local $Number=$EEPVersions_Nr[$v]
		Local $Paths[1]
		Local $Descriptions[1]
		For $i = 1 To IniRead($IniFileName, $IniSectionsVersions & $Number, "Count", 0)
			Local $path = IniRead($IniFileName, $IniSectionsVersions & $Number, "Path" & $i, "Konnte Pfad nicht lesen")
			Local $Description = IniRead($IniFileName, $IniSectionsVersions & $Number, "Description" & $i, "Konnte Beschreibung nicht lesen")
			_ArrayAdd($Paths, $path)
			_ArrayAdd($Descriptions, $Description)
		Next
;~ 		Local $ResBase = RegRead($EEPVersions_RegPath[$v], "ResBase")
;~ 		$CurrResFolderReg[$v] = _ArraySearch($Paths, $ResBase) - 1
		_ArrayAdd($ResourcenFolder_Paths, $Paths)
		_ArrayAdd($ResourcenFolder_Descriptions, $Descriptions)
		DisplayResourcenFolders($v)
	Next
EndFunc   ;==>ReadResourcenFolders

Func DisplayResourcenFolders($v = 1)
;~ 	MsgBox(0,"DisplayResourcenFolders",$v)
	GUISetState(@SW_LOCK, $GUI)
	Local $RegIcons[1]
	Local $LinkIcons[1]
	_GUICtrlListView_DeleteAllItems($ListViews_ResourcenFolders[$v])
	Local $homeRes = RegRead($EEPVersions_RegPath[$v], "Directory") & "\Resourcen"
	Local $ResBase = $homeRes
	If $EEPVersions_HasRegResBase[$v] Then $ResBase = RegRead($EEPVersions_RegPath[$v], "ResBase")
;~ 	If $ResBase="" Then 			;Ab EEP11 gibt es keinen ResBase-Eintrag mehr in der Registry
;~ 		$ResBase=$homeRes
;~ 		$ResBaseIsInRegistry = False
;~ 	EndIf
	Local $LinkFolder = LinkGetRealDir($homeRes)
;~ 	MsgBox(0,$ResBase,$LinkFolder)
;~ 	_ArrayDisplay($ResourcenFolder_Paths[$v],"Suche nach '"&$LinkFolder&"'")
	$CurrResFolderLink[$v] = _ArraySearch($ResourcenFolder_Paths[$v], $LinkFolder) - 1
;~ 	Local $Paths = $ResourcenFolder_Paths[$v]
;~ 	MsgBox(0,$LinkFolder,$Paths)
;~ 	MsgBox(0,"@error nach ArraySearch von LinkFolder",@error)
	If $ResBase = $homeRes And $CurrResFolderLink[$v]>-1 Then
		$CurrResFolderReg[$v] = $CurrResFolderLink[$v]
	Else
		$CurrResFolderReg[$v] = _ArraySearch($ResourcenFolder_Paths[$v], $ResBase) - 1
	EndIf
;~ 	MsgBox(0,$CurrResFolderLink[$v],$CurrResFolderReg[$v])
;~ 	MsgBox(0,$ActResFolderIndex,$ActLinkIndex)
	For $i = 1 To UBound($ResourcenFolder_Descriptions[$v]) - 1
		Local $Descriptions = $ResourcenFolder_Descriptions[$v]
		Local $Paths = $ResourcenFolder_Paths[$v]
		GUICtrlCreateListViewItem("||" & $Descriptions[$i] & "|" & $Paths[$i], $ListViews_ResourcenFolders[$v])
		If StringInStr(FileGetAttrib($Paths[$i]), "D") = 0 Then
			GUICtrlSetColor(-1, 0xCCCCCC)		;Nicht vorhandene Ordner grau färben
		EndIf
		_ArrayAdd($RegIcons,$IconKeins)
		_ArrayAdd($LinkIcons,$IconKeins)
	Next
;~ 	MsgBox(0,$CurrResFolderLink[$v],$CurrResFolderReg[$v])
	If ($ResBase = $homeRes Or $ResBase = $LinkFolder) And $CurrResFolderLink[$v] < -1 Then
;~ 		MsgBox(0,"EEP-Version "&$v,"Ein gemeinsamer Resourcen-Ordner")
		$CurrResFolderReg[$v] = _GUICtrlListView_GetItemCount($ListViews_ResourcenFolders[$v])
		$CurrResFolderLink[$v] = $CurrResFolderReg[$v]
		GUICtrlCreateListViewItem("||Aktueller Resourcen-Ordner|" & $LinkFolder, $ListViews_ResourcenFolders[$v])
		GUICtrlSetColor(-1, 0xFF0000)
		_ArrayAdd($RegIcons,$IconKeins)
		_ArrayAdd($LinkIcons,$IconKeins)
	Else
;~ 		MsgBox(0,"EEP-Version "&$v,"Nicht der gleiche Resourcen-Ordner")
;~ 		MsgBox(0,$CurrResFolderReg[$v],$CurrResFolderLink[$v])
		If $CurrResFolderReg[$v] < -1 Then
;~ 			MsgBox(0,"EEP-Version "&$v,"Resourcen-Ordner in der Registry")
			$CurrResFolderReg[$v] = _GUICtrlListView_GetItemCount($ListViews_ResourcenFolders[$v])
			If $EEPVersions_HasRegResBase[$v] Then
				GUICtrlCreateListViewItem("||Aktueller Resourcen-Ordner in der Registry|" & $ResBase, $ListViews_ResourcenFolders[$v])
			Else
				GUICtrlCreateListViewItem("||Aktueller Resourcen-Ordner|" & $ResBase, $ListViews_ResourcenFolders[$v])
			EndIf
			GUICtrlSetColor(-1, 0xFF0000)
			_ArrayAdd($RegIcons,$IconKeins)
			_ArrayAdd($LinkIcons,$IconKeins)
		EndIf
		If $CurrResFolderLink[$v] < -1 Then
;~ 			MsgBox(0,"EEP-Version "&$v,"Resourcen-Ordner per Link")
			$CurrResFolderLink[$v] = _GUICtrlListView_GetItemCount($ListViews_ResourcenFolders[$v])
			GUICtrlCreateListViewItem("||Aktuell verlinkter Resourcen-Ordner|" & $LinkFolder, $ListViews_ResourcenFolders[$v])
			GUICtrlSetColor(-1, 0xFF0000)
			_ArrayAdd($RegIcons,$IconKeins)
			_ArrayAdd($LinkIcons,$IconKeins)
		EndIf
	EndIf
	If $EEPVersions_HasRegResBase[$v] Then
		If $ResBase = $homeRes And $CurrResFolderLink[$v]>-1 Then
			_GUICtrlListView_SetItemImage($ListViews_ResourcenFolders[$v], $CurrResFolderReg[$v], $IconRegLinkAktiv)
			$RegIcons[$CurrResFolderReg[$v]+1]=$IconRegLinkAktiv
		Else
			_GUICtrlListView_SetItemImage($ListViews_ResourcenFolders[$v], $CurrResFolderReg[$v], $IconRegAktiv)
			$RegIcons[$CurrResFolderReg[$v]+1]=$IconRegAktiv
		EndIf
	EndIf
	_GUICtrlListView_SetItemImage($ListViews_ResourcenFolders[$v], $CurrResFolderLink[$v], $IconLinkAktiv, 1)
	$LinkIcons[$CurrResFolderLink[$v]+1]=$IconLinkAktiv
	_GUICtrlListView_SetColumnWidth($ListViews_ResourcenFolders[$v], 0, 24);+0*$EEPVersions_HasRegResBase[$v]) ;0, if ResBase not in Registry, 24 otherwise
	_GUICtrlListView_SetColumnWidth($ListViews_ResourcenFolders[$v], 1, 20);+4*Not $EEPVersions_HasRegResBase[$v]) ;24, if ResBase not in Registry, 20 otherwise
	_GUICtrlListView_SetColumnWidth($ListViews_ResourcenFolders[$v], 2, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($ListViews_ResourcenFolders[$v], 3, $LVSCW_AUTOSIZE)
 	_ArrayDelete($RegIcons,0)
 	_ArrayDelete($LinkIcons,0)
;~ 	_ArrayDisplay($RegIcons,"$RegIcons")
;~ 	_ArrayDisplay($LinkIcons,"$LinkIcons")

	$RegColumnsIcon[$v]=$RegIcons
	$LinkColumnsIcon[$v]=$LinkIcons
	$RegColumnIcon=$RegColumnsIcon[$EEPVersionAkt]
	$LinkColumnIcon=$LinkColumnsIcon[$EEPVersionAkt]
	GUISetState(@SW_UNLOCK, $GUI)
EndFunc   ;==>DisplayResourcenFolders

Func ResizeWindow()
;~ 	Local $Width = 0;_GUICtrlListView_GetColumnWidth($ListView_ResourcenFolders,0)+_GUICtrlListView_GetColumnWidth($ListView_ResourcenFolders,1)+30
;~ 	Local $Height = 0;_GUICtrlListView_GetItemCount($ListView_ResourcenFolders)*17+150
	Local $Width = IniRead($IniFileName, $IniSectionSettings, "Width", 680)
	Local $Height = IniRead($IniFileName, $IniSectionSettings, "Height", 180)
	;If $Width < 500 Then $Width = 500
	;If $Height < 180 Then $Height = 180
	Local $Left = IniRead($IniFileName, $IniSectionSettings, "PosX", (@DesktopWidth - $Width) / 2)
	Local $Top = IniRead($IniFileName, $IniSectionSettings, "PosY", (@DesktopHeight - $Height) / 2)
	WinMove($GUI, "", $Left, $Top, $Width, $Height)
	If ($Left<0 Or $Left>@DesktopWidth-$Width Or $Top<0 Or $Top>@DesktopHeight-$Height) And _
		MsgBox(36,$ToolName,"Das Fenster scheint außerhalb des Bildschirms zu liegen." & @CRLF & "Soll es zurückgeholt werden?")=6 Then
		If $Width>@DesktopWidth Then $Width=@DesktopWidth
		If $Height>@DesktopHeight Then $Height=@DesktopHeight
		If $Left<0 Then $Left=0
		If $Left>@DesktopWidth-$Width Then $Left=@DesktopWidth-$Width
		If $Top<0 Then $Top=0
		If $Top>@DesktopHeight-$Height Then $Top=@DesktopHeight-$Height
		WinMove($GUI, "", $Left, $Top, $Width, $Height)
	EndIf
EndFunc   ;==>ResizeWindow

Func Cancel()
	GUISetState(@SW_HIDE, @GUI_WinHandle)
EndFunc   ;==>Cancel

Func Close()
	Local $WindowPos = WinGetPos($ToolName)
	If @error=0 Then ;Fensterpositionen nur speichern, wenn es sie auch gibt
		IniWrite($IniFileName, $IniSectionSettings, "PosX", $WindowPos[0])
		IniWrite($IniFileName, $IniSectionSettings, "PosY", $WindowPos[1])
		IniWrite($IniFileName, $IniSectionSettings, "Width", $WindowPos[2])
		IniWrite($IniFileName, $IniSectionSettings, "Height", $WindowPos[3])
	EndIf
	IniWrite($IniFileName, $IniSectionSettings, "SelectedEEPVersion", $EEPVersionAkt)
	For $v = 1 To $EEPVersionsCount
		Local $Paths = $ResourcenFolder_Paths[$v]
		Local $Descriptions = $ResourcenFolder_Descriptions[$v]
		Local $Data[3 + UBound($Descriptions) * 2][2]
		$Data[1][0] = "Name"
		$Data[1][1] = '"' & IniRead($IniFileName, $IniSectionsVersions & $EEPVersions_Nr[$v], "Name", "") & '"'
		$Data[2][0] = "RegPath"
		$Data[2][1] = '"' & IniRead($IniFileName, $IniSectionsVersions & $EEPVersions_Nr[$v], "RegPath", "") & '"'
		$Data[3][0] = "exeName"
		$Data[3][1] = '"' & IniRead($IniFileName, $IniSectionsVersions & $EEPVersions_Nr[$v], "exeName", "") & '"'
		$Data[4][0] = "Count"
		$Data[4][1] = UBound($Descriptions) - 1
		For $i = 1 To UBound($Descriptions) - 1
			$Data[$i * 2 + 3][0] = "Path" & $i
			$Data[$i * 2 + 3][1] = '"' & $Paths[$i] & '"'
			$Data[$i * 2 + 4][0] = "Description" & $i
			$Data[$i * 2 + 4][1] = '"' & $Descriptions[$i] & '"'

		Next
		IniWriteSection($IniFileName, $IniSectionsVersions & $EEPVersions_Nr[$v], $Data)
	Next
	If $Restart = True Then ;Restart if flag is set
		Exit Run(@ScriptFullPath)
	EndIf
	Exit
EndFunc   ;==>Close

Func CMDdir($path, $Folder = "")
	If $Folder = "" Then
		Local $result = StringRegExp($path, "(.*)[/\\]([^/\\]+?)[/\\]?\z", 1)
		$path = $result[0]
		$Folder = $result[1]
	EndIf
	Local $cmd = Run(@ComSpec & " /c dir /AD /N", $path, @SW_HIDE, $STDOUT_CHILD)
	Local $line = ""
	While 1
		$line &= StdoutRead($cmd)
		If @error Then ExitLoop
		Sleep(1)
	WEnd
	Local $out = StringRegExp($line, "<(.*?)>\s*?" & $Folder & "(\s*?\[(.*?)\])?", 1)
	SetError(@error)
	Return $out
EndFunc   ;==>CMDdir

Func DirIsLink($path, $Folder = "")
	Local $out = CMDdir($path, $Folder)
;~ 	_ArrayDisplay($out)
	If @error = 0 Then
		If $out[0] = "DIR" Then
			Return False
		Else
			Return True
		EndIf
	EndIf
	Return False
EndFunc   ;==>DirIsLink

Func LinkGetRealDir($path, $Folder = "")
	If FileExists($path & "\" & $TxtFilesName) Then
		Local $DirPath = FileRead($path & "\" & $TxtFilesName)
;~ 		MsgBox(0,$ToolName,$DirPath)
		Return $DirPath
	EndIf
	Local $out = CMDdir($path)
	If @error = 0 Then
		If UBound($out) > 1 Then
			If $out[0] <> "DIR" Then
;~ 				MsgBox(0, "Der Link zeigt nach", $out[2])
				Local $offset=StringInStr($out[2],":")
				Return StringTrimLeft($out[2],$offset-2)
			EndIf
		EndIf
	EndIf
	Return ""
EndFunc   ;==>LinkGetRealDir

Func WM_NOTIFY($hWnd, $uMsg, $wParam, $lParam)
	Local $hWndFrom, $iIDFrom, $iCode, $tNMHDR, $hWndListView, $tInfo

	$tNMHDR = DllStructCreate($tagNMHDR, $lParam)
	$hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
	$iIDFrom = DllStructGetData($tNMHDR, "IDFrom")
	$iCode = DllStructGetData($tNMHDR, "Code")

	Switch $hWndFrom
		Case GUICtrlGetHandle($ListViews_ResourcenFolders[$EEPVersionAkt])
			Switch $iCode
				Case $NM_CLICK
					$tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
					Local $item = DllStructGetData($tInfo, "Index")
					Local $subItem = DllStructGetData($tInfo, "SubItem")
					Switch $subItem
						Case 0
							SwitchRegistryTo($item)
							_GUICtrlListView_SetItemSelected($ListView_ResourcenFolders, $item, False)
						Case 1
							SwitchLinkTo($item)
							_GUICtrlListView_SetItemSelected($ListView_ResourcenFolders, $item, False)
					EndSwitch
				Case $LVN_HOTTRACK
					$tInfo = DllStructCreate($tagNMLISTVIEW, $lParam)
					Local $item = DllStructGetData($tInfo, "Item")
					Local $subItem = DllStructGetData($tInfo, "SubItem")
					If $LastHoveredItem[0] <> $item Or $LastHoveredItem[1] <> $subItem Then
						If $LastHoveredItem[0]>=0 Then
							If $LastHoveredItem[1]=1 Then _GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $LastHoveredItem[0], $LinkColumnIcon[$LastHoveredItem[0]], 1)
							_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $LastHoveredItem[0], $RegColumnIcon[$LastHoveredItem[0]], 0)
							_GUICtrlListView_RedrawItems($ListView_ResourcenFolders, $LastHoveredItem[0], $LastHoveredItem[0])
						EndIf
						#cs
						Switch $LastHoveredItem[1]
							Case 0
								If $LastHoveredItem[0] <> $CurrResFolderReg[$EEPVersionAkt] Then
									_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $LastHoveredItem[0], $IconKeins, 0)
									_GUICtrlListView_RedrawItems($ListView_ResourcenFolders, $LastHoveredItem[0], $LastHoveredItem[0])
								EndIf
							Case 1
								If $LastHoveredItem[0] <> $CurrResFolderLink[$EEPVersionAkt] Then
									_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $LastHoveredItem[0], $IconKeins, 1)
								EndIf
								If $LastHoveredItem[0] <> $CurrResFolderReg[$EEPVersionAkt] Then
									_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $LastHoveredItem[0], $IconKeins, 0)
									_GUICtrlListView_RedrawItems($ListView_ResourcenFolders, $LastHoveredItem[0], $LastHoveredItem[0])
								EndIf
						EndSwitch
						#ce
						Switch $subItem
							Case 0
								If _GUICtrlListView_GetItemImage($ListView_ResourcenFolders,$item,0)<>$IconRegAktiv Then
									If $EEPVersions_HasRegResBase[$EEPVersionAkt] Then
										_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $item, $IconRegHover, 0)
									Else
										_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $item, $IconRegError, 0)
									EndIf
								EndIf
;~ 								If $item <> $CurrResFolderReg[$EEPVersionAkt] Then _GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $item, $IconRegHover, 0)
							Case 1
								If _GUICtrlListView_GetItemImage($ListView_ResourcenFolders,$item,0)<>$IconRegLinkAktiv And $EEPVersions_HasRegResBase[$EEPVersionAkt] Then
									_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $item, $IconRegLinkHover, 0)
								EndIf
								If _GUICtrlListView_GetItemImage($ListView_ResourcenFolders,$item,1)<>$IconLinkAktiv Then
									_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $item, $IconLinkHover, 1)
								EndIf
;~ 								If $item <> $CurrResFolderLink[$EEPVersionAkt] Then _GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $item, $IconRegLinkHover, 0)
;~ 								If $item <> $CurrResFolderLink[$EEPVersionAkt] Then _GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $item, $IconLinkHover, 1)
						EndSwitch
						$LastHoveredItem[0] = $item
						$LastHoveredItem[1] = $subItem
					EndIf
				Case $LVN_ITEMACTIVATE ;Doppelklick
					;MsgBox(0, $ToolName, "Activated")
			EndSwitch

		Case GUICtrlGetHandle($TabView_EEPVersions)
			Switch $iCode
				Case - 551
;~ 					;Local $ID=GUICtrlRead($TabView_EEPVersions)
;~ 					;MsgBox(0,$ToolName,"Mit den Tabs ist was los: "&$ID&$iCode)
					SetVersion(GUICtrlRead($TabView_EEPVersions) + 1)
				Case - 552
					If BitAND(WinGetState($ResourcenFolder_GUI), 2) = 2 Then Return 1
			EndSwitch

	EndSwitch
	Switch $uMsg
		Case $WM_GETMINMAXINFO ; Die Fenstergröße abfragen, minimale Größe fürs verkleinern setzen
			Local $MinMax = DllStructCreate("int ptReserved[2]; int ptMaxSize[2]; int ptMaxPosition[2]; int ptMinTrackSize[2]; int ptMaxTrackSize[2];", $lParam) ; DLLStruct auf den Pointer erstellen, zum bearbeiten der Werte
			DllStructSetData($MinMax, 4, 500, 1) ; Minimal 500 Pixel breit
			DllStructSetData($MinMax, 4, 180, 2) ; Minimal 180 Pixel hoch
	EndSwitch

EndFunc   ;==>WM_NOTIFY

Func SetVersion($number = -1)
;~ 	MsgBox(0,"EEP-Version Nr.",$number)
	$EEPVersionAkt = $number
;~ 	$IniSectionResourcenFolders = $IniSectionsVersions & $EEPVersionAkt
	$ListView_ResourcenFolders = $ListViews_ResourcenFolders[$number]
	$EEPRegPath=$EEPVersions_RegPath[$number]
	$RegColumnIcon=$RegColumnsIcon[$number]
	$LinkColumnIcon=$LinkColumnsIcon[$number]
EndFunc   ;==>SetVersion

#EndRegion Functions

While 1
	Sleep(1000)
WEnd
#cs
[Settings]
EEPVersionsCount=3
PosX=710
PosY=450
Width=763
Height=234
[EEPVersion1]
Name="bis EEP6"
RegPath="HKEY_CLASSES_ROOT\Software\Software Untergrund\EEXP"
exeName="EEP.exe"
Count=0
[EEPVersion2]
Name=EEP7
RegPath="HKEY_LOCAL_MACHINE\SOFTWARE\Trend\EEP 7.00\EEXP"
exeName="EEP7.exe"
[EEPVersion3]
Name="EEP8"
RegPath="HKEY_LOCAL_MACHINE\SOFTWARE\Trend\EEP 8.00\EEXP"
exeName="EEP8.exe"
Count=0
#ce
