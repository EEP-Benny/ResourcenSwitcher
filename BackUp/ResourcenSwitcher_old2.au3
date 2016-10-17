#AutoIt3Wrapper_Res_Icon_Add=folder_add.ico
#AutoIt3Wrapper_Res_Icon_Add=folder_edit.ico
#AutoIt3Wrapper_Res_Icon_Add=folder_delete.ico
#AutoIt3Wrapper_Res_Icon_Add=database.ico
#AutoIt3Wrapper_Res_Icon_Add=database_go.ico
#AutoIt3Wrapper_Res_Icon_Add=link.ico
#AutoIt3Wrapper_Res_Icon_Add=link_go.ico
#AutoIt3Wrapper_Icon=ResSwitchmultiple.ico

;Todo: 	Aktuellen Ordner herausfinden

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
#EndRegion - Include Parameters

#Region - Options
Opt("GUIOnEventMode", 1)
Opt("GUICoordMode", 0)
Opt("TrayIconHide", 1)
Opt("GUICloseOnEsc", 0)
Opt("MustDeclareVars", 1)
#EndRegion - Options
Global $ToolName = "ResourcenSwitcher"
Global $EEPRegPath = "HKEY_CLASSES_ROOT\Software\Software Untergrund\EEXP"

Global $IniFileName = @ScriptDir & "\" & $ToolName & ".ini"
Global $IniSectionSettings = "Settings"
Global $IniSectionsVersions = "EEPVersion"

Global $TxtFilesName=$ToolName&".txt"

Global $EEPVersionsCount = IniRead($IniFileName, $IniSectionSettings, "EEPVersionsCount", 0)
If ($EEPVersionsCount < 1) Then
	$EEPVersionsCount = 3
	IniWrite($IniFileName, $IniSectionSettings, "EEPVersionsCount", 3)
	Local $SectionEEP6[3][2] = [["Name", '"bis EEP6"'], _
			["RegPath", '"HKEY_CLASSES_ROOT\Software\Software Untergrund\EEXP"'], _
			["exeName", '"EEP.exe"']]
	IniWriteSection($IniFileName, $IniSectionsVersions & "1", $SectionEEP6, 0)
	Local $SectionEEP7[3][2] = [["Name", "EEP7"], _
			["RegPath", '"HKEY_LOCAL_MACHINE\SOFTWARE\Trend\EEP 7.00\EEXP"'], _
			["exeName", '"EEP7.exe"']]
	IniWriteSection($IniFileName, $IniSectionsVersions & "2", $SectionEEP7, 0)
	Local $SectionEEP8[3][2] = [["Name", "EEP8"], _
			["RegPath", '"HKEY_LOCAL_MACHINE\SOFTWARE\Trend\EEP 8.00\EEXP"'], _
			["exeName", '"EEP8.exe"']]
	IniWriteSection($IniFileName, $IniSectionsVersions & "3", $SectionEEP8, 0)
EndIf

Global $EEPVersions_Name[1]
Global $EEPVersions_RegPath[1]
Global $EEPVersions_Nr[1]
For $i = 1 To $EEPVersionsCount
	RegRead(IniRead($IniFileName,$IniSectionsVersions&$i,"RegPath",""),"Directory")
	If @error Then
;~ 		MsgBox(0,$ToolName,IniRead($IniFileName,$IniSectionsVersions&$i,"Name","???")&" gibt es nicht")
	Else
		_ArrayAdd($EEPVersions_Name, IniRead($IniFileName, $IniSectionsVersions & $i, "Name", ""))
		_ArrayAdd($EEPVersions_RegPath, IniRead($IniFileName, $IniSectionsVersions & $i, "RegPath", ""))
		_ArrayAdd($EEPVersions_Nr, $i)
	EndIf
Next
$EEPVersionsCount = UBound($EEPVersions_Name) - 1

Global $EEPVersionAkt = 1

Global $IniSectionResourcenFolders = $IniSectionsVersions & $EEPVersionAkt

Global $ResourcenFolder_Descriptions[1]
Global $ResourcenFolder_Paths[1]

Global $EntryToEdit = 0
Global $CurrResFolderLink[$EEPVersionsCount+1]
Global $CurrResFolderReg[$EEPVersionsCount+1]

;Unterschied zwischen Fenstergröße und Clientsize rausfinden
Local $TestGUI = GUICreate("Test-GUI", 500, 300, Default, Default, BitOR($WS_SIZEBOX, $WS_MINIMIZEBOX))
Local $Size = WinGetClientSize($TestGUI)
Global $ClientDiff[2] = [500 - $Size[0], 300 - $Size[1]]
;~ MsgBox(0,"Größenunterschied",$ClientDiff[0]&"|"&$ClientDiff[1])
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
_GUIImageList_AddIcon($ImageList, @ScriptFullPath, 7);0 - Reg aktiv
_GUIImageList_AddIcon($ImageList, @ScriptFullPath, 8);1 - Reg Hover
_GUIImageList_AddIcon($ImageList, @ScriptFullPath, 9);2 - Link aktiv
_GUIImageList_AddIcon($ImageList, @ScriptFullPath, 10);3 - Link Hover

;ImageList für die EEP-Icons
Global $ImageList_EEPIcons = _GUIImageList_Create(16, 16, 4, 1)

;Tabs zum Auswählen der EEP-Version
Global $TabView_EEPVersions = GUICtrlCreateTab(10, 10, 265, 130)
GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)
GUISetOnEvent(-1, "Close")
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
	_ArrayAdd($ListViews_ResourcenFolders, GUICtrlCreateListView("E|M|Beschreibung|Pfad", 20, 40, 210, 85, BitOR($GUI_SS_DEFAULT_LISTVIEW, $LVS_NOSORTHEADER)))
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

GUICtrlCreateTabItem("")

;Label Hilfe
Global $Label_Help = GUICtrlCreateLabel("Hilfe", 0, -92, 25, 15, $GUI_SS_DEFAULT_LABEL + $SS_RIGHT)
GUICtrlSetFont(-1, 8.5, 400, 4)
GUICtrlSetColor(-1, 0x0000FF)
GUICtrlSetCursor(-1, 0)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT + $GUI_DOCKWIDTH)
GUICtrlSetOnEvent($Label_Help, "ShowHelp")
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
ResizeWindow()
SetVersion(1)
GUISetState(@SW_SHOW, $GUI)
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
	If $Index >= 0 Then
		Local $path = _GUICtrlListView_GetItemText($ListView_ResourcenFolders, $Index, 3)
	Else
		Local $path = RegRead($EEPRegPath, "Directory") & "\Resourcen"
		$Index = $CurrResFolderLink[$EEPVersionAkt]
	EndIf
	If RegRead($EEPRegPath, "ResBase")<>$path Then
		RegWrite($EEPRegPath, "ResBase", "REG_SZ", $path)
		If @error Then
			MsgBox(48, $ToolName, "Bitte starte dieses Programm mit Schreibrechten in der Registry, z.B. als Administrator." & @CRLF & _
					"Die Zugriffsrechte können sich auf folgenden Pfad/Schlüssel beschränken:" & @CRLF & _
					"HKEY_CLASSES_ROOT\Software\Software Untergrund\EEXP")
			DisplayResourcenFolders($EEPVersionAkt)
			Return
		EndIf
	EndIf
	_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $CurrResFolderReg[$EEPVersionAkt], -1, 0)
;~ 	_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $Index, 0, 0)
	$CurrResFolderReg[$EEPVersionAkt] = $Index
 	DisplayResourcenFolders($EEPVersionAkt)
;~ 	;MsgBox(64,$ToolName,"Resourcen-Ordner wurde gewechselt nach"&@CRLF&$ItemTextArray[4],Default,$GUI)
EndFunc   ;==>SwitchRegistryTo

Func SwitchLinkTo($Index)
	Local $path = _GUICtrlListView_GetItemText($ListView_ResourcenFolders, $Index, 3)
	Local $homeRes = RegRead($EEPRegPath, "Directory") & "\Resourcen"
	If FileExists($homeRes) And Not DirIsLink(RegRead($EEPRegPath, "Directory"), "Resourcen") Then
		MsgBox(16, $ToolName, $homeRes & " ist noch ein vollwertiger Resourcen-Ordner." & @CRLF & @CRLF & "Bitte benenne diesen Ordner um (z.B. in Resourcen_Voll), " & @CRLF & "damit an seiner Stelle eine Verknüpfung erstellt werden kann.")
		Return
	EndIf
	If FileCreateNTFSLink($path, $homeRes, 1) Then
		_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $CurrResFolderLink[$EEPVersionAkt], -1, 1)
;~ 		_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $Index, 2, 1)
		$CurrResFolderLink[$EEPVersionAkt] = $Index
		If LinkGetRealDir($homeRes)<>$path Then
			If FileExists($homeRes&"\"&$TxtFilesName) Then FileSetAttrib($homeRes&"\"&$TxtFilesName,"-H")
			Local $file=FileOpen($homeRes&"\"&$TxtFilesName,2)
			FileWrite($homeRes&"\"&$TxtFilesName,$path)
			FileClose($file)
			FileSetAttrib($homeRes&"\"&$TxtFilesName,"+H")
		EndIf
	EndIf
	SwitchRegistryTo(-1)
EndFunc   ;==>SwitchLinkTo

Func ShowHelp()
	Local $PDFtype = RegRead("HKEY_CLASSES_ROOT\.pdf", "")
	Local $PDFapplication = RegRead("HKEY_CLASSES_ROOT\" & $PDFtype & "\Shell\Open\Command", "")
	Run(StringReplace($PDFapplication, "%1", @ScriptDir & "\Beschreibung.pdf"))
EndFunc   ;==>ShowHelp

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
	Global $EntryToEdit = $SelectedIndex + 1
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
	If $EntryToEdit > UBound($ResourcenFolder_Descriptions[$EEPVersionAkt]) - 1 Then
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
		Local $Paths[1]
		Local $Descriptions[1]
		For $i = 1 To IniRead($IniFileName, $IniSectionsVersions & $v, "Count", 0)
			Local $path = IniRead($IniFileName, $IniSectionsVersions & $v, "Path" & $i, "Konnte Pfad nicht lesen")
			Local $Description = IniRead($IniFileName, $IniSectionsVersions & $v, "Description" & $i, "Konnte Beschreibung nicht lesen")
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
	GUISetState(@SW_LOCK,$GUI)
;~ 	;_ArrayDisplay($ResourcenFolder_Paths[$v])
	_GUICtrlListView_DeleteAllItems($ListViews_ResourcenFolders[$v])
	Local $homeRes=RegRead($EEPVersions_RegPath[$v],"Directory")&"\Resourcen"
	Local $ResBase = RegRead($EEPVersions_RegPath[$v], "ResBase")
	Local $LinkFolder=LinkGetRealDir($homeRes)
;~ 	MsgBox(0,$ResBase,$LinkFolder)
	$CurrResFolderLink[$v] = _ArraySearch($ResourcenFolder_Paths[$v], $LinkFolder) - 1
	If $ResBase=$homeRes Then
		$CurrResFolderReg[$v]=$CurrResFolderLink[$v]
	Else
		$CurrResFolderReg[$v] = _ArraySearch($ResourcenFolder_Paths[$v], $ResBase) - 1
	EndIf
;~ 	MsgBox(0,$ActResFolderIndex,$ActLinkIndex)
	For $i = 1 To UBound($ResourcenFolder_Descriptions[$v]) - 1
		Local $Descriptions = $ResourcenFolder_Descriptions[$v]
		Local $Paths = $ResourcenFolder_Paths[$v]
		GUICtrlCreateListViewItem("||" & $Descriptions[$i] & "|" & $Paths[$i], $ListViews_ResourcenFolders[$v])
	Next
;~ 	MsgBox(0,$CurrResFolderLink[$v],$CurrResFolderReg[$v])
	If ($ResBase=$homeRes Or $ResBase=$LinkFolder) And $CurrResFolderLink[$v]<-1 Then
		$CurrResFolderReg[$v] = _GUICtrlListView_GetItemCount($ListViews_ResourcenFolders[$v])
		$CurrResFolderLink[$v]=$CurrResFolderReg[$v]
		GUICtrlCreateListViewItem("||Aktueller Resourcen-Ordner|" & $LinkFolder, $ListViews_ResourcenFolders[$v])
		GUICtrlSetColor(-1,0xFF0000)
		GUICtrlSetBkColor(-1,0xFFFF00)
	Else
		If $CurrResFolderReg[$v] < -1 Then
			$CurrResFolderReg[$v] = _GUICtrlListView_GetItemCount($ListViews_ResourcenFolders[$v])
			GUICtrlCreateListViewItem("||Aktueller Resourcen-Ordner in der Registry|" & $ResBase, $ListViews_ResourcenFolders[$v])
		EndIf
		If $CurrResFolderLink[$v] < -1 Then
			$CurrResFolderLink[$v] = _GUICtrlListView_GetItemCount($ListViews_ResourcenFolders[$v])
			GUICtrlCreateListViewItem("||Aktuell verlinkter Resourcen-Ordner|" & $LinkFolder, $ListViews_ResourcenFolders[$v])
		EndIf
	EndIf
	_GUICtrlListView_SetItemImage($ListViews_ResourcenFolders[$v], $CurrResFolderReg[$v], 0)
	_GUICtrlListView_SetItemImage($ListViews_ResourcenFolders[$v], $CurrResFolderLink[$v], 2,1)
	_GUICtrlListView_SetColumnWidth($ListViews_ResourcenFolders[$v], 0, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($ListViews_ResourcenFolders[$v], 1, $LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_SetColumnWidth($ListViews_ResourcenFolders[$v], 2, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($ListViews_ResourcenFolders[$v], 3, $LVSCW_AUTOSIZE)
	GUISetState(@SW_UNLOCK,$GUI)
EndFunc   ;==>DisplayResourcenFolders

Func ResizeWindow()
;~ 	Local $Width = 0;_GUICtrlListView_GetColumnWidth($ListView_ResourcenFolders,0)+_GUICtrlListView_GetColumnWidth($ListView_ResourcenFolders,1)+30
;~ 	Local $Height = 0;_GUICtrlListView_GetItemCount($ListView_ResourcenFolders)*17+150
	Local $Width = IniRead($IniFileName, $IniSectionSettings, "Width", 0)
	Local $Height = IniRead($IniFileName, $IniSectionSettings, "Height", 0)
	If $Height < 180 Then $Height = 180
	If $Width < 500 Then $Width = 500
	Local $Left = IniRead($IniFileName, $IniSectionSettings, "PosX", (@DesktopWidth - $Width) / 2)
	Local $Top = IniRead($IniFileName, $IniSectionSettings, "Top", (@DesktopHeight - $Height) / 2)
	WinMove($GUI, "", $Left, $Top, $Width, $Height)
EndFunc   ;==>ResizeWindow

Func Cancel()
	GUISetState(@SW_HIDE, @GUI_WinHandle)
EndFunc   ;==>Cancel

Func Close()
	Local $WindowPos = WinGetPos($ToolName)
	IniWrite($IniFileName, $IniSectionSettings, "PosX", $WindowPos[0])
	IniWrite($IniFileName, $IniSectionSettings, "PosY", $WindowPos[1])
	IniWrite($IniFileName, $IniSectionSettings, "Width", $WindowPos[2])
	IniWrite($IniFileName, $IniSectionSettings, "Height", $WindowPos[3])
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
	Exit
EndFunc   ;==>Close

Func CMDdir($path,$Folder="")
	If $Folder = "" Then
		Local $result = StringRegExp($path, "(.*)[/\\]([^/\\]+?)[/\\]?\z", 1)
		$path = $result[0]
		$Folder = $result[1]
	EndIf
	Local $cmd = Run(@ComSpec & " /c dir /AD /N", $path, @SW_HIDE, $STDOUT_CHILD)
	Local $line = "";
	While 1
		$line &= StdoutRead($cmd)
		If @error Then ExitLoop
	WEnd
	Local $out= StringRegExp($line, "<(.*?)>\s*?" & $Folder&"(\s*?\[(.*?)\])?", 1)
	SetError(@error)
	Return $out
EndFunc

Func DirIsLink($path, $Folder = "")
	Local $out = CMDdir($path,$Folder)
;~ 	_ArrayDisplay($out)
	If @error = 0 Then
		If $out[0] = "DIR" Then
			Return False
		Else
			Return True
		EndIf
	EndIf
EndFunc   ;==>DirIsLink

Func LinkGetRealDir($path, $Folder="")
	If FileExists($path&"\"&$TxtFilesName) Then 
		Local $DirPath=FileRead($path&"\"&$TxtFilesName)
;~ 		MsgBox(0,$ToolName,$DirPath)
		Return $DirPath
	EndIf
	Local $out=CMDdir($path)
	If @error = 0 Then
		If UBound($out)>1 Then
			If $out[0] <> "DIR" Then
				MsgBox(0,"Der Link zeigt nach",$out[1])
				Return $out[1]
			EndIf
		EndIf
	EndIf
	Return ""
EndFunc

Global $LastHoveredItem[2] = [-1, -1]
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
						Switch $LastHoveredItem[1]
							Case 0
								If $LastHoveredItem[0] <> $CurrResFolderReg[$EEPVersionAkt] Then
									_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $LastHoveredItem[0], -1, 0)
									_GUICtrlListView_RedrawItems($ListView_ResourcenFolders, $LastHoveredItem[0], $LastHoveredItem[0])
								EndIf
							Case 1
								If $LastHoveredItem[0] <> $CurrResFolderLink[$EEPVersionAkt] Then
									_GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $LastHoveredItem[0], -1, 1)
								EndIf
						EndSwitch
						Switch $subItem
							Case 0
								If $item <> $CurrResFolderReg[$EEPVersionAkt] Then _GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $item, 1, 0)
							Case 1
								If $item <> $CurrResFolderLink[$EEPVersionAkt] Then _GUICtrlListView_SetItemImage($ListView_ResourcenFolders, $item, 3, 1)
						EndSwitch
						$LastHoveredItem[0] = $item
						$LastHoveredItem[1] = $subItem
					EndIf
				Case $LVN_ITEMACTIVATE
					MsgBox(0, $ToolName, "Activated")
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
;~ 	;MsgBox(0,"EEP-Version Nr.",$number)
	$EEPVersionAkt = $number
	$IniSectionResourcenFolders = $IniSectionsVersions & $EEPVersionAkt
	$ListView_ResourcenFolders = $ListViews_ResourcenFolders[$number]
EndFunc   ;==>SetVersion

#EndRegion Functions

While 1
	Sleep(1000)
WEnd