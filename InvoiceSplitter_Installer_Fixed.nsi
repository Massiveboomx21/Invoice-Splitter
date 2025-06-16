; Invoice Splitter NSIS Installer
; Version 1.3

!include "MUI2.nsh"
!include "FileFunc.nsh"

; General Configuration
Name "Invoice Splitter 1.3"
OutFile "InvoiceSplitter_Setup_v1.3.exe"
InstallDir "$PROGRAMFILES\InvoiceSplitter"
InstallDirRegKey HKLM "Software\InvoiceSplitter" "Install_Dir"
RequestExecutionLevel admin

; Interface Settings
!define MUI_ABORTWARNING

; Icon settings - check if exists
!if /FileExists "dist\InvoiceSplitter\resources\ico\arrows_16382055.ico"
  !define MUI_ICON "dist\InvoiceSplitter\resources\ico\arrows_16382055.ico"
  !define MUI_UNICON "dist\InvoiceSplitter\resources\ico\arrows_16382055.ico"
!else if /FileExists "resources\ico\arrows_fixed.ico"
  !define MUI_ICON "resources\ico\arrows_fixed.ico"
  !define MUI_UNICON "resources\ico\arrows_fixed.ico"
!endif

; Language Selection
!define MUI_LANGDLL_REGISTRY_ROOT "HKCU"
!define MUI_LANGDLL_REGISTRY_KEY "Software\InvoiceSplitter"
!define MUI_LANGDLL_REGISTRY_VALUENAME "Installer Language"

; Pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "LICENSE.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES

; Finish page
!define MUI_FINISHPAGE_RUN "$INSTDIR\InvoiceSplitter.exe"
!define MUI_FINISHPAGE_RUN_TEXT "Εκκίνηση του Invoice Splitter"
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; Languages
!insertmacro MUI_LANGUAGE "Greek"
!insertmacro MUI_LANGUAGE "English"

; Version Information
VIProductVersion "1.3.0.0"
VIAddVersionKey "ProductName" "Invoice Splitter"
VIAddVersionKey "CompanyName" "Your Company"
VIAddVersionKey "LegalCopyright" "Copyright 2025"
VIAddVersionKey "FileDescription" "Invoice Splitter Installer"
VIAddVersionKey "FileVersion" "1.3.0"

; Main Section
Section "Invoice Splitter" SEC01
  
  SectionIn RO
  
  SetOutPath "$INSTDIR"
  
  ; Copy all files from dist/InvoiceSplitter
  File /r "dist\InvoiceSplitter\*.*"
  
  ; Create logs directory
  CreateDirectory "$INSTDIR\logs"
  
  ; Write registry keys
  WriteRegStr HKLM "Software\InvoiceSplitter" "Install_Dir" "$INSTDIR"
  
  ; Create uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  
  ; Add to Add/Remove Programs
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\InvoiceSplitter" "DisplayName" "Invoice Splitter 3"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\InvoiceSplitter" "UninstallString" "$INSTDIR\Uninstall.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\InvoiceSplitter" "DisplayIcon" "$INSTDIR\InvoiceSplitter.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\InvoiceSplitter" "Publisher" "Your Company"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\InvoiceSplitter" "DisplayVersion" "1.3.0"
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\InvoiceSplitter" "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\InvoiceSplitter" "NoRepair" 1
  
  ; Calculate size
  ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
  IntFmt $0 "0x%08X" $0
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\InvoiceSplitter" "EstimatedSize" "$0"

SectionEnd

; Start Menu Shortcuts
Section "Start Menu Shortcuts" SEC02

  CreateDirectory "$SMPROGRAMS\Invoice Splitter"
  CreateShortcut "$SMPROGRAMS\Invoice Splitter\Invoice Splitter.lnk" "$INSTDIR\InvoiceSplitter.exe" "" "$INSTDIR\InvoiceSplitter.exe" 0
  CreateShortcut "$SMPROGRAMS\Invoice Splitter\Uninstall.lnk" "$INSTDIR\Uninstall.exe" "" "$INSTDIR\Uninstall.exe" 0

SectionEnd

; Desktop Shortcut
Section "Desktop Shortcut" SEC03

  CreateShortcut "$DESKTOP\Invoice Splitter.lnk" "$INSTDIR\InvoiceSplitter.exe" "" "$INSTDIR\InvoiceSplitter.exe" 0

SectionEnd

; Descriptions
LangString DESC_SEC01 ${LANG_GREEK} "Τα βασικά αρχεία της εφαρμογής"
LangString DESC_SEC02 ${LANG_GREEK} "Συντομεύσεις στο Start Menu"
LangString DESC_SEC03 ${LANG_GREEK} "Συντόμευση στην επιφάνεια εργασίας"

LangString DESC_SEC01 ${LANG_ENGLISH} "Core application files"
LangString DESC_SEC02 ${LANG_ENGLISH} "Start Menu shortcuts"
LangString DESC_SEC03 ${LANG_ENGLISH} "Desktop shortcut"

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC01} $(DESC_SEC01)
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC02} $(DESC_SEC02)
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC03} $(DESC_SEC03)
!insertmacro MUI_FUNCTION_DESCRIPTION_END

; Uninstaller Section
Section "Uninstall"
  
  ; Delete files and directories
  Delete "$INSTDIR\InvoiceSplitter.exe"
  Delete "$INSTDIR\Uninstall.exe"
  Delete "$INSTDIR\version.txt"
  Delete "$INSTDIR\README.txt"
  Delete "$INSTDIR\LICENSE.txt"
  
  RMDir /r "$INSTDIR\resources"
  RMDir /r "$INSTDIR\modules"
  RMDir /r "$INSTDIR\ui"
  RMDir /r "$INSTDIR\logs"
  RMDir /r "$INSTDIR\_internal"
  
  RMDir "$INSTDIR"
  
  ; Delete shortcuts
  Delete "$DESKTOP\Invoice Splitter.lnk"
  Delete "$SMPROGRAMS\Invoice Splitter\*.*"
  RMDir "$SMPROGRAMS\Invoice Splitter"
  
  ; Delete registry keys
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\InvoiceSplitter"
  DeleteRegKey HKLM "Software\InvoiceSplitter"

SectionEnd

; Functions
Function .onInit
  !insertmacro MUI_LANGDLL_DISPLAY
FunctionEnd