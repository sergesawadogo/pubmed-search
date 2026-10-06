; Installateur Windows 10 / 11 de PubMed Search (NSIS 3)
; Compilation (Linux ou Windows) : makensis -DVERSION=1.3.0 -DSRCDIR=<dossier PubMedSearch> installer.nsi
Unicode true
Target amd64-unicode
!include "MUI2.nsh"
!include "WinVer.nsh"
!include "x64.nsh"

!ifndef VERSION
  !define VERSION "1.3.0"
!endif
!ifndef SRCDIR
  !define SRCDIR "..\..\dist\PubMedSearch"
!endif
!ifndef OUTFILE
  !define OUTFILE "..\..\dist\PubMedSearch-Setup-${VERSION}.exe"
!endif
!define APPNAME "PubMed Search"
!define EXENAME "PubMedSearch.exe"
!define UNINSTKEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\PubMedSearch"

Name "${APPNAME} ${VERSION}"
OutFile "${OUTFILE}"
InstallDir "$PROGRAMFILES64\${APPNAME}"
InstallDirRegKey HKLM "Software\PubMedSearch" "InstallDir"
RequestExecutionLevel admin
SetCompressor /SOLID lzma
SetCompressorDictSize 64
BrandingText "${APPNAME} ${VERSION}"
VIProductVersion "${VERSION}.0"
VIAddVersionKey "ProductName" "${APPNAME}"
VIAddVersionKey "ProductVersion" "${VERSION}"
VIAddVersionKey "FileVersion" "${VERSION}"
VIAddVersionKey "FileDescription" "Installateur de ${APPNAME}"
VIAddVersionKey "CompanyName" "Serge Sawadogo"
VIAddVersionKey "LegalCopyright" "Serge Sawadogo"

!define MUI_ICON "..\..\resources\icon.ico"
!define MUI_UNICON "..\..\resources\icon.ico"
!define MUI_ABORTWARNING
!define MUI_FINISHPAGE_RUN "$INSTDIR\${EXENAME}"
!define MUI_FINISHPAGE_RUN_TEXT "Lancer ${APPNAME}"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "French"

Function .onInit
  ${IfNot} ${RunningX64}
    MessageBox MB_ICONSTOP "${APPNAME} nécessite Windows 10 ou 11 en 64 bits."
    Abort
  ${EndIf}
  ${IfNot} ${AtLeastWin10}
    MessageBox MB_ICONSTOP "${APPNAME} nécessite Windows 10 ou Windows 11."
    Abort
  ${EndIf}
  SetRegView 64
  SetShellVarContext all
FunctionEnd

Section "${APPNAME} (requis)" SecMain
  SectionIn RO
  SetOutPath "$INSTDIR"
  RMDir /r "$INSTDIR\_internal"
  File /r "${SRCDIR}\*.*"
  WriteUninstaller "$INSTDIR\Desinstaller.exe"
  CreateShortcut "$SMPROGRAMS\${APPNAME}.lnk" "$INSTDIR\${EXENAME}" "" "$INSTDIR\${EXENAME}" 0
  WriteRegStr HKLM "Software\PubMedSearch" "InstallDir" "$INSTDIR"
  WriteRegStr HKLM "${UNINSTKEY}" "DisplayName" "${APPNAME}"
  WriteRegStr HKLM "${UNINSTKEY}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKLM "${UNINSTKEY}" "Publisher" "Serge Sawadogo"
  WriteRegStr HKLM "${UNINSTKEY}" "DisplayIcon" "$INSTDIR\${EXENAME}"
  WriteRegStr HKLM "${UNINSTKEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKLM "${UNINSTKEY}" "UninstallString" '"$INSTDIR\Desinstaller.exe"'
  WriteRegDWORD HKLM "${UNINSTKEY}" "NoModify" 1
  WriteRegDWORD HKLM "${UNINSTKEY}" "NoRepair" 1
  ; Liens pubmedsearch:// (extension Firefox « Envoyer vers PubMed Search »)
  WriteRegStr HKLM "Software\Classes\pubmedsearch" "" "URL:PubMed Search"
  WriteRegStr HKLM "Software\Classes\pubmedsearch" "URL Protocol" ""
  WriteRegStr HKLM "Software\Classes\pubmedsearch\DefaultIcon" "" "$INSTDIR\${EXENAME},0"
  WriteRegStr HKLM "Software\Classes\pubmedsearch\shell\open\command" "" '"$INSTDIR\${EXENAME}" "%1"'
SectionEnd

Section "Raccourci sur le Bureau" SecDesktop
  CreateShortcut "$DESKTOP\${APPNAME}.lnk" "$INSTDIR\${EXENAME}" "" "$INSTDIR\${EXENAME}" 0
SectionEnd

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SecMain} "L'application, son Python intégré et le script pubmed_search.py fourni."
  !insertmacro MUI_DESCRIPTION_TEXT ${SecDesktop} "Ajoute une icône PubMed Search sur le Bureau."
!insertmacro MUI_FUNCTION_DESCRIPTION_END

Function un.onInit
  SetRegView 64
  SetShellVarContext all
FunctionEnd

Section "Uninstall"
  Delete "$SMPROGRAMS\${APPNAME}.lnk"
  Delete "$DESKTOP\${APPNAME}.lnk"
  RMDir /r "$INSTDIR\_internal"
  Delete "$INSTDIR\${EXENAME}"
  Delete "$INSTDIR\Desinstaller.exe"
  RMDir "$INSTDIR"
  DeleteRegKey HKLM "${UNINSTKEY}"
  DeleteRegKey HKLM "Software\Classes\pubmedsearch"
  DeleteRegKey HKLM "Software\PubMedSearch"
  ; Les réglages de l'utilisateur (%LOCALAPPDATA%\pubmed-search-gui) sont conservés.
SectionEnd
