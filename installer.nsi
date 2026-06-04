; PageCapture NSIS Installer Script
; Requires NSIS 3.x — compile with: makensis installer.nsi

Unicode True

!define APP_NAME     "PageCapture"
!define APP_VERSION  "1.0"
!define APP_EXE      "PageCapture.exe"
!define INSTALL_DIR  "$PROGRAMFILES64\PageCapture"

Name         "${APP_NAME} ${APP_VERSION}"
OutFile      "PageCapture_Setup.exe"
InstallDir   "${INSTALL_DIR}"
InstallDirRegKey HKCU "Software\PageCapture" "InstallDir"
RequestExecutionLevel admin
SetCompressor lzma

;--- Pages ---
!include "MUI2.nsh"
!define MUI_ABORTWARNING
!define MUI_ICON "assets\icon.ico"
!define MUI_UNICON "assets\icon.ico"
!define MUI_WELCOMEPAGE_TITLE "Welcome to PageCapture Setup"
!define MUI_WELCOMEPAGE_TEXT "PageCapture lets you capture text, links, and images from any browser or chat window.$\n$\nPress Win+Shift+C or click the system tray icon to get started."

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

;--- Install ---
Section "Install"
  SetOutPath "$INSTDIR"

  ; Copy all built files from PyInstaller output
  File /r "dist\PageCapture\*.*"

  ; Write uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  ; Desktop shortcut
  CreateShortcut "$DESKTOP\PageCapture.lnk" \
    "$INSTDIR\${APP_EXE}" "" \
    "$INSTDIR\${APP_EXE}" 0

  ; Start Menu shortcut
  CreateDirectory "$SMPROGRAMS\PageCapture"
  CreateShortcut "$SMPROGRAMS\PageCapture\PageCapture.lnk" \
    "$INSTDIR\${APP_EXE}" "" \
    "$INSTDIR\${APP_EXE}" 0
  CreateShortcut "$SMPROGRAMS\PageCapture\Uninstall.lnk" \
    "$INSTDIR\Uninstall.exe"

  ; Add to Windows startup (optional — runs in tray on login)
  WriteRegStr HKCU \
    "Software\Microsoft\Windows\CurrentVersion\Run" \
    "PageCapture" \
    '"$INSTDIR\${APP_EXE}"'

  ; Registry for uninstaller
  WriteRegStr HKCU "Software\PageCapture" "InstallDir" "$INSTDIR"
  WriteRegStr HKLM \
    "Software\Microsoft\Windows\CurrentVersion\Uninstall\PageCapture" \
    "DisplayName" "PageCapture"
  WriteRegStr HKLM \
    "Software\Microsoft\Windows\CurrentVersion\Uninstall\PageCapture" \
    "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegStr HKLM \
    "Software\Microsoft\Windows\CurrentVersion\Uninstall\PageCapture" \
    "DisplayVersion" "${APP_VERSION}"

SectionEnd

;--- Uninstall ---
Section "Uninstall"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir /r "$INSTDIR"
  Delete "$DESKTOP\PageCapture.lnk"
  RMDir /r "$SMPROGRAMS\PageCapture"
  DeleteRegKey HKCU "Software\PageCapture"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PageCapture"
  DeleteRegValue HKCU \
    "Software\Microsoft\Windows\CurrentVersion\Run" \
    "PageCapture"
SectionEnd
