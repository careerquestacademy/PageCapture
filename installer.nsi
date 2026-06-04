; PageCapture v3.1 NSIS Installer

Unicode True

!define APP_NAME    "PageCapture"
!define APP_VERSION "3.1"
!define APP_EXE     "PageCapture.exe"

Name         "${APP_NAME} ${APP_VERSION}"
OutFile      "PageCapture_Setup.exe"
InstallDir   "$PROGRAMFILES64\PageCapture"
InstallDirRegKey HKCU "Software\PageCapture" "InstallDir"
RequestExecutionLevel admin
SetCompressor lzma

!include "MUI2.nsh"
!define MUI_ABORTWARNING
!define MUI_ICON "assets\icon.ico"
!define MUI_UNICON "assets\icon.ico"
!define MUI_WELCOMEPAGE_TITLE "Welcome to PageCapture"
!define MUI_WELCOMEPAGE_TEXT "PageCapture lets you save any webpage as a clean, fully editable Word document.$\n$\nIncludes text, images, and clickable links.$\n$\nNote: Windows may warn you this is from an unknown publisher. Click 'More info' then 'Run anyway' — this is normal for new software."

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

Section "Install"
  SetOutPath "$INSTDIR"
  File /r "app\dist\PageCapture\*.*"

  ; Firefox extension folder
  SetOutPath "$INSTDIR\extension"
  File /r "extension\*.*"

  ; Native Messaging host manifest
  File "app\com.careerquest.pagecapture.json"

  WriteUninstaller "$INSTDIR\Uninstall.exe"

  ; Desktop shortcut
  CreateShortcut "$DESKTOP\PageCapture.lnk" \
    "$INSTDIR\${APP_EXE}" "" "$INSTDIR\${APP_EXE}" 0

  ; Start Menu
  CreateDirectory "$SMPROGRAMS\PageCapture"
  CreateShortcut "$SMPROGRAMS\PageCapture\PageCapture.lnk" \
    "$INSTDIR\${APP_EXE}" "" "$INSTDIR\${APP_EXE}" 0
  CreateShortcut "$SMPROGRAMS\PageCapture\Uninstall.lnk" \
    "$INSTDIR\Uninstall.exe"

  ; App registry entries
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

  ; ── Firefox Native Messaging registration ──────────────────────────────────
  ; Tells Firefox where to find the host manifest so it can launch the app.
  ; HKCU = current user only (no extra admin rights needed beyond install).
  WriteRegStr HKCU \
    "Software\Mozilla\NativeMessagingHosts\com.careerquest.pagecapture" \
    "" \
    "$INSTDIR\com.careerquest.pagecapture.json"

  MessageBox MB_OK "PageCapture installed!$\n$\nTo use it in Firefox:$\n1. Open Firefox$\n2. Go to about:addons$\n3. Click the gear icon > Install Add-on From File$\n4. Browse to: $INSTDIR\extension\manifest.json$\n$\nThen click the PageCapture button in your Firefox toolbar to capture any page."

SectionEnd

Section "Uninstall"
  Delete "$INSTDIR\Uninstall.exe"
  RMDir /r "$INSTDIR"
  Delete "$DESKTOP\PageCapture.lnk"
  RMDir /r "$SMPROGRAMS\PageCapture"

  ; App registry cleanup
  DeleteRegKey HKCU "Software\PageCapture"
  DeleteRegKey HKLM \
    "Software\Microsoft\Windows\CurrentVersion\Uninstall\PageCapture"

  ; Firefox Native Messaging cleanup
  DeleteRegKey HKCU \
    "Software\Mozilla\NativeMessagingHosts\com.careerquest.pagecapture"

SectionEnd
