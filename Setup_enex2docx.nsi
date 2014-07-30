; example2.nsi
;
; This script is based on example1.nsi, but it remember the directory, 
; has uninstall support and (optionally) installs start menu shortcuts.
;
; It will install example2.nsi into a directory that the user selects,

;--------------------------------

!define APP "enex2docx"
!define COM "HIRAOKA HYPERS TOOLS, Inc."
!define VER "0.1"
!define APV "0_1"

; The name of the installer
Name "${APP} ${VER}"

; The file to write
OutFile "Setup_${APP}_${APV}.exe"

; The default installation directory
InstallDir "$PROGRAMFILES\${APP}"

; Registry key to check for directory (so if you install again, it will 
; overwrite the old one automatically)
InstallDirRegKey HKLM "Software\${COM}\${APP}" "Install_Dir"

; Request application privileges for Windows Vista
RequestExecutionLevel admin

;--------------------------------

; Pages

Page components
Page directory
Page instfiles

UninstPage uninstConfirm
UninstPage instfiles

;--------------------------------

; The stuff to install
Section ""

  SectionIn RO
  
  ; Set output path to the installation directory.
  SetOutPath $INSTDIR
  
  ; Put file there
  File /r /x "*.vshost.*" "bin\DEBUG\*.*"
  
  ; Write the installation path into the registry
  WriteRegStr HKLM "SOFTWARE\${COM}\${APP}" "Install_Dir" "$INSTDIR"
  
  ; Write the uninstall keys for Windows
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP}" "DisplayName" "${APP}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP}" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP}" "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP}" "NoRepair" 1
  WriteUninstaller "uninstall.exe"
  
SectionEnd

; Optional section (can be disabled by the user)
Section "スタートメニューへ登録"

  CreateDirectory "$SMPROGRAMS\${APP}"
  CreateShortCut "$SMPROGRAMS\${APP}\${APP}削除.lnk" "$INSTDIR\uninstall.exe" "" "$INSTDIR\uninstall.exe" 0
  CreateShortCut "$SMPROGRAMS\${APP}\EvernoteをWord2007文書に変換(${APP}).lnk" "$INSTDIR\${APP}.exe" "/select"
  
SectionEnd

;--------------------------------

; Uninstaller

Section "Uninstall"
  
  ; Remove registry keys
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP}"
  DeleteRegKey HKLM "Software\${COM}\${APP}"

  ; Remove files and uninstaller
  Delete "$INSTDIR\DocX.dll"
  Delete "$INSTDIR\DocX.pdb"
  Delete "$INSTDIR\DocX.xml"
  Delete "$INSTDIR\enex2docx.exe"
  Delete "$INSTDIR\enex2docx.exe.config"
  Delete "$INSTDIR\enex2docx.pdb"
  Delete "$INSTDIR\uninstall.exe"

  ; Remove shortcuts, if any
  Delete "$SMPROGRAMS\${APP}\*.lnk"

  ; Remove directories used
  RMDir "$SMPROGRAMS\${APP}"
  RMDir "$INSTDIR"

SectionEnd
