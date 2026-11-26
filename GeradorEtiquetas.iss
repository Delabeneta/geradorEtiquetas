; Instalador do Gerador de Etiquetas Pimaco (Inno Setup 6)
; Gerado a partir da pasta dist\GeradorEtiquetas (PyInstaller --onedir).
; A versão vem do build.py (/DMyAppVersion=...); o valor abaixo é só um padrão.

#ifndef MyAppVersion
  #define MyAppVersion "3.0.0"
#endif
#define MyAppName "Gerador de Etiquetas Pimaco"
#define MyAppPublisher "Delabeneta"
#define MyAppExeName "GeradorEtiquetas.exe"

[Setup]
; ATENÇÃO: o AppId identifica o programa nas atualizações. NUNCA mude este GUID,
; senão uma nova versão será instalada ao lado da antiga em vez de substituí-la.
AppId={{5E5D0E83-BA69-4CCB-A75C-E0E2C45AE596}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
VersionInfoVersion={#MyAppVersion}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; Instala só para o usuário atual: não pede senha de administrador (nem nas atualizações).
PrivilegesRequired=lowest
OutputDir=installer
; Nome fixo, para o updater sempre encontrar o mesmo asset na release do GitHub.
OutputBaseFilename=GeradorEtiquetas-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
UninstallDisplayIcon={app}\{#MyAppExeName}
#ifexist "icone.ico"
SetupIconFile=icone.ico
#endif
; Para assinar o instalador pelo próprio Inno Setup, configure um SignTool no IDE
; (Tools > Configure Sign Tools) e descomente:
; SignTool=minhaassinatura

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Área de Trabalho"; GroupDescription: "Atalhos:"; Flags: unchecked

[Files]
Source: "dist\GeradorEtiquetas\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Instalação manual: oferece abrir o programa ao final.
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir o {#MyAppName}"; Flags: nowait postinstall skipifsilent
; Atualização automática (modo silencioso): reabre o programa sozinho.
Filename: "{app}\{#MyAppExeName}"; Flags: nowait; Check: WizardSilent
