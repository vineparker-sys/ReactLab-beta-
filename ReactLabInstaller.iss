; -- ReactLabInstaller.iss --
; Script de instalação para o aplicativo ReactLab

[Setup]
; Configurações gerais
AppName=ReactLab
AppVersion=1.0
DefaultDirName={pf}\ReactLab
DefaultGroupName=ReactLab
OutputBaseFilename=ReactLabInstaller
SetupIconFile=assets\logo.ico
Compression=lzma
SolidCompression=yes

; Arquivo de licença (opcional)
;LicenseFile=LICENSE.txt

[Languages]
Name: "portuguese"; MessagesFile: "compiler:Languages\Portuguese.isl"

[Files]
; Executável principal
Source: "release_build\ReactLab.exe"; DestDir: "{app}"; Flags: ignoreversion

; Pasta assets
Source: "release_build\assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs

; Pasta locales (idiomas)
Source: "release_build\locales\*"; DestDir: "{app}\locales"; Flags: ignoreversion recursesubdirs createallsubdirs

; Pasta themes (temas)
Source: "release_build\themes\*"; DestDir: "{app}\themes"; Flags: ignoreversion recursesubdirs createallsubdirs

; Outros arquivos necessários
;Source: "release_build\outros\*"; DestDir: "{app}\outros"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\ReactLab"; Filename: "{app}\ReactLab.exe"; IconFilename: "{app}\assets\logo.ico"
Name: "{group}\Desinstalar ReactLab"; Filename: "{uninstallexe}"
Name: "{commondesktop}\ReactLab"; Filename: "{app}\ReactLab.exe"; IconFilename: "{app}\assets\logo.ico"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Área de Trabalho para todos os usuários"; GroupDescription: "Atalhos"

[Run]
; Executar o aplicativo após a instalação (opcional)
;Filename: "{app}\ReactLab.exe"; Description: "{cm:LaunchProgram,ReactLab}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Remover arquivos ou pastas adicionais na desinstalação (se necessário)
;Type: filesandordirs; Name: "{app}\alguma_pasta"

[Messages]
; Personalização de mensagens (opcional)

[CustomMessages]
; Mensagens personalizadas para suporte a vários idiomas
;LaunchProgram=Iniciar {#StringChange(AppName, '&', '&&')}
