; Arquivo de configuração do Inno Setup para o aplicativo
[Setup]
AppName=Torp Controle EPI
AppVersion=1.0
DefaultDirName={userappdata}\TorpControleEPI
DefaultGroupName=Torp Controle EPI
OutputDir=output
OutputBaseFilename=TorpControleEPI_Instalador
Compression=lzma
SolidCompression=yes

[Files]
; Inclui o executável gerado pelo PyInstaller e o banco de dados
Source: "dist\main.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\database.db"; DestDir: "{app}"; Flags: ignoreversion


[Icons]
; Atalho no menu Iniciar
Name: "{group}\Torp Controle EPI"; Filename: "{app}\main.exe"; WorkingDir: "{app}"; IconFilename: "{app}\icon.ico"
; Atalho na área de trabalho
Name: "{userdesktop}\Torp Controle EPI"; Filename: "{app}\main.exe"; WorkingDir: "{app}"; IconFilename: "{app}\icon.ico"

[Run]
; Executar o programa após a instalação (opcional)
Filename: "{app}\main.exe"; Description: "Executar Torp Controle EPI"; Flags: nowait postinstall skipifsilent
