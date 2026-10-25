; ============================================================================
; Instalador - Fechamento para Gráfica (CorelDRAW)
; Compilar com o Inno Setup 6 (Build > Compile)
;
; O instalador detecta as instalações do CorelDRAW e copia FechamentoGrafica.gms para as respectivas pastas Draw\GMS.
;
; ESTRUTURA ESPERADA:
;
;   FechamentoGrafica.iss
;   dist\FechamentoGrafica.gms    <- projeto de macros criado no Corel (ATUALIZE antes de compilar)
;   dist\icone.ico                <- ícone (formato .ico)
;   dist\MeuBotao.cdws            <- opcional (workspace com o botão)
;
; Saída: pasta "dist_setup\"
; ============================================================================

#define NomeApp     "Fechamento para Gráfica (CorelDRAW)"
#define VersaoApp   "1.1.0"
#define ArquivoGMS  "FechamentoGrafica.gms"

[Setup]
AppId={{1C565EED-738D-4F33-9AF1-2C71830D0C54}
AppName={#NomeApp}
AppVersion={#VersaoApp}
AppPublisher=Marcos
DefaultDirName={autopf}\FechamentoGrafica
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=admin
; instalador de 64 bits: {commonpf} = C:\Program Files, {commonpf32} = C:\Program Files (x86)
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=dist_setup
OutputBaseFilename=Instalador_FechamentoGrafica_{#VersaoApp}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupIconFile=dist\icone.ico
UninstallDisplayIcon={app}\icone.ico
UninstallDisplayName={#NomeApp}

[Languages]
Name: "ptbr"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Files]
Source: "dist\{#ArquivoGMS}"; DestDir: "{app}\gms"; Flags: ignoreversion
Source: "dist\icone.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\MeuBotao.cdws"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

[UninstallDelete]
Type: files; Name: "{app}\instalados.txt"
Type: dirifempty; Name: "{app}\gms"
Type: dirifempty; Name: "{app}"

[Code]
var
  PaginaVersoes: TInputOptionWizardPage;
  PastasGMS: TStringList;

{ Verifica se o CorelDRAW está aberto (o arquivo .gms fica travado enquanto o Corel roda) }
function CorelEmExecucao: Boolean;
var
  Codigo: Integer;
begin
  Result := False;

  if Exec(ExpandConstant('{cmd}'),
          '/C tasklist | find /I "CorelDRW.exe" > nul',
          '', SW_HIDE, ewWaitUntilTerminated, Codigo) then
  begin
    if Codigo = 0 then
    begin
      Result := True;
      Exit;
    end;
  end;

  if Exec(ExpandConstant('{cmd}'),
          '/C tasklist | find /I "CorelDRW64.exe" > nul',
          '', SW_HIDE, ewWaitUntilTerminated, Codigo) then
  begin
    if Codigo = 0 then
    begin
      Result := True;
      Exit;
    end;
  end;
end;

{ Procura em <Base>\* as pastas que tenham Draw\GMS (uma por versão instalada) }
procedure ProcurarEm(const Base, Rotulo: String);
var
  Achado: TFindRec;
  Pasta: String;
begin
  if FindFirst(Base + '\*', Achado) then
  begin
    try
      repeat
        if ((Achado.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0) and
           (Achado.Name <> '.') and (Achado.Name <> '..') then
        begin
          Pasta := Base + '\' + Achado.Name + '\Draw\GMS';
          if DirExists(Pasta) then
          begin
            PastasGMS.Add(Pasta);
            PaginaVersoes.Add(Achado.Name + ' (' + Rotulo + ')');
          end;
        end;
      until not FindNext(Achado);
    finally
      FindClose(Achado);
    end;
  end;
end;

procedure ProcurarVersoes;
begin
  ProcurarEm(ExpandConstant('{commonpf}\Corel'), 'Program Files');
  ProcurarEm(ExpandConstant('{commonpf32}\Corel'), 'Program Files (x86)');
end;

procedure InitializeWizard;
var
  i: Integer;
begin
  PastasGMS := TStringList.Create;
  PaginaVersoes := CreateInputOptionPage(wpWelcome,
    'Instalações do CorelDRAW',
    'Onde instalar a macro?',
    'Marque as instalações do CorelDRAW onde a macro deve ser instalada. ' +
    'Feche o CorelDRAW antes de continuar.',
    False, False);
  ProcurarVersoes;
  for i := 0 to PaginaVersoes.CheckListBox.Items.Count - 1 do
    PaginaVersoes.Values[i] := True;
  if PastasGMS.Count = 0 then
    PaginaVersoes.SubCaptionLabel.Caption :=
      'Nenhuma instalação do CorelDRAW foi encontrada em Program Files\Corel. ' +
      'Verifique se o CorelDRAW está instalado neste computador.';
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  i: Integer;
  Algum: Boolean;
begin
  Result := True;
  if CurPageID = PaginaVersoes.ID then
  begin
    Algum := False;
    for i := 0 to PaginaVersoes.CheckListBox.Items.Count - 1 do
      if PaginaVersoes.Values[i] then Algum := True;
    if not Algum then
    begin
      MsgBox('Marque pelo menos uma instalação do CorelDRAW.', mbInformation, MB_OK);
      Result := False;
    end
    else if CorelEmExecucao then
    begin
      MsgBox('O CorelDRAW está aberto. Feche-o e clique em Avançar novamente.', mbInformation, MB_OK);
      Result := False;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  i: Integer;
  Origem, Destino, Lista: String;
begin
  if CurStep = ssPostInstall then
  begin
    Origem := ExpandConstant('{app}\gms\{#ArquivoGMS}');
    Lista := '';
    for i := 0 to PastasGMS.Count - 1 do
    begin
      if PaginaVersoes.Values[i] then
      begin
        Destino := PastasGMS[i] + '\{#ArquivoGMS}';
        { FALSE = sobrescreve se já existir (necessário para atualizar/reinstalar) }
        if CopyFile(Origem, Destino, False) then
          Lista := Lista + Destino + #13#10
        else
          MsgBox(
            'Não foi possível copiar a macro para:' + #13#10 +
            Destino + #13#10 +
            'Verifique se o CorelDRAW está fechado.',
            mbError,
            MB_OK
          );
      end;
    end;
    SaveStringToFile(ExpandConstant('{app}\instalados.txt'), Lista, False);
  end;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = wpFinished then
    WizardForm.FinishedLabel.Caption :=
      'A macro foi instalada com sucesso.' + #13#10#13#10 +
      'Abra o CorelDRAW e execute a macro "FecharDocumentoAtivo" ' +
      '(Ferramentas > Macros > Gerenciador de macros), ' +
      'ou crie um botão arrastando-a em Ferramentas > Personalização > Comandos > Macros.' + #13#10#13#10 +
      'O botão pronto (MeuBotao.cdws) está em:' + #13#10 + ExpandConstant('{app}');
end;

{ ---- desinstalação ---- }
function InitializeUninstall: Boolean;
begin
  Result := True;
  if CorelEmExecucao then
  begin
    MsgBox('O CorelDRAW está aberto. Feche-o e execute a desinstalação novamente.', mbInformation, MB_OK);
    Result := False;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Linhas: TArrayOfString;
  i: Integer;
begin
  if CurUninstallStep = usUninstall then
  begin
    if LoadStringsFromFile(ExpandConstant('{app}\instalados.txt'), Linhas) then
    begin
      for i := 0 to GetArrayLength(Linhas) - 1 do
      begin
        if Trim(Linhas[i]) <> '' then
          DeleteFile(Trim(Linhas[i]));
      end;
    end;
  end;
end;
