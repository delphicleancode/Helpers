unit cxGrid.Helper;

interface
  uses
    System.Classes,
    System.SysUtils,
    System.IniFiles,
    Vcl.Dialogs,
    cxGridDBTableView,
    cxFilter,
    cxGrid,
    cxDBPivotGrid,
    cxCustomPivotGrid,
    cxCustomData,
    cxGridExportLink,
    cxGridDBBandedTableView;

type

  TcxGridsHelper = class helper for TcxGrid
  public
    procedure Exportar;
  end;

  TcxGridHelper = class helper for TcxGridDBTableView
  public
    procedure Filtrar(const FieldName: string; FilterOperator: TcxFilterOperatorKind; const Value: string);
    procedure SalvarLayout(FormName: string; Posicao: Integer; const NomeArquivo: string);
    procedure SalvarLayoutPadrao(const FormName: string; Sobreescrever: Boolean = False);
    procedure RestaurarLayout(FormName: string; Posicao: Integer);
    procedure RestaurarLayoutPadrao(const FormName: string);
    procedure RemoverLayout(FormName: string; Posicao: Integer);
    procedure Layouts(FormName: string; Componente: TStrings);
    procedure AdicionarItemSumario(const IndiceColuna: Integer; TipoExpressao: TcxSummaryKind; const Formato: string; Posicao: TcxSummaryPosition = spFooter);
    procedure AdicionarGrupoSumario(const IndiceColuna: Integer; TipoExpressao: TcxSummaryKind; const Formato: string; Posicao: TcxSummaryPosition = spFooter);
    procedure ClearDefaultGroupSummaryItems;
  end;

  TcxPivotGridHelper = class helper for TcxDBPivotGrid
  public
    procedure Exportar;
    procedure SetEvents;
    procedure GetStoredPropertiesHelper(Sender: TObject; AProperties: TStrings);
    procedure GetStoredPropertyValueHelper(Sender: TObject; const AName: string; var AValue: Variant);
    procedure SetStoredPropertyValueHelper(Sender: TObject; const AName: string; const AValue: Variant);

    procedure SalvarLayout(FormName: string; Posicao: Integer; const NomeArquivo: string);
    procedure SalvarLayoutPadrao(const FormName: string; Sobreescrever: Boolean = False);
    procedure RestaurarLayout(FormName: string; Posicao: Integer);
    procedure RestaurarLayoutPadrao(const FormName: string);
    procedure RemoverLayout(FormName: string; Posicao: Integer);
    procedure Layouts(FormName: string; Componente: TStrings);
  end;

  TcxGridDBBandedTableViewHelper = class helper for TcxGridDBBandedTableView
  public
    procedure Filtrar(const FieldName: string; FilterOperator: TcxFilterOperatorKind; const Value: string);
    procedure SalvarLayout(FormName: string; Posicao: Integer; const NomeArquivo: string);
    procedure SalvarLayoutPadrao(const FormName: string; Sobreescrever: Boolean = False);
    procedure RestaurarLayout(FormName: string; Posicao: Integer);
    procedure RestaurarLayoutPadrao(const FormName: string);
    procedure RemoverLayout(FormName: string; Posicao: Integer);
    procedure Layouts(FormName: string; Componente: TStrings);
    procedure AdicionarItemSumario(const IndiceColuna: Integer; TipoExpressao: TcxSummaryKind; const Formato: string; Posicao: TcxSummaryPosition = spFooter);
    procedure AdicionarGrupoSumario(const IndiceColuna: Integer; TipoExpressao: TcxSummaryKind; const Formato: string; Posicao: TcxSummaryPosition = spFooter);
    procedure ClearDefaultGroupSummaryItems;
  end;

implementation
  uses
    cxExportPivotGridLink;

procedure TcxGridHelper.Layouts(FormName: string; Componente: TStrings);
var
  IniLayouts: TIniFile;
  Secao: string;
  Items: TStrings;
  I: Integer;
begin
  try
    IniLayouts := TIniFile.Create('LayoutGrid.ini');

    Secao := FormName + Self.Name;

    if IniLayouts.SectionExists(Secao) then
    begin
      Items := TStringList.Create;
      IniLayouts.ReadSectionValues(Secao, Items);

      for I := 0 to Pred(Items.Count) do
        Componente.Add(IniLayouts.ReadString(Secao, IntToStr(I), ''));

    end
    else
    begin
      IniLayouts.WriteString(Secao, '0', 'Padrao');
      IniLayouts.WriteString(Secao+'Files', '0', FormName + Self.Name +'Custom.lyg');
    end;

  finally
    FreeAndNil(IniLayouts);
  end;
end;

procedure TcxGridHelper.AdicionarGrupoSumario(const IndiceColuna: Integer; TipoExpressao: TcxSummaryKind; const Formato: string; Posicao: TcxSummaryPosition);
begin
  Self.DataController.Summary.DefaultGroupSummaryItems.Add(Self.Columns[IndiceColuna], Posicao, TipoExpressao, Formato);
end;

procedure TcxGridHelper.AdicionarItemSumario(const IndiceColuna: Integer; TipoExpressao: TcxSummaryKind; const Formato: string; Posicao: TcxSummaryPosition);
begin
  Self.DataController.Summary.FooterSummaryItems.Add(Self.Columns[IndiceColuna], Posicao, TipoExpressao, Formato);
end;

procedure TcxGridHelper.ClearDefaultGroupSummaryItems;
begin
  if Self.DataController.Summary.DefaultGroupSummaryItems.Count > 0 then
    Self.DataController.Summary.DefaultGroupSummaryItems.Clear;
end;

procedure TcxGridsHelper.Exportar;
var
  FileExt : String;
  SaveDialog: TSaveDialog;
begin
  SaveDialog := nil;
  try
    SaveDialog := TSaveDialog.Create(nil);

    SaveDialog.Filter :=
      'Excel (*.xlsx) |*.xlsx|'+
      'Excel 97-2003 (*.xls) |*.xls|'+
      'Arquivo CSV (*.csv) |*.csv|'+
      'Arquivo Texto (*.txt) |*.txt|'+
      'XML (*.xml) |*.xml|'+
      'BrOffice (*.ods) |*.ods|'+
      'Página Web (*.html)|*.html';

    SaveDialog.Title := 'Exportar Dados';
    SaveDialog.DefaultExt:= 'xlsx';

    if SaveDialog.Execute then
    begin
      FileExt := LowerCase(ExtractFileExt(SaveDialog.FileName));

      if FileExt = '.xlsx' then
        ExportGridToXLSX(SaveDialog.FileName,Self, False, True, False)
      else if FileExt = '.xls' then
        ExportGridToExcel(SaveDialog.FileName,Self, False, True, False)
      else if FileExt = '.csv' then
        ExportGridToText(SaveDialog.FileName,Self, True, True, ';','"','"', 'csv')
      else if FileExt = '.txt' then
        ExportGridToText(SaveDialog.FileName,Self, False)
      else if FileExt = '.xml' then
        ExportGridToXML(SaveDialog.FileName,Self, False)
      else if FileExt = '.ods' then
        ExportGridToExcel(SaveDialog.FileName,Self, False, True, False)
      else if FileExt = '.html' then
        ExportGridToHTML(SaveDialog.FileName,Self, False);
    end;
  finally
    FreeAndNil(SaveDialog);
  end;
end;

procedure TcxGridHelper.Filtrar(const FieldName: string; FilterOperator: TcxFilterOperatorKind; const Value: string);
var
  AItemList: TcxFilterCriteriaItemList;
begin
  with DataController.Filter do
  begin
    BeginUpdate;
    try
      Root.Clear;
      Root.BoolOperatorKind := fboOr;
      AItemList := Root.AddItemList(fboOr);
      AItemList.AddItem(Self.FindItemByName(FieldName), FilterOperator, Value, Value);
    finally
      EndUpdate;
    end;
  end;
end;

procedure TcxGridHelper.SalvarLayout(FormName: string; Posicao: Integer; const NomeArquivo: string);
var
  IniLayouts: TIniFile;
  Secao: string;
  Arquivo: string;
begin
  try
    IniLayouts := TIniFile.Create('LayoutGrid.ini');

    Secao   := FormName + Self.Name;
    Arquivo := Secao + NomeArquivo +'.lyg';

    IniLayouts.WriteString(Secao, IntToStr(Posicao), NomeArquivo);
    IniLayouts.WriteString(Secao+'Files', IntToStr(Posicao), Arquivo);

    StoreToIniFile(Arquivo);

  finally
    if Assigned(IniLayouts) then
      FreeAndNil(IniLayouts);
  end;
end;

procedure TcxGridHelper.RemoverLayout(FormName: string; Posicao: Integer);
var
  IniLayouts: TIniFile;
  Secao: string;
  Arquivo: string;
  Chave: string;
begin
  try
    IniLayouts := TIniFile.Create('LayoutGrid.ini');

    Secao   := FormName + Self.Name;
    Chave   := IntToStr(Posicao);
    Arquivo := IniLayouts.ReadString(Secao+'Files', Chave, '');

    if FileExists(PWideChar(Arquivo)) then
      DeleteFile(PWideChar(Arquivo));

    IniLayouts.DeleteKey(Secao, Chave);
    IniLayouts.DeleteKey(Secao+'Files', Chave);

  finally
    if Assigned(IniLayouts) then
      FreeAndNil(IniLayouts);
  end;
end;

procedure TcxGridHelper.RestaurarLayout(FormName: string; Posicao: Integer);
var
  IniLayouts: TIniFile;
  Secao: string;
  Arquivo: string;
begin
  try
    IniLayouts := TIniFile.Create('LayoutGrid.ini');

    Secao   := FormName + Self.Name;
    Arquivo := IniLayouts.ReadString(Secao+'Files', IntToStr(Posicao), '');

    if FileExists(Arquivo) then
      RestoreFromIniFile(Arquivo);

  finally
    if Assigned(IniLayouts) then
      FreeAndNil(IniLayouts);
  end;
end;

procedure TcxGridHelper.RestaurarLayoutPadrao(const FormName: string);
var
  ArqGridCustom: string;
begin
  ArqGridCustom := FormName + Self.Name +'Custom.lyg';

  if FileExists(ArqGridCustom) then
    RestoreFromIniFile(ArqGridCustom);
end;

procedure TcxGridHelper.SalvarLayoutPadrao(const FormName: string; Sobreescrever: Boolean);
var
  ArqGridCustom: string;
begin
  ArqGridCustom := FormName + Self.Name +'Custom.lyg';

  if not FileExists(ArqGridCustom) or Sobreescrever then
    StoreToIniFile(ArqGridCustom);
end;

procedure TcxPivotGridHelper.Exportar;
var
  sFileExt : String;
  eSaveDialog: TSaveDialog;
begin

  eSaveDialog := nil;
  try
    eSaveDialog := TSaveDialog.Create(nil);

    eSaveDialog.Filter :=
      'Excel 97-2003 (*.xls) |*.xls|'+
      'Arquivo CSV (*.csv) |*.csv|'+
      'Arquivo Texto (*.txt) |*.txt|'+
      'XML (*.xml) |*.xml|'+
      'Página Web (*.html)|*.html';

    eSaveDialog.Title := 'Exportar Dados';
    eSaveDialog.DefaultExt:= 'xls';

    if eSaveDialog.Execute then
    begin
      sFileExt := LowerCase(ExtractFileExt(eSaveDialog.FileName));

      if sFileExt = '.xls' then
        cxExportPivotGridToExcel(eSaveDialog.FileName, Self, False, True)
      else if sFileExt = '.csv' then
        cxExportPivotGridToText(eSaveDialog.FileName, Self, True, ';', '"', '"', 'csv')
//      else if sFileExt = '.txt' then
//        cxExportPivotGridToText(eSaveDialog.FileName, Self, False, 'txt')
      else if sFileExt = '.xml' then
        cxExportPivotGridToXML(eSaveDialog.FileName, Self, False)
      else if sFileExt = '.html' then
        cxExportPivotGridToHTML(eSaveDialog.FileName, Self, False);
    end;
  finally
    FreeAndNil(eSaveDialog);
  end;
end;

procedure TcxPivotGridHelper.GetStoredPropertiesHelper(Sender: TObject; AProperties: TStrings);
begin
  AProperties.Add('DataFieldArea');
end;

procedure TcxPivotGridHelper.GetStoredPropertyValueHelper(Sender: TObject; const AName: string; var AValue: Variant);
begin
  if AName = 'DataFieldArea' then
    AValue := TcxCustomPivotGrid(Sender).OptionsDataField.Area;

end;

procedure TcxPivotGridHelper.Layouts(FormName: string; Componente: TStrings);
var
  IniLayouts: TIniFile;
  Secao: string;
  Items: TStrings;
  I: Integer;
begin
  SetEvents;

  try
    IniLayouts := TIniFile.Create('LayoutGrid.ini');

    Secao := FormName + Self.Name;

    if IniLayouts.SectionExists(Secao) then
    begin
      Items := TStringList.Create;
      IniLayouts.ReadSectionValues(Secao, Items);

      for I := 0 to Pred(Items.Count) do
        Componente.Add(IniLayouts.ReadString(Secao, IntToStr(I), ''));

    end
    else
    begin
      IniLayouts.WriteString(Secao, '0', 'Padrao');
      IniLayouts.WriteString(Secao+'Files', '0', FormName + Self.Name +'Custom.lyg');
    end;

  finally
    FreeAndNil(IniLayouts);
  end;
end;

procedure TcxPivotGridHelper.SalvarLayout(FormName: string; Posicao: Integer; const NomeArquivo: string);
var
  IniLayouts: TIniFile;
  Secao: string;
  Arquivo: string;
begin
  SetEvents;

  try

    IniLayouts := TIniFile.Create('LayoutGrid.ini');

    Secao   := FormName + Self.Name;
    Arquivo := Secao + NomeArquivo +'.lyg';

    IniLayouts.WriteString(Secao, IntToStr(Posicao), NomeArquivo);
    IniLayouts.WriteString(Secao+'Files', IntToStr(Posicao), Arquivo);

    StoreToIniFile(Arquivo);

  finally
    if Assigned(IniLayouts) then
      FreeAndNil(IniLayouts);
  end;
end;

procedure TcxPivotGridHelper.RemoverLayout(FormName: string; Posicao: Integer);
var
  IniLayouts: TIniFile;
  Secao: string;
  Arquivo: string;
  Chave: string;
begin
  SetEvents;

  try
    IniLayouts := TIniFile.Create('LayoutGrid.ini');

    Secao   := FormName + Self.Name;
    Chave   := IntToStr(Posicao);
    Arquivo := IniLayouts.ReadString(Secao+'Files', Chave, '');

    if FileExists(PWideChar(Arquivo)) then
      DeleteFile(PWideChar(Arquivo));

    IniLayouts.DeleteKey(Secao, Chave);
    IniLayouts.DeleteKey(Secao+'Files', Chave);

  finally
    if Assigned(IniLayouts) then
      FreeAndNil(IniLayouts);
  end;
end;

procedure TcxPivotGridHelper.RestaurarLayout(FormName: string; Posicao: Integer);
var
  IniLayouts: TIniFile;
  Secao: string;
  Arquivo: string;
begin
  SetEvents;

  try
    IniLayouts := TIniFile.Create('LayoutGrid.ini');

    Secao   := FormName + Self.Name;
    Arquivo := IniLayouts.ReadString(Secao+'Files', IntToStr(Posicao), '');

    if FileExists(Arquivo) then
      RestoreFromIniFile(Arquivo);

  finally
    if Assigned(IniLayouts) then
      FreeAndNil(IniLayouts);
  end;
end;

procedure TcxPivotGridHelper.RestaurarLayoutPadrao(const FormName: string);
var
  ArqGridCustom: string;
begin
  SetEvents;

  ArqGridCustom := FormName + Self.Name +'Custom.lyg';

  if FileExists(ArqGridCustom) then
    RestoreFromIniFile(ArqGridCustom);
end;

procedure TcxPivotGridHelper.SalvarLayoutPadrao(const FormName: string; Sobreescrever: Boolean);
var
  ArqGridCustom: string;
begin
  SetEvents;

  ArqGridCustom := FormName + Self.Name +'Custom.lyg';

  if not FileExists(ArqGridCustom) or Sobreescrever then
    StoreToIniFile(ArqGridCustom);
end;

procedure TcxPivotGridHelper.SetEvents;
begin
  Self.OnGetStoredProperties    := GetStoredPropertiesHelper;
  Self.OnGetStoredPropertyValue := GetStoredPropertyValueHelper;
  Self.OnSetStoredPropertyValue := SetStoredPropertyValueHelper;
end;

procedure TcxPivotGridHelper.SetStoredPropertyValueHelper(Sender: TObject; const AName: string; const AValue: Variant);
begin
  if AName = 'DataFieldArea' then
    TcxCustomPivotGrid(Sender).OptionsDataField.Area := AValue;
end;

procedure TcxGridDBBandedTableViewHelper.AdicionarGrupoSumario(const IndiceColuna: Integer; TipoExpressao: TcxSummaryKind;const Formato: string; Posicao: TcxSummaryPosition);
begin
  Self.DataController.Summary.DefaultGroupSummaryItems.Add(Self.Columns[IndiceColuna], Posicao, TipoExpressao, Formato);
end;

procedure TcxGridDBBandedTableViewHelper.AdicionarItemSumario(const IndiceColuna: Integer; TipoExpressao: TcxSummaryKind; const Formato: string; Posicao: TcxSummaryPosition);
begin
  Self.DataController.Summary.FooterSummaryItems.Add(Self.Columns[IndiceColuna], Posicao, TipoExpressao, Formato);
end;

procedure TcxGridDBBandedTableViewHelper.ClearDefaultGroupSummaryItems;
begin
  if Self.DataController.Summary.DefaultGroupSummaryItems.Count > 0 then
    Self.DataController.Summary.DefaultGroupSummaryItems.Clear;
end;

procedure TcxGridDBBandedTableViewHelper.Filtrar(const FieldName: string; FilterOperator: TcxFilterOperatorKind; const Value: string);
var
  AItemList: TcxFilterCriteriaItemList;
begin
  with DataController.Filter do
  begin
    BeginUpdate;
    try
      Root.Clear;
      Root.BoolOperatorKind := fboOr;
      AItemList := Root.AddItemList(fboOr);
      AItemList.AddItem(Self.FindItemByName(FieldName), FilterOperator, Value, Value);
    finally
      EndUpdate;
    end;
  end;
end;

procedure TcxGridDBBandedTableViewHelper.Layouts(FormName: string; Componente: TStrings);
var
  IniLayouts: TIniFile;
  Secao: string;
  Items: TStrings;
  I: Integer;
begin
  try
    IniLayouts := TIniFile.Create('LayoutGrid.ini');

    Secao := FormName + Self.Name;

    if IniLayouts.SectionExists(Secao) then
    begin
      Items := TStringList.Create;
      IniLayouts.ReadSectionValues(Secao, Items);

      for I := 0 to Pred(Items.Count) do
        Componente.Add(IniLayouts.ReadString(Secao, IntToStr(I), ''));

    end
    else
    begin
      IniLayouts.WriteString(Secao, '0', 'Padrao');
      IniLayouts.WriteString(Secao+'Files', '0', FormName + Self.Name +'Custom.lyg');
    end;

  finally
    FreeAndNil(IniLayouts);
  end;
end;

procedure TcxGridDBBandedTableViewHelper.RemoverLayout(FormName: string; Posicao: Integer);
var
  IniLayouts: TIniFile;
  Secao: string;
  Arquivo: string;
  Chave: string;
begin
  try
    IniLayouts := TIniFile.Create('LayoutGrid.ini');

    Secao   := FormName + Self.Name;
    Chave   := IntToStr(Posicao);
    Arquivo := IniLayouts.ReadString(Secao+'Files', Chave, '');

    if FileExists(PWideChar(Arquivo)) then
      DeleteFile(PWideChar(Arquivo));

    IniLayouts.DeleteKey(Secao, Chave);
    IniLayouts.DeleteKey(Secao+'Files', Chave);

  finally
    if Assigned(IniLayouts) then
      FreeAndNil(IniLayouts);
  end;
end;

procedure TcxGridDBBandedTableViewHelper.RestaurarLayout(FormName: string; Posicao: Integer);
var
  IniLayouts: TIniFile;
  Secao: string;
  Arquivo: string;
begin
  try
    IniLayouts := TIniFile.Create('LayoutGrid.ini');

    Secao   := FormName + Self.Name;
    Arquivo := IniLayouts.ReadString(Secao+'Files', IntToStr(Posicao), '');

    if FileExists(Arquivo) then
      RestoreFromIniFile(Arquivo);

  finally
    if Assigned(IniLayouts) then
      FreeAndNil(IniLayouts);
  end;
end;

procedure TcxGridDBBandedTableViewHelper.RestaurarLayoutPadrao(const FormName: string);
var
  ArqGridCustom: string;
begin
  ArqGridCustom := FormName + Self.Name +'Custom.lyg';

  if FileExists(ArqGridCustom) then
    RestoreFromIniFile(ArqGridCustom);
end;

procedure TcxGridDBBandedTableViewHelper.SalvarLayout(FormName: string; Posicao: Integer; const NomeArquivo: string);
var
  IniLayouts: TIniFile;
  Secao: string;
  Arquivo: string;
begin
  try
    IniLayouts := TIniFile.Create('LayoutGrid.ini');

    Secao   := FormName + Self.Name;
    Arquivo := Secao + NomeArquivo +'.lyg';

    IniLayouts.WriteString(Secao, IntToStr(Posicao), NomeArquivo);
    IniLayouts.WriteString(Secao+'Files', IntToStr(Posicao), Arquivo);

    StoreToIniFile(Arquivo);

  finally
    if Assigned(IniLayouts) then
      FreeAndNil(IniLayouts);
  end;
end;

procedure TcxGridDBBandedTableViewHelper.SalvarLayoutPadrao(const FormName: string; Sobreescrever: Boolean);
var
  ArqGridCustom: string;
begin
  ArqGridCustom := FormName + Self.Name +'Custom.lyg';

  if not FileExists(ArqGridCustom) or Sobreescrever then
    StoreToIniFile(ArqGridCustom);
end;

end.
