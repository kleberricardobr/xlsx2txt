unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComObj, Grids, DB, JvMemoryDataset, StrUtils;

type
  TForm1 = class(TForm)
    Button2: TButton;
    Button3: TButton;
    OpenDialog1: TOpenDialog;
    odXLS: TOpenDialog;
    StringGrid1: TStringGrid;
    qImportaPlan: TJvMemoryData;
    qImportaPlandt_pedido: TDateTimeField;
    qImportaPlandt_pagamento: TDateTimeField;
    qImportaPlandt_estorno: TDateTimeField;
    qImportaPlandt_prevista: TDateTimeField;
    qImportaPlanref_pedido: TStringField;
    qImportaPlanentrega: TLargeintField;
    qImportaPlantipo: TStringField;
    qImportaPlanvalor: TFloatField;
    qImportaPlandt_liberacao: TDateTimeField;
    lbStatus: TLabel;
    Button1: TButton;

    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    procedure LoadToDataSet;
    { Private declarations }
  public
    { Public declarations }
    function XlsxToTxt(pFileXLSX: String; pFileTXT: String; pCols: String): PAnsiChar;
  end;

  //usand stdcall ao invés de cdecl estava dando acess violation
  TSaveAsTXT = function(pFileXLSX, pFileTXT, pColsName: PAnsiChar): PAnsiChar; cdecl;

  TIndiceCols = record
  private
    DataPedido:    Byte;
    DataPagamento: Byte;
    DataEstorno:   Byte;
    DataLiberacao: Byte;
    DataPrevista:  Byte;
    RefPedido:     Byte;
    Entrega:       Byte;
    Tipo:          Byte;
    Valor:         Byte;
   public
    procedure SetIndice(pIndice: Integer; pCol: String);
  end;

var
  Form1: TForm1;
  Handler: THandle;

const
  xlCellTypeLastCell = 11;
  cCols = 'DATAPEDIDO;DATAPAGAMENTO;DATAESTORNO;DATALIBERAÇÃO;DATAPREVISTAPGTO;REF.PEDIDO;ENTREGA;TIPO;VALOR;';
  cFileDest = 'xlsx.txt';


implementation

{$R *.dfm}


//procedure TForm1.Button1Click(Sender: TObject);
//var
//  GetSheetName: TGetSheetName;
//  vByte: PAnsiChar;
//begin
//
//  try
//  //  ShowMessage( IntToStr(GetLastError) );
//   if (Handler = 0) then
//      raise Exception.Create('Erro ao carrega Dll. Código: ' + IntToStr(GetLastError));
//
//    @GetSheetName := GetProcAddress(Handler, 'GetSheetName');
//    if @GetSheetName <> nil then
//       vByte := GetSheetName(PAnsiChar('teste.xlsx'));
//
//     if vByte <> nil then
//        ShowMessage(vByte);
//  finally
//   // freelibrary(HD);
//  end;
//end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  ShowMessage('Não use stdcall, use cdecl');
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  ShowMessage(String( XlsxToTxt('teste.xlsx', cFileDest, cCols) ));
  lbStatus.Caption := 'Deu certo!!! Sem acess violation rs!';
  Button1.Click;
  //LoadToDataSet;
end;

procedure TForm1.LoadToDataSet;
var
 vFile: TStringList;
 vRow: TStringList;
 vLinha: Integer;
 vColuna: Integer;
 vCabecalho: TIndiceCols;
 vValorStr: String;

 procedure SetValues;
 begin
   if vColuna = vCabecalho.DataPedido then
   begin
     qImportaPlandt_pedido.AsDateTime := FloatToDateTime(StrToFloatDef(vValorStr, 0));
     exit;
   end;

   if vColuna = vCabecalho.DataPagamento then
   begin
     qImportaPlandt_pagamento.AsDateTime := FloatToDateTime(StrToFloatDef(vValorStr, 0));
     exit;
   end;

   if vColuna = vCabecalho.DataEstorno then
   begin
     qImportaPlandt_estorno.AsDateTime   := FloatToDateTime(StrToFloatDef(vValorStr, 0));
     exit;
   end;

   if vColuna = vCabecalho.DataLiberacao then
   begin
     qImportaPlandt_liberacao.AsDateTime := FloatToDateTime(StrToFloatDef(vValorStr, 0));
     exit;
   end;

   if vColuna = vCabecalho.DataPrevista then
   begin
     qImportaPlandt_prevista.AsDateTime  := FloatToDateTime(StrToFloatDef(vValorStr, 0));
     exit;
   end;

   if vColuna = vCabecalho.RefPedido then
   begin
     qImportaPlanref_pedido.AsString := vValorStr;
     exit;
   end;

   if vColuna = vCabecalho.Entrega then
   begin
     qImportaPlanentrega.AsLargeInt := StrToInt64Def(vValorStr, 0);
     exit;
   end;

   if vColuna = vCabecalho.Tipo then
   begin
     qImportaPlantipo.AsString := vValorStr;
     exit;
   end;

   if vColuna = vCabecalho.Valor then
   begin
     qImportaPlanvalor.AsFloat := StrToFloatDef(vValorStr, 0);
     exit;
   end;
 end;

begin
 vFile := TStringList.Create;
 vRow  := TStringList.Create;
 try
   vFile.LoadFromFile(cFileDest);

   vRow.Delimiter := ';';
   vRow.StrictDelimiter := true;

      if qImportaPlan = nil then showMessage('sim');

  // qImportaPlan.Close;
  // qImportaPlan.Open;

    for vLinha := 0 to vFile.Count - 1 do
    begin
      vRow.DelimitedText := Copy(vFile.Strings[vLinha], 1, Length(vFile.Strings[vLinha]) - 1 ) ;

      for vColuna := 0 to vRow.Count - 1 do
      begin
        vValorStr := vRow.Strings[vColuna];

        if vLinha = 0 then
        begin
          vCabecalho.SetIndice(vColuna, vRow.Strings[vColuna]);
          continue;
        end;

        //qImportaPlan.Append;

        //if Trim(vValorStr) <> '' then SetValues;

        //qImportaPlan.Post;
      end;
    end;
 finally
   FreeAndNil(vFile);
   FreeAndNil(vRow);
 end;
end;

procedure TForm1.Button3Click(Sender: TObject);
var
  XLSAplicacao, AbaXLS: OLEVariant;
  RangeMatrix: Variant;
  lastRow: Integer;
  lastColumn: Integer;
  k: Integer;
  r: integer;
begin
  if not odXLS.Execute then Exit;

  try
    try
      XLSAplicacao := CreateOleObject('Excel.Application');      
    except
    on E: Exception do
      begin
        ShowMessage(
        'Falha ao iniciar Excel. Certifique-se de que a aplicação esteja realmente instalada no computador.' + sLineBreak +
        'Erro: '+E.Message);
        Abort;
      end;
    end;

    XLSAplicacao.Workbooks.Open(odXLS.FileName);
    
    XLSAplicacao.Visible := false;
    XLSAplicacao.WorkSheets[1].Activate;

    AbaXLS := XLSAplicacao.Workbooks[ExtractFileName(odXLS.FileName)].WorkSheets[1];
    AbaXLS.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;

    lastRow := XLSAplicacao.ActiveCell.Row;
    lastColumn := XLSAplicacao.ActiveCell.Column;

    RangeMatrix := XLSAplicacao.Range['A1', XLSAplicacao.Cells.Item[lastColumn, lastRow]].Value;

    ShowMessage(IntToStr(lastRow));
    ShowMessage(IntToStr(lastColumn));
    k := 1;
    repeat
      for r := 1 to lastColumn do
        ShowMessage(RangeMatrix[k, r]);
      Inc(k, 1);
    until k > lastColumn;
    
    RangeMatrix := Unassigned;
    
  finally
    
  end;

end;

function TForm1.XlsxToTxt(pFileXLSX: String; pFileTXT: String; pCols: String): PAnsiChar;
var
  SaveAsTXT: TSaveAsTXT;
  vError: Integer;
  vFileXLSX: PAnsiChar;
  vCols: PAnsiChar;
  vFileTXT: PAnsiChar;
begin
  Result := '';

  vFileXLSX := PAnsiChar(AnsiString(pFileXLSX));
  vFileTXT  := PAnsiChar(AnsiString(pFileTXT));
  vCols     := PAnsiChar(AnsiString(pCols));

  Handler := LoadLibrary('xlsx2txt.dll');
  vError := GetLastError;
  try
    if Handler = 0 then
      raise Exception.Create('Erro ao carregar Dll "xlsx2txt.dll". Código: ' + IntToStr(vError));

    @SaveAsTXT := GetProcAddress(Handler, 'SaveAsTXT');
    if @SaveAsTXT <> nil then
        Result := SaveAsTXT(vFileXLSX, vFileTXT, vCols);
  finally
    @SaveAsTXT := nil;
    FreeLibrary(Handle)
  end;
end;

{ TIndiceCols }

procedure TIndiceCols.SetIndice(pIndice: Integer; pCol: String);
const
 vACols: array [1..9] of string = ('DATAPEDIDO',
                                   'DATAPAGAMENTO',
                                   'DATAESTORNO',
                                   'DATALIBERAÇÃO',
                                   'DATAPREVISTAPGTO',
                                   'REF.PEDIDO',
                                   'ENTREGA',
                                   'TIPO',
                                   'VALOR');
begin
  case AnsiIndexStr( AnsiUpperCase(StringReplace(pCol, ' ', '',  [rfReplaceAll])), vACols ) of
    0: DataPedido     := pIndice;
    1: DataPagamento  := pIndice;
    2: DataEstorno    := pIndice;
    3: DataLiberacao  := pIndice;
    4: DataPrevista   := pIndice;
    5: RefPedido      := pIndice;
    6: Entrega        := pIndice;
    7: Tipo           := pIndice;
    8: Valor          := pIndice;
  end;
end;

end.
