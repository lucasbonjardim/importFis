unit UMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.Phys.FB, FireDAC.Phys.FBDef, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, FireDAC.VCLUI.Wait, Data.DB, Vcl.Grids,
  Vcl.DBGrids, Vcl.ExtCtrls, FireDAC.Phys.IBBase, FireDAC.Comp.UI,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client, Vcl.Buttons, uImportExcel,
  Vcl.ComCtrls, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxStyles, dxSkinsCore, dxSkinMetropolis, dxSkinMetropolisDark,
  dxSkinOffice2013DarkGray, dxSkinOffice2013LightGray, dxSkinOffice2013White,
  dxSkinOffice2016Colorful, dxSkinOffice2016Dark, dxSkinVisualStudio2013Blue,
  dxSkinVisualStudio2013Dark, dxSkinVisualStudio2013Light, dxSkinVS2010,
  dxSkinscxPCPainter, cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit,
  cxNavigator, cxDBData, cxGridLevel, cxClasses, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  Data.Win.ADODB, Vcl.ValEdit, Vcl.Menus;

type
  TForm1 = class(TForm)
    FDBcoERP: TFDConnection;
    qryFieldProduto: TFDQuery;
    FDGUIxWaitCursor1: TFDGUIxWaitCursor;
    FDPhysFBDriverLink1: TFDPhysFBDriverLink;
    OpenDialog1: TOpenDialog;
    ProgressBar1: TProgressBar;
    FDQueryProduto: TFDQuery;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    ts_log: TTabSheet;
    PanelUnder: TPanel;
    Panel1: TPanel;
    ScrollBox2: TScrollBox;
    StringGrid1: TStringGrid;
    Panel2: TPanel;
    ValueListEditor1: TValueListEditor;
    pnlTop: TPanel;
    Label2: TLabel;
    Label1: TLabel;
    btnImport: TSpeedButton;
    Label3: TLabel;
    btnLoadExcel: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton1: TSpeedButton;
    edtfdb: TEdit;
    edtExcel: TEdit;
    BitBtn1: TBitBtn;
    btnConectar: TBitBtn;
    Panel3: TPanel;
    Memo1: TMemo;
    Label4: TLabel;
    TabSheet2: TTabSheet;
    memo2: TMemo;
    Panel4: TPanel;
    Label5: TLabel;
    ComboBox1: TComboBox;
    Label6: TLabel;
    CheckBox1: TCheckBox;
    Panel5: TPanel;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    procedure BitBtn1Click(Sender: TObject);
    procedure btnConectarClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure btnLoadExcelClick(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure btnImportClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure StringGrid1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SpeedButton1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure StringGrid1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
  private
    ComboBoxes: array of TComboBox;
    removePontuacao: boolean;
    ColunaSelecionada: Integer;
    importExcel :TImportExcel;
   function loadDB(): Boolean;
    function desconectDB: Boolean;
    function dbConected: Boolean;
    function getFieldList: TStrings;
    procedure CriarComboBoxes;
    procedure CriarComboBoxes1(const Coluna: Integer);
    procedure ComboBoxChange(Sender: TObject);
    procedure ClearValueListEditor(AValueListEditor: TValueListEditor);
    function buscaTipoPlanilha: Integer;
    function RemoverPontuacao(s: string): string;
    function RemoverZerosAEsquerda(const s: string): string;

  public
    FFieldsItems: TStrings;
  end;

var
  Form1: TForm1;

implementation
uses Unit2,  RegularExpressions;


{$R *.dfm}

procedure TForm1.BitBtn1Click(Sender: TObject);
var
 i, j: Integer;
begin
  OpenDialog1.Title      := 'Importar Planilha';
  OpenDialog1.DefaultExt := '.xls';
  OpenDialog1.Filter     := 'Arquivos XLS|*.xls';
  If OpenDialog1.Execute then
  begin
    edtExcel.Text :=     OpenDialog1.FileName;
  end;
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
begin
  ClearValueListEditor(ValueListEditor1);

end;

procedure TForm1.BitBtn3Click(Sender: TObject);
begin
 ValueListEditor1.Strings.Delete(ValueListEditor1.Strings.Count-1);
end;

procedure TForm1.btnConectarClick(Sender: TObject);
var
 i, j: Integer;
begin
  OpenDialog1.Title      := 'Selecione Banco SAT FACIL ';
  OpenDialog1.DefaultExt := '.fdb';
  OpenDialog1.Filter     := 'Arquivos FDB|*.fdb';
  If OpenDialog1.Execute then
  begin
    edtfdb.Text :=     OpenDialog1.FileName;

    if(edtfdb.Text = '') then
    begin
      showMessage('Selecione o banco de dados');
      exit;
    end;
  end;
end;

procedure TForm1.btnImportClick(Sender: TObject);
var
  RowIndex, ColIndex, COD_BARRAS_Index,TotalRows: Integer;
  COD_BARRAS, Campo, Valor, SQLQuery: string;
begin
  if(dbConected) = False then
  begin
    ShowMessage('Faça a importação da planilha!');
    Exit;
  end;

  FDQueryProduto.Close;

  // Encontrar o índice da coluna 'COD_BARRAS'
  COD_BARRAS_Index := -1;
  for ColIndex := 0 to StringGrid1.ColCount - 1 do
  begin
    if StringGrid1.Cells[ColIndex, 0] = 'COD_BARRA' then
    begin
      COD_BARRAS_Index := ColIndex;
      Break;
    end;
  end;

  if COD_BARRAS_Index = -1 then
  begin
    ShowMessage('Coluna COD_BARRA não encontrada na StringGrid.');
    Exit;
  end;

    // Definir o máximo para o ProgressBar como o número total de linhas
  ProgressBar1.Max := StringGrid1.RowCount - 1;
  TotalRows := ProgressBar1.Max;

  for RowIndex := 1 to StringGrid1.RowCount - 1 do
  begin
    COD_BARRAS := StringGrid1.Cells[COD_BARRAS_Index, RowIndex];


    if COD_BARRAS = 'SEM GTIN' then
    begin
      Continue;
    end;

    if COD_BARRAS = '' then
    begin
      Continue;
    end;


    SQLQuery := 'UPDATE produtos SET ';

    // Percorre todas as colunas exceto a 'COD_BARRAS'
    for ColIndex := 0 to High(ComboBoxes) do
    begin
      Campo := ComboBoxes[ColIndex].Text;

      // Verifica se o campo é 'Selecione...'
      if Campo = 'Selecione...' then
        Continue; // Pula para a próxima iteração

      Valor := StringGrid1.Cells[ColIndex, RowIndex];


      if Valor = 'SEM GTIN' then
        Continue; // Pula para a próxima iteração

      if Valor = '' then
        Continue; // Pula para a próxima iteração

      // Adiciona cada atualização à string SQL
      SQLQuery := SQLQuery + Campo + ' = ' + QuotedStr(Valor) + ', ';
    end;

    // Remove a vírgula extra no final
    Delete(SQLQuery, Length(SQLQuery) - 1, 2);

    // Adiciona a cláusula WHERE
    SQLQuery := SQLQuery + ' WHERE CODIGO_BARRAS = ' + QuotedStr(COD_BARRAS);

    try
      FDQueryProduto.SQL.Text := SQLQuery;
      FDQueryProduto.ExecSQL;
    except
      on E: Exception do
        ShowMessage('Erro ao atualizar produto com COD_BARRAS = ' + COD_BARRAS +
          ': ' + E.Message);
    end;

    ProgressBar1.Position := RowIndex;
    Application.ProcessMessages; // Atualiza a interface gráfica
  end;
   ProgressBar1.Position := 0;
  ShowMessage('Finalizado!');
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  edtfdb.Text := '';
  edtExcel.Text :='';
end;

procedure TForm1.FormResize(Sender: TObject);
begin
   Application.ProcessMessages;
end;


procedure TForm1.FormShow(Sender: TObject);
begin
  FDBcoERP.Connected := false;
end;

procedure TForm1.btnLoadExcelClick(Sender: TObject);
var
  i, j: Integer;
begin
  if (edtExcel.Text = '') then
  begin
    ShowMessage('Selecione uma planilha!');
    Exit;
  end;

  if (edtfdb.Text = '') then
  begin
    ShowMessage('Selecione banco de dados!');
    Exit;
  end;

  if (loadDB) then
  begin
    FFieldsItems := getFieldList;

    try
      StringGrid1.RowCount := 1;
      importExcel := TImportExcel.Create(Self);
      importExcel.ExcelFile := edtExcel.Text;
      importExcel.ExcelParaStringGrid(StringGrid1, ProgressBar1);
    finally
    end;
  end;
end;

procedure TForm1.CriarComboBoxes;
var
  j: Integer;
begin
  // Limpar ComboBoxes existentes antes de criar novos
  for j := 0 to Length(ComboBoxes) - 1 do
    FreeAndNil(ComboBoxes[j]);

  // Adicionar ComboBoxes à StringGrid
  SetLength(ComboBoxes, StringGrid1.ColCount);

  for j := 0 to StringGrid1.ColCount - 1 do
  begin
    ComboBoxes[j] := TComboBox.Create(ScrollBox2);
    ComboBoxes[j].Parent := ScrollBox2;
    ComboBoxes[j].Items.Add('Selecione...');
    if Assigned(FFieldsItems) then
      ComboBoxes[j].Items.AddStrings(FFieldsItems);

    // Posicionar o ComboBox acima da coluna correspondente
    ComboBoxes[j].Top := ScrollBox2.Top - ComboBoxes[j].Height;
    ComboBoxes[j].Left := StringGrid1.CellRect(j, 0).Left;
    ComboBoxes[j].Width := StringGrid1.ColWidths[j];
    ComboBoxes[j].Style := csDropDownList;
    ComboBoxes[j].ItemIndex := 0;
    ComboBoxes[j].OnChange := ComboBoxChange; // Adiciona um evento OnChange para processar as seleções
  end;

    ValueListEditor1.Strings.Clear;
  for j := 0 to StringGrid1.ColCount - 1 do
  begin
    ValueListEditor1.InsertRow(StringGrid1.Cells[j, 0], ComboBoxes[j].Items.CommaText, True);
  end;

end;

procedure TForm1.CriarComboBoxes1(const Coluna: Integer);
var
  ComboBox: TComboBox;
begin
  // Certificar-se de que o array ComboBoxes está inicializado
  if Length(ComboBoxes) <> StringGrid1.ColCount then
    SetLength(ComboBoxes, StringGrid1.ColCount);

  // Verificar se o ComboBox já existe na coluna clicada
  if Assigned(ComboBoxes[Coluna]) then
    Exit; // Se já existe, não criar outro

  // Criar ComboBox
  ComboBoxes[Coluna] := TComboBox.Create(Self);
  ComboBox := ComboBoxes[Coluna];
  ComboBox.Parent := Self;  // Anexar ao formulário
  ComboBox.Items.Add('Selecione...');
  if Assigned(FFieldsItems) then
    ComboBox.Items.AddStrings(FFieldsItems);

  // Posicionar o ComboBox acima da coluna correspondente
  ComboBox.Top := StringGrid1.Top - ComboBox.Height;
  ComboBox.Left := StringGrid1.Left + StringGrid1.CellRect(Coluna, 0).Left;
  ComboBox.Width := StringGrid1.ColWidths[Coluna];
  ComboBox.Style := csDropDownList;
  ComboBox.ItemIndex := 0;
end;



function TForm1.LoadDB: Boolean;
begin
  try
    FDBcoERP.Params.Database := edtFdb.Text;
    FDBcoERP.Connected := True;
    Result := FDBcoERP.Connected;
  except
    on E: Exception do
    begin
      ShowMessage('Erro ao conectar ao banco de dados: ' + E.Message);
      Result := False;
    end;
  end;
end;

function TForm1.getFieldList(): TStrings;
var
FFieldsItems : TStrings;
  field: TField;
begin
  FFieldsItems := TStringList.Create;
  qryFieldProduto.Active := False;
  qryFieldProduto.Active := True;
  qryFieldProduto.First;

  while not qryFieldProduto.Eof do
  begin
    for field in qryFieldProduto.Fields do
      FFieldsItems.Add(field.AsString);
    qryFieldProduto.Next;
  end;

  Result :=  FFieldsItems;
end;

function TForm1.desconectDB: Boolean;
begin
  try
    FDBcoERP.Connected := false;
    Result := not FDBcoERP.Connected;
  except
    on E: Exception do
    begin
      ShowMessage('Erro ao desconectar ao banco de dados: ' + E.Message);
      Result := False;
    end;
  end;
end;
function TForm1.dbConected: Boolean;
begin
  Result:= FDBcoERP.Connected ;
end;

function TForm1.buscaTipoPlanilha: Integer;
var i,LINHA_INDEX: Integer;
begin
  // Encontrar a linha da grid que tem campo cod barra   ou cod.
  LINHA_INDEX := -1;
  for i := 0 to StringGrid1.ColCount - 1 do
  begin
    if StringGrid1.Cells[i, 0] = 'COD_BARRA' then
    begin
      LINHA_INDEX := 0;
      Break;
    end;

    if StringGrid1.Cells[i, 0] = 'COD' then
    begin
      LINHA_INDEX := 0;
      Break;
    end;

    if StringGrid1.Cells[i, 0] = 'EAN' then
    begin
      LINHA_INDEX := 0;
      Break;
    end;
  end;

  if(LINHA_INDEX = -1) then
  begin
    for i := 0 to StringGrid1.ColCount - 1 do
    begin
      if StringGrid1.Cells[i, 2] = 'EAN' then
      begin
        LINHA_INDEX := 2;
        Break;
      end;
    end;
  end;

  Result :=  LINHA_INDEX;
end;

function TForm1.RemoverPontuacao(s: string): string;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length(s) do
  begin
    if s[i] in ['0'..'9'] then
      Result := Result + s[i];
  end;
end;

function TForm1.RemoverZerosAEsquerda(const s: string): string;
var
  i: Integer;
begin
  i := 1;
  // Encontrar o primeiro caractere não zero
  while (i <= Length(s)) and (s[i] = '0') do
    Inc(i);

  // Retornar a parte da string após os zeros à esquerda
  Result := Copy(s, i, Length(s));
end;



procedure TForm1.SpeedButton1Click(Sender: TObject);
var
  RowIndex, ColIndex, COD_BARRAS_Index,TotalRows,i,j, ColKey_Index,tipoPlanilha: Integer;
  COD_BARRAS, VALUE_KEY, Campo, FIELD, SQLQuery, campoWhere: string;
  errorCount , sucessCount: Integer;
begin
  Memo1.Clear;
  memo2.Clear;
  errorCount :=0;
  sucessCount :=0;
  if(dbConected) = False then
  begin
    ShowMessage('Faça a importação da planilha!');
    Exit;
  end;

  if(ComboBox1.ItemIndex=0)then
  begin
   ShowMessage('Selecione campo na planilha para a condição WHERE.!');
    Exit;
  end;

   if ValueListEditor1.Strings.Count = 0 then
   begin
    ShowMessage('Selecione pelo menos 1 campo da planilha!');
    Exit;
   end;

  FDQueryProduto.Close;


  tipoPlanilha :=    buscaTipoPlanilha();


  if tipoPlanilha = -1 then
  begin
    ShowMessage('Planilha nao identificada. não é possivel determinal quais elementos compoe colunas de condição where. ');
    Exit;
  end;

  // Encontrar o índice da coluna 'COD_BARRAS'
  COD_BARRAS_Index := -1;
  for ColIndex := 0 to StringGrid1.ColCount - 1 do
  begin
    if StringGrid1.Cells[ColIndex, tipoPlanilha] = ComboBox1.Text then
    begin
      COD_BARRAS_Index := ColIndex;
      Break;
    end;
  end;


  if COD_BARRAS_Index = -1 then
  begin
    ShowMessage('Coluna (' +  ComboBox1.Text+') para update, condição WHERE não encontrada na StringGrid.');
    Exit;
  end;

    // Definir o máximo para o ProgressBar como o número total de linhas
  ProgressBar1.Max := StringGrid1.RowCount - 1;
  TotalRows := ProgressBar1.Max;

  for RowIndex := 1 to StringGrid1.RowCount - 1 do
  begin
    COD_BARRAS := StringGrid1.Cells[COD_BARRAS_Index, RowIndex];
    COD_BARRAS := RemoverPontuacao(COD_BARRAS);
    COD_BARRAS := RemoverZerosAEsquerda(COD_BARRAS);

    if COD_BARRAS = 'SEM GTIN' then
    begin
      errorCount :=errorCount+1;
      Memo1.Lines.Add('ERRO: Codigo barra: ' + 'SEM GTIN' )  ;
      Continue;
    end;

    if COD_BARRAS = 'EAN' then
    begin
      errorCount :=errorCount+1;
      Memo1.Lines.Add('ERRO: Codigo barra: ' + COD_BARRAS + '= EAN' )  ;
      Continue;
    end;

    if COD_BARRAS = 'COD_BARRA' then
    begin
      errorCount :=errorCount+1;
      Memo1.Lines.Add('ERRO: Codigo barra: ' + COD_BARRAS + '= COD_BARRA' )  ;
      Continue;
    end;


    if COD_BARRAS = '' then
    begin
      errorCount :=errorCount+1;
      Memo1.Lines.Add('ERRO: Codigo barra: ' + COD_BARRAS + ' VAZIO.' )  ;
      Continue;
    end;

    j := 1;
   while j <= ValueListEditor1.RowCount - 1 do
    begin
      ProgressBar1.Position := RowIndex;
      Application.ProcessMessages;
      Campo :=  ValueListEditor1.Keys[j];

      // Encontrar o índice da coluna

      ColKey_Index := -1;
      for ColIndex := 0 to StringGrid1.ColCount - 1 do
      begin
        if StringGrid1.Cells[ColIndex, tipoPlanilha] = Campo then
        begin
          ColKey_Index := ColIndex;
          Break;
        end;
      end;

      VALUE_KEY := StringGrid1.Cells[ColKey_Index, RowIndex];
      VALUE_KEY := VALUE_KEY.Replace(',','.');
      FIELD := ValueListEditor1.Values[Campo];

      if(CheckBox1.Checked)then
       begin
        VALUE_KEY := RemoverPontuacao(VALUE_KEY);
       end;

      SQLQuery := 'UPDATE produtos SET ';
      SQLQuery := SQLQuery + (FIELD)  + ' = ' + QuotedStr(VALUE_KEY) + ', ';


      // Remove a vírgula extra no final
      Delete(SQLQuery, Length(SQLQuery) - 1, 2);


      if (ComboBox1.Text = 'COD_ITEM') or (ComboBox1.Text = 'COD.') or (ComboBox1.Text = 'COD') then
      begin
        campoWhere := 'CODIGO';
      end
      else
      begin
        campoWhere := 'CODIGO_BARRAS';
      end;



      SQLQuery := SQLQuery + ' WHERE  '+  campoWhere + ' = '+  QuotedStr(COD_BARRAS);
      memo2.Lines.Add(SQLQuery);
      Label5.Caption :=  SQLQuery;

      try
        FDQueryProduto.SQL.Text := SQLQuery;
        FDQueryProduto.ExecSQL;
        sucessCount :=sucessCount +1;
      except
      on E: Exception do
        begin
          errorCount := errorCount+1;
          Memo1.Lines.Add('Erro ao atualizar produto com Clausula WHERE: = ' + COD_BARRAS + ': ' + E.Message )  ;
        end;

      end;
       Inc(j);
    end;
  end;
   ProgressBar1.Position := 0;
   ClearValueListEditor(ValueListEditor1);
   Label5.Caption := '';

  if(Memo1.Lines.Count>0) then
  begin
    ShowMessage('Finalizado com '+ IntToStr(errorCount)+' erros e ' + IntToStr(sucessCount) +' sucessos. verifique LOG!');
    PageControl1.ActivePage :=ts_log;
  end else
  begin
    ShowMessage('Finalizado sem nenhum erro! '   + IntToStr(sucessCount) +' itens processados!');
  end;


end;

procedure TForm1.ClearValueListEditor(AValueListEditor: TValueListEditor);
var
  i: Integer;
begin
  // Itera sobre os itens e os remove
  for i := AValueListEditor.Strings.Count - 1 downto 0 do
    AValueListEditor.Strings.Delete(i);
end;



procedure TForm1.SpeedButton2Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TForm1.StringGrid1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  ColunaClicada: Integer;
  Form2: TForm2;
    Chave: string;
  Indice: Integer;
begin
  // Obtém a coluna clicada
  ColunaClicada := StringGrid1.Col;

  // Verifica se a coluna clicada é válida
  if ColunaClicada >= 0 then
  begin
    Chave := StringGrid1.Cells[ColunaClicada, 0];

    if(Chave ='') then
    begin
      Chave := StringGrid1.Cells[ColunaClicada, 2];
    end;

    if(Chave ='') then
    begin
      ShowMessage('Nenhuma coluna selecionada');
      exit;
    end;
    // Cria o formulário secundário
    Form2 := TForm2.Create(Self);
    try
      // Define o nome da coluna no ComboBox
      Form2.ComboBox1.Items.AddStrings(FFieldsItems);
      Form2.Panel1.Caption :='COLUNA: ' +  Chave;

      // Mostra o formulário secundário
      if Form2.ShowModal = mrOK then
      begin
          // Inicializa o índice como -1, indicando que a chave não foi encontrada
          Indice := -1;
          // Procura manualmente pela chave no ValueListEditor1.Keys
          for Indice := 0 to ValueListEditor1.ColCount - 1 do
          begin
            if ValueListEditor1.Keys[Indice] = Chave then
              Break; // Chave encontrada, interrompe o loop
          end;

       if Indice < ValueListEditor1.ColCount then
        begin
          // Se a chave já existe, substitui o valor
          ValueListEditor1.Values[Chave] := Form2.ComboBox1.Text;
        end
        else
        begin
          // Se a chave não existe, adiciona uma nova linha
          ValueListEditor1.InsertRow(Chave, Form2.ComboBox1.Text, True);
        end;

      end;
    finally
      // Libera a memória do formulário secundário
      Form2.Free;
    end;
  end;
end;

procedure TForm1.StringGrid1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var
  Col, Row: Integer;
begin
  // Converte as coordenadas do mouse para coluna e linha
  StringGrid1.MouseToCell(X, Y, Col, Row);

  // Verifica se a coluna e a linha são válidas
  if (Col >= 0) and (Row >= 0) and (Col < StringGrid1.ColCount) and (Row < StringGrid1.RowCount) then
  begin
    // Define o hint com base nas informações da célula
    StringGrid1.Hint := 'Coluna: ' + IntToStr(Col) + ', Linha: ' + IntToStr(Row) +
      ', Valor: ' + StringGrid1.Cells[Col, Row];

    // Ativa o hint
    Application.ActivateHint(Mouse.CursorPos);
  end
  else
  begin
    // Limpa o hint se o mouse estiver fora da grade
    StringGrid1.Hint := '';
  end;
end;

procedure TForm1.ComboBoxChange(Sender: TObject);
var
  CampoSelecionado: string;
begin
  // Processa a seleção do ComboBox e atualiza o ValueListEditor
  if Sender is TComboBox then
  begin
    CampoSelecionado := TComboBox(Sender).Text;

    // Atualiza o ValueListEditor com a chave (nome da coluna) e valor (campo selecionado)
    ValueListEditor1.Values[StringGrid1.Cells[ColunaSelecionada, 0]] := CampoSelecionado;

    // Esconde o ComboBox após a seleção
    TComboBox(Sender).Visible := False;
  end;
end;

end.
