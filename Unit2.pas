unit Unit2;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, System.StrUtils,
  Vcl.ExtCtrls;

type
  TForm2 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    ComboBox1: TComboBox;
    btnOK: TBitBtn;
    procedure btnOKClick(Sender: TObject);
  private

  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation
uses
UMain;

{$R *.dfm}

procedure TForm2.btnOKClick(Sender: TObject);
begin
  if(ComboBox1.Text ='Selecione...') then
  begin
    ShowMessage('Selecione um campo da tabela');
     ModalResult := mrCancel;
    exit;
  end;
  ModalResult := mrOK;
end;



end.
