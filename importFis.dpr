program importFis;

uses
  Vcl.Forms,
  UMain in 'UMain.pas' {Form1},
  Unit2 in 'Unit2.pas' {Form2},
  uImportExcel in 'uImportExcel.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.Run;
end.
