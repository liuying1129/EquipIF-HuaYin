program ProHuaYin;

uses
  Forms,
  UfrmMain in 'UfrmMain.pas' {frmMain},
  UCommFunction in 'UCommFunction.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.ShowMainForm:=false;
  Application.Run;
end.
