program Clipboard;

uses
  Forms,
  f_prin in 'f_prin.pas' {frmPrin};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmPrin, frmPrin);
  Application.Run;
end.
