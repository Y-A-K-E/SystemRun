program SystemRun;



{$R *.dres}

uses
  Vcl.Forms,
  Main in 'Main.pas' {MainForm},
  RunElevatedSupport in 'RunElevatedSupport.pas',
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Metro Blue');
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
