program DirectoryBuilder;

uses
  Vcl.Forms,
  Directory.Common in '..\source\Directory.Common.pas',
  DirectoryReader.Excel in '..\source\DirectoryReader.Excel.pas',
  DirectoryWriter.Word in '..\source\DirectoryWriter.Word.pas',
  fmMain in '..\source\fmMain.pas' {MainForm};

{$R *.res}

begin
  ReportMemoryLeaksOnShutdown := True;
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
