unit fmMain;

interface

uses
  Winapi.Messages, Winapi.Windows, System.Classes, System.SysUtils,
  System.Variants, Vcl.CheckLst, Vcl.Controls, Vcl.Dialogs, Vcl.Forms,
  Vcl.Graphics, Vcl.StdCtrls,
  Directory.Common;

type
  TMainForm = class(TForm)
    Button1: TButton;
    Button2: TButton;
    CheckListBox1: TCheckListBox;
    Memo1: TMemo;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    FCatalog: TCatalog;
    procedure DrawArea(aCatalog: TCatalog);
  public

  end;

var
  MainForm: TMainForm;

implementation

uses
  DirectoryReader.Excel, DirectoryWriter.Word;

{$R *.dfm}

procedure TMainForm.Button1Click(Sender: TObject);
begin
  var F: string;
  if PromptForFileName(F, 'Excel Workbook(*.xlsx)|*.xlsx|All files (*.*)|*.*') then begin
    FCatalog := TMemberReader.Read(F);
    DrawArea(FCatalog);
  end;
end;

procedure TMainForm.Button2Click(Sender: TObject);
begin
  var o := TMemberWriter.Create;
  try
    var Items: TArray<Integer> := [];
    for var i := 0 to CheckListBox1.Count - 1 do
      if CheckListBox1.Checked[i] then
        Items := Items + [i];

    o.Generate(FCatalog, Items);
  finally
    o.Free;
  end;
end;

procedure TMainForm.DrawArea(aCatalog: TCatalog);
begin
  var i := 0;
  for var o in aCatalog.Areas do begin
    var MemberCount := 0;
    for var h in o.Houses do
      Inc(MemberCount, h.Count);
    Inc(i);

    var s := Format('%d. %s, %d, %d', [i, o.NameEN, o.Count, MemberCount]);
    Memo1.Lines.Add(s);

    CheckListBox1.Items.Add(i.ToString + '. ' + o.NameCN + ' ' + o.NameEN);
  end;
  CheckListBox1.CheckAll(TCheckBoxState.cbChecked, False, False);
end;

end.
