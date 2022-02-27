unit DirectoryReader.Excel;

interface

uses
  Excel2010,
  Directory.Common;

type
  // Reference: http://delphiprogrammingdiary.blogspot.com/2018/02/excel-automation-in-delphi.html
  TMemberReader = class
  const BeginCol  = 'A';
  const AreaCN    = 'A';
  const AreaEN    = 'B';
  const NameCN    = 'A';
  const NameEN    = 'B';
  const Phone     = 'C';
  const Address1  = 'D';
  const Address2  = 'E';
  const PostCode  = 'F';
  const City      = 'G';
  const EndCol    = City;
  type
    TRowKind = (rk_None, rk_Area, rk_HouseHead, rk_HouseMember);
  private
    FApp: TExcelApplication;
    FWorkBook: ExcelWorkbook;
    FWorkSheet: ExcelWorksheet;
    LCID: Cardinal;
    function _: OleVariant; inline;
    function GetRowKind(Row: ExcelRange): TRowKind;
    procedure Log(s: string);
    function Process: TCatalog;
  public
    constructor Create(aFileName: string);
    destructor Destroy; override;
    class function Read(aFileName: string): TCatalog;
  end;

implementation

uses
  Winapi.Windows, System.IOUtils, System.SysUtils, System.Variants;

constructor TMemberReader.Create(aFileName: string);
begin
  LCID := GetUserDefaultLCID;

  FApp := TExcelApplication.Create(nil);
  FApp.Connect;
  FApp.Visible[LCID] := True;
  FWorkBook := FApp.Workbooks.Open(aFileName, _, _, _, _, _, _, _, _, _, _, _, _, _, _, LCID);
  FWorkSheet := FWorkBook.Worksheets[1] as ExcelWorksheet;
  FWorkSheet.Activate(LCID);
end;

destructor TMemberReader.Destroy;
begin
  FWorkBook.Close(False, _, _, LCID);
  FApp.Disconnect;
  FApp.Quit;
  FApp.Free;
  inherited;
end;

function TMemberReader.GetRowKind(Row: ExcelRange): TRowKind;
var s: string;
begin
  for var i := 1 to Row.Columns.Count do begin
    s := Row.Item[1, i];
    if not s.Trim.IsEmpty then
      Break
    else if i = Row.Columns.Count then
      Exit(rk_none);
  end;

  if (Row.Range[AreaCN + '1', _].Font.Bold) and (Row.Range[AreaEN + '1', _].Font.Bold) then Exit(rk_Area);

  s := Row.Item[1, Address1];
  if s.IsEmpty then Exit(rk_HouseMember);

  Exit(rk_HouseHead)
end;

procedure TMemberReader.Log(s: string);
begin
//  TFile.AppendAllText(TPath.GetTempPath + 'debug.txt', s + sLineBreak);
end;

function TMemberReader.Process: TCatalog;
begin
  var id: string;
  var Area: TArea;
  var CurrentArea: PArea;
  var iEmptyRows: Integer := 0;
  var CurrentRow: ExcelRange;
  var CurrentHouse: PHouse;

  CurrentRow := FWorksheet.Range[BeginCol + 2.ToString, EndCol + 2.ToString];

  while iEmptyRows < 10 do begin
    var rowKind := GetRowKind(CurrentRow);

    if rowKind = rk_None then
      Inc(iEmptyRows)
    else begin
      iEmptyRows := 0;

      if rowKind = rk_Area then begin
        CurrentArea := Result.Add(TArea.Create(CurrentRow.Item[1, AreaCN], CurrentRow.Item[1, AreaEN]));
        Log(CurrentArea.NameEN + CurrentArea.NameCN);
      end else begin
        var m := THouseMember.Create(CurrentRow.Item[1, NameCN], CurrentRow.Item[1, NameEN], CurrentRow.Item[1, Phone]);
        if rowKind = rk_HouseHead then begin
          var h := THouse.Create(CurrentRow.Item[1, Address1], CurrentRow.Item[1, Address2], CurrentRow.Item[1, PostCode], CurrentRow.Item[1, City], m);
          CurrentHouse := CurrentArea.Add(h);
          Log(m.NameCN + m.NameEN + CurrentHouse.Address1);
        end else if rowKind = rk_HouseMember then begin
          CurrentHouse.Add(m);
          Log(m.NameCN + m.NameEN + m.Phone);
        end;
      end;
    end;

    CurrentRow := CurrentRow.Offset[1, 0];
  end;
end;

class function TMemberReader.Read(aFileName: string): TCatalog;
begin
  var o := Create(aFileName);
  try
    Result := o.Process;
  finally
    o.Free;
  end;
end;

function TMemberReader._: OleVariant;
begin
  Result := EmptyParam;
end;

end.
