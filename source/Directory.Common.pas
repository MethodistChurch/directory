unit Directory.Common;

interface

type
  PHouseMember = ^THouseMember;
  THouseMember = record
  strict private
    FPhone: string;
    FNameCN: string;
  private
    function GetNameCN_Display: string;
    procedure SetPhone(const Value: string);
    function GetPhoneDisplay: string;
    procedure SetNameCN(const Value: string);
  public
    NameEN: string;
    function IsEmpty: Boolean;
    constructor Create(aNameCN, aNameEN, aPhone: string);
    function GetPhoneAsWhatsappLink(out aWhatsappURL: string): Boolean;
    property NameCN: string read FNameCN write SetNameCN;
    property NameCN_Display: string read GetNameCN_Display;
    property Phone: string read FPhone write SetPhone;
    property PhoneDisplay: string read GetPhoneDisplay;
  end;

  PHouse = ^THouse;
  THouse = record
    Address1: string;
    Address2: string;
    PostCode: string;
    City:     string;
    Members:  TArray<THouseMember>;
  public
    class operator Initialize(out Dest: THouse);
    constructor Create(aAddress1, aAddress2, aPostCode, aCity: string; aHouseHead:
        THouseMember);
    function Add(o: THouseMember): PHouseMember;
    function Count: Integer;
    function IsEmpty: Boolean;
  end;

  PArea = ^TArea;
  TArea = record
  public
    NameEN: string;
    NameCN: string;
    Houses: TArray<THouse>;
    class operator Initialize(out Dest: TArea);
    constructor Create(aNameCN, aNameEN: string);
    function Add(o: THouse): PHouse;
    function Count: Integer;
    function IsEmpty: Boolean;
  end;

  TCatalog = record
  private
    FAreas: TArray<TArea>;
  public
    class operator Initialize(out Dest: TCatalog);
    function Add(o: TArea): PArea;
    function Count: Integer;
    property Areas: TArray<TArea> read FAreas;
  end;

function GetStringLength(aStr: string): Integer;

implementation

uses System.SysUtils, System.Character;

function GetStringLength(aStr: string): Integer;
begin
  var i := 0;
  for var c in aStr do begin
    if c.IsSurrogate then
      Inc(i)
    else
      Inc(i, 2);
  end;
  Result := i div 2;
end;

constructor THouseMember.Create(aNameCN, aNameEN, aPhone: string);
begin
  NameCN := aNameCN.Trim;
  NameEN := aNameEN.Trim;
  SetPhone(aPhone);
end;

function THouseMember.GetNameCN_Display: string;
begin
  if GetStringLength(FNameCN) = 2 then
    Result := FNameCN.Insert(1, #$3000)
  else
    Result := FNameCN;
end;

function THouseMember.GetPhoneAsWhatsappLink(out aWhatsappURL: string): Boolean;
begin
  if FPhone.StartsWith('0') then begin
    aWhatsappURL := Format('https://wa.me/6', [FPhone]);
    Result := True;
  end else begin
    aWhatsappURL := '';
    Result := False;
  end;
end;

function THouseMember.GetPhoneDisplay: string;
begin
  if FPhone.StartsWith('01') then begin
    var i := (FPhone.Length - 3) div 2;
    Result := FPhone.Substring(0, 3) + '-'
            + FPhone.Substring(3, i) + ' '
            + FPhone.Substring(3 + i);
  end else if FPhone.StartsWith('0') then begin
    var i := (FPhone.Length - 2) div 2;
    Result := FPhone.Substring(0, 2) + '-'
            + FPhone.Substring(2, i) + ' '
            + FPhone.Substring(2 + i);
  end else if FPhone.StartsWith('+') then begin
    var i := (FPhone.Length - 3) div 2;
    Result := FPhone.Substring(0, 3) + ' '
            + FPhone.Substring(3, i) + ' '
            + FPhone.Substring(3 + i);
  end else begin
    var i := FPhone.Length div 2;
    Result := FPhone.Substring(0, i) + ' '
            + FPhone.Substring(i);
  end;
end;

function THouseMember.IsEmpty: Boolean;
begin
  Result := NameCN.IsEmpty and NameEN.IsEmpty and Phone.IsEmpty;
end;

procedure THouseMember.SetNameCN(const Value: string);
begin
  FNameCN := Value;
end;

procedure THouseMember.SetPhone(const Value: string);
begin
  FPhone := Value.Replace('-', '', [rfReplaceAll])
                 .Replace(' ', '', [rfReplaceAll])
                 .Trim;
end;

function THouse.Add(o: THouseMember): PHouseMember;
begin
  if o.IsEmpty then
    Exit(nil)
  else begin
    Members := Members + [o];
    Result := @Members[High(Members)];
  end;
end;

function THouse.Count: Integer;
begin
  Result := Length(Members);
end;

constructor THouse.Create(aAddress1, aAddress2, aPostCode, aCity: string;
    aHouseHead: THouseMember);
begin
  Address1 := aAddress1.Trim;
  Address2 := aAddress2.Trim;
  PostCode := aPostCode.Trim;
  City := aCity.Trim;
  Add(aHouseHead);
end;

function THouse.IsEmpty: Boolean;
begin
  Result := Count = 0;
end;

class operator THouse.Initialize(out Dest: THouse);
begin
  Dest.Members := [];
end;

function TArea.Add(o: THouse): PHouse;
begin
  Houses := Houses + [o];
  Result := @Houses[High(Houses)];
end;

function TArea.Count: Integer;
begin
  Result := Length(Houses);
end;

constructor TArea.Create(aNameCN, aNameEN: string);
begin
  NameCN := aNameCN.Trim;
  NameEN := aNameEN.Trim;
end;

class operator TArea.Initialize(out Dest: TArea);
begin
  Dest.Houses := [];
end;

function TArea.IsEmpty: Boolean;
begin
  Result := Count = 0;
end;

function TCatalog.Add(o: TArea): PArea;
begin
  FAreas := FAreas + [o];
  Result := @FAreas[High(FAreas)];
end;

function TCatalog.Count: Integer;
begin
  Result := Length(FAreas);
end;

class operator TCatalog.Initialize(out Dest: TCatalog);
begin
  Dest.FAreas := [];
end;

end.
