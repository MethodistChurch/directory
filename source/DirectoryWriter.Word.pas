unit DirectoryWriter.Word;

interface

uses
  Word2010,
  Directory.Common;

type
  TMemberWriter = class
    const vbTab = #9;
    const vbLf = #11;
    const vbCr = #13;
    const vbCrLf = #13#11;
    const IsDebug = False;

    DefaultFont = 'Segoe UI';
    DefaultFont_FarEast = 'PMingLiU';

    const s_Directory = 'Directory';
    const s_SectionTitle = 'SectionTitle';
    const s_SectionTitleCN = 'SectionTitleCN';
    const s_AreaSeq = 'AreaSeq';
    const s_Area = 'Area';
    const s_Member = 'Member';
    const s_Member_4 = 'Member_4';
    const s_MemberAddress = 'Member Address';
    const s_MemberAddressParagraph = 'Member Address Paragraph';
    const s_MemberTableHeader = 'Member Table Header';

    const TOC_CN = #$76EE + #$5F55;
    const TOC_EN = 'Table of Contents';

    const U_Dot = #$2022;
    const U_Delim = #$3000;
    const U_Gap = U_Delim + U_Dot + U_Delim;
    const U_AreaDelim = #$0020 + U_Dot + #$0020;

    const CN_Name = #$59D3 + #$540D;
    const CN_Email = #$96FB + #$90F5;
    const CN_Phone = #$7535 + #$8BDD;
    const CN_Address = #$5730 + #$5740;

    const CN_ChineseIndex = #$4E2D + #$6587 + #$7D22 + #$5F15;
    const EN_ChineseIndex = 'Chinese Name Index';

    const CN_EnglishIndex = #$82F1 + #$6587 + #$7D22 + #$5F15;
    const EN_EnglishIndex = 'English Name Index';

    const CTrue = -1;
    const CFalse = 0;
  private
    ActiveDocument: WordDocument;
    CurrentSelection: WordSelection;
    WordApp: TWordApplication;
    function NewArea(o: TArea): WordDocument;
    function NewMasterDoc(DocName, aTitleDoc: string; aDocs: TArray<string>):
        WordDocument;
    procedure B03_SetupDocument;
    procedure B04_SetupStyles;
    procedure B05_DrawAreaTitle(Area_EN, Area_CN: string);
    function B06_DrawAreaTable: Table;
    procedure B09_DrawHouse(o: THouse);
    procedure B10_Finalize(aDoc: WordDocument);
    procedure Setup;
    procedure TearDown;
    function _: OleVariant; inline;
    function SplitStringByWord(aStr: string; aMaxLength: Integer = 30):
        TArray<string>;
  public
    procedure Generate(C: TCatalog; Indexes: TArray<Integer>; aWorkingDir: string);
  end;

implementation

uses
  Winapi.Windows, System.Generics.Collections, System.IOUtils, System.SysUtils,
  System.Variants,
  Spring;

procedure TMemberWriter.Setup;
begin
  WordApp := TWordApplication.Create(nil);
  WordApp.Connect;
  WordApp.Visible := True;
end;

procedure TMemberWriter.TearDown;
begin
  WordApp.Disconnect;
  WordApp.Free;
end;

function TMemberWriter.SplitStringByWord(aStr: string; aMaxLength: Integer =
    30): TArray<string>;
begin
  Result := [];

  while aStr.Length > 0 do begin
    var A := aStr.Split([' ']);
    var i := 0;
    while aStr.Length > aMaxLength do begin
      Inc(i);
      aStr := string.Join(' ', Copy(A, 0, Length(A) - i));
    end;
    Result := Result + [aStr];
    aStr := string.Join(' ', Copy(A, Length(A) - i, i));
  end;
end;

function TMemberWriter.NewArea(o: TArea): WordDocument;
begin
  ActiveDocument := WordApp.Documents.Add(_, _, _, _);

  B03_SetupDocument;
  B04_SetupStyles;
  B05_DrawAreaTitle(o.NameEN, o.NameCN);

  var T := B06_DrawAreaTable;

  for var h in o.Houses do begin
    with T.Rows.Add(_) do begin
      Select;
      for var i in [wdBorderTop, wdBorderBottom] do begin
        with Borders.Item(i) do begin
          LineStyle := wdLineStyleSingle;
          LineWidth := wdLineWidth100pt;
          Visible := True;
        end;
      end;
    end;
    B09_DrawHouse(h);
  end;
  T.Rows.Item(1).HeadingFormat := CTrue;

  B10_Finalize(ActiveDocument);
  Result := ActiveDocument;
end;

function TMemberWriter.NewMasterDoc(DocName, aTitleDoc: string; aDocs:
    TArray<string>): WordDocument;
begin
  ActiveDocument := WordApp.Documents.Add(_, _, _, _);
  ActiveDocument.SaveAs2(DocName, _, _, _, _, _, _, _, _, _, _, _, _, _, _, _, _);

  B03_SetupDocument;
  B04_SetupStyles;

  with CurrentSelection do begin
    with Range.Sections.First.Footers.Item(wdHeaderFooterPrimary) do begin
      with PageNumbers do begin
        NumberStyle := wdPageNumberStyleLowercaseRoman;
        RestartNumberingAtSection := False;
        StartingNumber := 1;
        Add(_, _);
      end;
    end;

    Range.Sections.First.PageSetup.TextColumns.SetCount(1);
    Fields.Add(Range, wdFieldIncludeText, Format('"/../%s"', [aTitleDoc]), False);

    MoveUntil('"', _);
    MoveRight(wdCharacter, 1, _);
    Fields.Add(Range, wdFieldFileName, '\p', False);
    EndKey(wdStory, _);

    InsertBreak(wdSectionBreakEvenPage);
    Set_Style(ActiveDocument.Styles.Item(s_SectionTitleCN));
    TypeText(TOC_CN);
    InsertBreak(wdSectionBreakContinuous);
    Range.Sections.First.PageSetup.TextColumns.SetCount(2);
    Range.Sections.First.PageSetup.TextColumns.LineBetween := CTrue;
    Fields.Add(Range, wdFieldTOC, '\o \f C', False);

    InsertBreak(wdSectionBreakNextPage);
    Range.Sections.First.PageSetup.TextColumns.SetCount(1);
    TypeText(TOC_EN);
    Set_Style(ActiveDocument.Styles.Item(s_SectionTitle));

    InsertBreak(wdSectionBreakContinuous);
    Range.Sections.First.PageSetup.TextColumns.SetCount(2);
    Range.Sections.First.PageSetup.TextColumns.LineBetween := CTrue;
    Fields.Add(Range, wdFieldTOC, '\o \f E', False);
  end;

  var bIsFirstDoc := True;
  for var s in aDocs do begin
    with CurrentSelection do begin
      InsertBreak(wdSectionBreakOddPage);
      var curSection := Range.Sections.First;
      curSection.PageSetup.DifferentFirstPageHeaderFooter := CFalse;
      curSection.PageSetup.TextColumns.SetCount(1);
      Fields.Add(Range, wdFieldIncludeText, Format('"/../%s"', [s]), False);

      MoveUntil('"', _);
      MoveRight(wdCharacter, 1, _);
      Fields.Add(Range, wdFieldFileName, '\p', False);
      EndKey(wdStory, _);

      if bIsFirstDoc then begin
        with curSection.Footers.Item(wdHeaderFooterPrimary) do begin
          LinkToPrevious := False;
          with PageNumbers do begin
            NumberStyle := wdPageNumberStyleArabic;
            RestartNumberingAtSection := True;
            StartingNumber := 1;
            Add(_, _);
          end;
        end;
        bIsFirstDoc := False;
      end else begin
        with curSection.Footers.Item(wdHeaderFooterPrimary) do begin
          LinkToPrevious := True;
          with PageNumbers do begin
            RestartNumberingAtSection := False;
          end;
        end;
      end;
    end;
  end;

  with CurrentSelection do begin
    InsertBreak(wdSectionBreakOddPage);
    Set_Style(ActiveDocument.Styles.Item(s_SectionTitleCN));
    Fields.Add(Range, wdFieldTOCEntry, '"A.' + vbTab + CN_ChineseIndex + '" \f C', False);
    Fields.Add(Range, wdFieldTOCEntry, '"A.' + vbTab + EN_ChineseIndex + '" \f E', False);
    TypeText(CN_ChineseIndex);
    Fields.Add(Range, wdFieldIndex, '\c 4 \h "" \f "c"\z "1028"\o "S"', False);
  end;

  with CurrentSelection do begin
    InsertBreak(wdSectionBreakOddPage);
    Set_Style(ActiveDocument.Styles.Item(s_SectionTitle));
    Fields.Add(Range, wdFieldTOCEntry, '"B.' + vbTab + CN_EnglishIndex + '" \f C', False);
    Fields.Add(Range, wdFieldTOCEntry, '"B.' + vbTab + EN_EnglishIndex + '" \f E', False);
    TypeText(EN_EnglishIndex);
    Fields.Add(Range, wdFieldIndex, '\c 3 \h "—A—" \f "e"', False);
  end;

  ActiveDocument.Fields.Update;
  for var i := 1 to ActiveDocument.TablesOfContents.Count do
    ActiveDocument.TablesOfContents.Item(i).Update;

  CurrentSelection.HomeKey(wdStory, _);
  with ActiveDocument.ActiveWindow.View do begin
    ShowFieldCodes := False;
    ShowHiddenText := False;
    type_ := wdPrintView;
    Zoom.PageRows := 1;
    Zoom.PageColumns := 2;
  end;

  Result := ActiveDocument;
end;

procedure TMemberWriter.B03_SetupDocument;
begin
  CurrentSelection := WordApp.Selection;

  // Paper Size and orientation
  with ActiveDocument.PageSetup do begin
    PaperSize     := wdPaperA5;
    Orientation   := wdOrientPortrait;
    TopMargin     := WordApp.MillimetersToPoints(10);
    BottomMargin  := WordApp.MillimetersToPoints(10);
    LeftMargin    := WordApp.MillimetersToPoints(12.7);
    RightMargin   := WordApp.MillimetersToPoints(8);
    MirrorMargins := CTrue;
    OddAndEvenPagesHeaderFooter := CTrue;
    HeaderDistance := TopMargin;
    FooterDistance := BottomMargin - 2;
  end;

  // Show hidden text before render
  ActiveDocument.ActiveWindow.View.ShowHiddenText := True;

  // Show field codes before render
  ActiveDocument.ActiveWindow.View.ShowFieldCodes := True;

  // View Draft
  ActiveDocument.ActiveWindow.View.Draft := True;

  // Hide Spelling and Grammar Errors
  ActiveDocument.ShowGrammaticalErrors := False;
  ActiveDocument.ShowSpellingErrors := False;

  // View Grid Lines
  ActiveDocument.ActiveWindow.View.TableGridlines := False;
end;

procedure TMemberWriter.B04_SetupStyles;
begin
  // Define style: Normal
  with ActiveDocument.Styles.Add(s_Directory, wdStyleTypeParagraph) do begin
    Set_BaseStyle(ActiveDocument.Styles.Item(wdStyleNormal));
    Font.Name := DefaultFont;
    Font.NameFarEast := DefaultFont_FarEast;
    Font.Size := 10;
    ParagraphFormat.LineSpacingRule := wdLineSpaceSingle;
  end;

  // Define style: TOC1
  with ActiveDocument.Styles.Item(wdStyleTOC1) do begin
    Font.Name := DefaultFont;
    Font.NameFarEast := DefaultFont_FarEast;
    Font.Size := 12;
    ParagraphFormat.SpaceAfter := 8;//WordApp.MillimetersToPoints(3);
    with ParagraphFormat.TabStops do begin;
      ClearAll;
      Add(WordApp.MillimetersToPoints(8),    wdAlignTabLeft, _);
      Add(WordApp.MillimetersToPoints(ActiveDocument.PageSetup.TextColumns.Item(1).Width / 2),   wdAlignTabRight, wdTabLeaderDots);
    end;
    ParagraphFormat.TabHangingIndent(1);
  end;

  // Define style: SectionTitle
  with ActiveDocument.Styles.Add(s_SectionTitle, wdStyleTypeParagraph) do begin
    Font.Bold := CTrue;
    Font.Size := 20;
    Set_NextParagraphStyle(ActiveDocument.Styles.Item(s_Directory));
    ParagraphFormat.SpaceAfter := WordApp.MillimetersToPoints(5);
    ParagraphFormat.Alignment := wdAlignParagraphCenter;
  end;

  // Define style: SectionTitleCN
  with ActiveDocument.Styles.Add(s_SectionTitleCN, wdStyleTypeParagraph) do begin
    Font.Name := DefaultFont;
    Font.NameFarEast := DefaultFont_FarEast;
    Font.Bold := CTrue;
    Font.Size := 20;
    Set_NextParagraphStyle(ActiveDocument.Styles.Item(s_Directory));
    ParagraphFormat.SpaceAfter := WordApp.MillimetersToPoints(5);
    ParagraphFormat.Alignment := wdAlignParagraphCenter;
  end;

  // Define style: AreaSeq
  with ActiveDocument.Styles.Add(s_AreaSeq, wdStyleTypeParagraph) do begin
    Set_BaseStyle(ActiveDocument.Styles.Item(s_Directory));
    Set_NextParagraphStyle(ActiveDocument.Styles.Item(s_Directory));
    Font.Bold := CTrue;
    Font.Size := 140;
    Font.ColorIndex := wdBlack;
    ParagraphFormat.SpaceBefore := 0;
    ParagraphFormat.SpaceAfter := 0;
    ParagraphFormat.Alignment := wdAlignParagraphRight;
  end;

  // Define style: Area
  with ActiveDocument.Styles.Add(s_Area, wdStyleTypeParagraph) do begin
    Set_BaseStyle(ActiveDocument.Styles.Item(s_Directory));
    Set_NextParagraphStyle(ActiveDocument.Styles.Item(s_Directory));
    Font.Bold := CTrue;
    Font.Size := 28;
    Font.ColorIndex := wdBlack;
    ParagraphFormat.SpaceBefore := 0;
    ParagraphFormat.SpaceAfter :=  5;
    ParagraphFormat.Alignment := wdAlignParagraphLeft;
    with ParagraphFormat.TabStops do begin;
      ClearAll;
      Add(WordApp.MillimetersToPoints(45),    wdAlignTabRight, _);
      Add(WordApp.MillimetersToPoints(50),    wdAlignTabLeft, _);
      Add(WordApp.MillimetersToPoints(60),    wdAlignTabLeft, _);
    end;
    ParagraphFormat.TabHangingIndent(3);
//    With ParagraphFormat.Borders do begin
//      Enable := CTrue;
//      DistanceFromTop := 2;
//      DistanceFromBottom := 2;
//      DistanceFromLeft := 2;
//      DistanceFromRight := 2;
//    end;
  end;

  // Define style: Member has 3 chinese name characters
  with ActiveDocument.Styles.Add(s_Member, wdStyleTypeParagraph) do begin
    Set_BaseStyle(ActiveDocument.Styles.Item(s_Directory));
    ParagraphFormat.SpaceAfter := 0;
    with ParagraphFormat.TabStops do begin;
      Add(WordApp.MillimetersToPoints(11.5),  wdAlignTabLeft,  _);
      Add(WordApp.MillimetersToPoints(53),    wdAlignTabLeft,  _);
      Add(WordApp.MillimetersToPoints(127.5), wdAlignTabRight, _);
    end;
  end;

  // Define style: Member has 4 chinese name characters
  with ActiveDocument.Styles.Add(s_Member_4, wdStyleTypeParagraph) do begin
    Set_BaseStyle(ActiveDocument.Styles.Item(s_Directory));
    ParagraphFormat.SpaceAfter := 0;
    with ParagraphFormat.TabStops do begin;
      Add(WordApp.MillimetersToPoints(15),    wdAlignTabLeft,  _);
      Add(WordApp.MillimetersToPoints(53),    wdAlignTabLeft,  _);
      Add(WordApp.MillimetersToPoints(127.5), wdAlignTabRight, _);
    end;
  end;

  // Define style: Member Address
  with ActiveDocument.Styles.Add(s_MemberAddress, wdStyleTypeParagraph) do begin
    Set_BaseStyle(ActiveDocument.Styles.Item(s_Directory));
    ParagraphFormat.SpaceAfter := 0;
  end;

  // Define style: Member Address Paragraph
  with ActiveDocument.Styles.Add(s_MemberAddressParagraph, wdStyleTypeParagraph) do begin
    Set_BaseStyle(ActiveDocument.Styles.Item(s_Directory));
    ParagraphFormat.SpaceAfter := 0;
    with ParagraphFormat.TabStops do begin
      Add(WordApp.MillimetersToPoints(11.5), wdAlignTabLeft,  _);
      Add(WordApp.MillimetersToPoints(105),  wdAlignTabRight, _);
      Add(WordApp.MillimetersToPoints(128),  wdAlignTabRight, _);
    end;
  end;

  // Define style: Member Table Header
  with ActiveDocument.Styles.Add(s_MemberTableHeader, wdStyleTypeParagraph) do begin
    Set_BaseStyle(ActiveDocument.Styles.Item(s_Directory));
    Font.Bold := CTrue;
    Font.Size := 10;
    ParagraphFormat.SpaceBefore := 0;
    ParagraphFormat.SpaceAfter := 0;
    with ParagraphFormat.TabStops do begin
      Add(WordApp.MillimetersToPoints(11.5),  wdAlignTabLeft,  _);
      Add(WordApp.MillimetersToPoints(53),    wdAlignTabLeft,  _);
      Add(WordApp.MillimetersToPoints(127.5), wdAlignTabRight, _);
    end;
  end;
end;

procedure TMemberWriter.B05_DrawAreaTitle(Area_EN, Area_CN: string);
begin
  with CurrentSelection do begin
    // Add TOC entry: Chinese
    MoveUp(wdParagraph, 1, _);
    Fields.Add(Range, wdFieldTOCEntry, '".' + vbTab + Area_CN.Replace('/', U_AreaDelim, [rfReplaceAll]) + '" \f C', False);
    MoveUp(wdParagraph, 1, _);
    var F := Fields.Add(Range, wdFieldSequence, 'Area_CN', False);
    F.Cut;
    MoveUntil('.', wdForward);
    Paste;

    // Add TOC entry: English
    MoveUp(wdParagraph, 1, _);
    Fields.Add(Range, wdFieldTOCEntry, '".' + vbTab + Area_EN.Replace('/', U_AreaDelim, [rfReplaceAll]) + '" \f E', False);
    MoveUp(wdParagraph, 1, _);
    F := Fields.Add(Range, wdFieldSequence, 'Area_EN', False);
    F.Cut;
    MoveUntil('.', wdForward);
    Paste;

    MoveDown(wdParagraph, 1, _);
    Set_Style(ActiveDocument.Styles.Item(s_AreaSeq));
    Fields.Add(Range, wdFieldSequence, s_AreaSeq, False);
    TypeText(vbCR);

    // Draw Area Text
    var c := Area_CN.Split(['/']);
    var e := Area_EN.Split(['/']);
    while Length(c) < Length(e) do c := c + [''];

    MoveDown(wdParagraph, 1, _);
    Set_Style(ActiveDocument.Styles.Item(s_Area));
    for var i := Low(e) to High(e) do begin
      TypeText(vbTab + c[i]);
      TypeText(vbTab + U_Dot + vbTab + e[i] + vbCR);
    end;
    CurrentSelection.Set_Style(ActiveDocument.Styles.Item(wdStyleNormal));
    TypeText(vbCR);
    InsertBreak(wdPageBreak);
  end;
end;

function TMemberWriter.B06_DrawAreaTable: Table;
begin
  Result := CurrentSelection.Range.Tables.Add(CurrentSelection.Range, 1, 1, _, _);

  with Result do begin
    PreferredWidthType := wdPreferredWidthPercent;
    PreferredWidth := 100;
    ApplyStyleHeadingRows := True;
    AllowAutoFit := True;
    Rows.AllowBreakAcrossPages := CFalse;
    LeftPadding := 0;
    RightPadding := 0;
    TopPadding := 0;
    BottomPadding := 0;
  end;

  with CurrentSelection do begin
    Set_Style(ActiveDocument.Styles.Item(s_MemberTableHeader));
    InsertAfter(CN_Name + U_Dot + 'Name' + vbTab + CN_Phone + U_Dot + 'Phone' + vbTab + CN_Address + U_Dot + 'Address');
  end;
end;

procedure TMemberWriter.B09_DrawHouse(o: THouse);
var styles: array[Boolean] of Style;
begin
  // Prepare styles
  styles[False] := ActiveDocument.Styles.Item(s_Member);
  styles[True] := ActiveDocument.Styles.Item(s_Member_4);

  // House Addresses
  var Addr := Shared.Make(TQueue<string>.Create);

  Addr.Enqueue(o.Address1);
  var A3 := string.Join(' ',
    string
      .Join(' ', TArray<string>.Create(o.PostCode, o.City))
      .Split([' '], TStringSplitOptions.ExcludeEmpty)
  );

  var A2 := SplitStringByWord(o.Address2);
  case Length(A2) of
    0: Addr.Enqueue(A3);
    1: begin
      Addr.Enqueue(A2[0]);
      Addr.Enqueue(A3);
    end
    else begin
      A2 := A2 + [A3];
      Addr.Enqueue(A2[0]);
      for var s in SplitStringByWord(string.Join(' ', A2, 1, Length(A2) - 1)) do Addr.Enqueue(s)
    end;
  end;

  for var i := 0 to o.Count - 1 do begin
    var M := o.Members[i];
    with CurrentSelection do begin
      Set_Style(styles[GetStringLength(M.NameCN_Display) > 3]);

      if not M.NameCN_Display.IsEmpty then begin
        TypeText(M.NameCN_Display);
        Fields.Add(Range, wdFieldIndexEntry, Format('"%s" \f "c"', [M.NameCN_Display]), False);
      end;

      TypeText(vbTab);

      if not M.NameEN.IsEmpty then begin
        TypeText(M.NameEN);
        Fields.Add(Range, wdFieldIndexEntry, Format('"%s" \f "e"', [M.NameEN]), False);
      end;

      TypeText(vbTab + M.PhoneDisplay);
      if Addr.Count > 0 then
        TypeText(vbTab + Addr.Dequeue);
      if i < o.Count - 1 then
        TypeText(vbCr);
    end;
  end;

  while Addr.Count > 0 do begin
    with CurrentSelection do begin
      TypeText(vbCr);
      Set_Style(styles[False]);
      TypeText(vbTab + vbTab + vbTab + Addr.Dequeue);
    end;
  end;
end;

procedure TMemberWriter.B10_Finalize(aDoc: WordDocument);
begin
  CurrentSelection.HomeKey(wdStory, _);
  with aDoc.ActiveWindow.View do begin
    ShowFieldCodes := False;
    ShowHiddenText := False;
    type_ := wdPrintView;
    Zoom.PageRows := 1;
    Zoom.PageColumns := 2;
  end;
end;

procedure TMemberWriter.Generate(C: TCatalog; Indexes: TArray<Integer>;
    aWorkingDir: string);
begin
  Setup;
  try
    var docNames: TArray<string> := [];
    for var i in Indexes do begin
      var doc := NewArea(C.Areas[i]);

      var docName := Format('%s\%.2d-', [aWorkingDir, i + 1]);
      for var n in C.Areas[i].NameEN do begin
        if TPath.IsValidFileNameChar(n) then
          docName := docName + n
        else
          docName := docName + '_';
      end;

      doc.SaveAs2(docName, _, _, _, _, _, _, _, _, _, _, _, _, _, _, _, _);
      docNames := docNames + [doc.Name];
      doc.Close(_, _, _);
    end;

    docNames := TDirectory.GetFiles(aWorkingDir, '*.docx',
      function(const Path: string; const SearchRec: TSearchRec): Boolean
      begin
        Result := not SameText(SearchRec.Name, 'Directory.docx') and
                  not SameText(SearchRec.Name, 'Title.docx');
      end
    );

    for var i := Low(docNames) to High(docNames) do
      docNames[i] := TPath.GetFileName(docNames[i]);

    var doc := NewMasterDoc(Format('%s\%s', [aWorkingDir, 'Directory']), 'Title.docx', docNames);
    doc.Save;
    doc.Close(_, _, _);
  finally
    TearDown;
  end;
end;

function TMemberWriter._: OleVariant;
begin
  Result := EmptyParam;
end;

end.
