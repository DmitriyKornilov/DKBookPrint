unit UUtils;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Graphics, fpspreadsheetgrid, fpstypes,
  DK_Fonts, DK_SheetWriter, DK_Vector, DK_Matrix;

const
  COLOR_LIGHTBLUE  = $00FBDEBB;
  COLOR_GRAY       = $00D6D6D6;
  COLOR_GREEN      = $00CCE3CC;
  PAGES_PER_SHEET = 4; //for А4 sheet

type

  { TPagesSheet }

  TPagesSheet = class (TObject)
  private
    FGrid: TsWorksheetGrid;
    FWriter: TSheetWriter;
    FFontName: String;
    FFontSize: Single;
    const
      DATA_COL_WIDTH = 50;
      EMPTY_COL_WIDTH = 25;
      DATA_ROW_HEIGHT = 45;
      EMPTY_ROW_HEIGHT = 10;
      FIRST_COL = 1;
  public
    constructor Create(const AGrid: TsWorksheetGrid);
    destructor  Destroy; override;
    procedure Draw(const AFacePages1, AFacePages2, ABackPages1, ABackPages2: TIntVector;
                   const AColumnName1, AColumnName2, AColumnName3: String);
                   //    STR_SIDE+' 1', STR_SIDE+' 2', STR_SHEET_NUMBER
  end;

  function PageNumberVectorsToString(const V1, V2: TIntVector): String;

  function PageNumbersDocToBooks(const APagesPerBook: Integer;
                                 const ADocPageNumbers: TIntVector): TIntMatrix;
  procedure PlacePageNumbersOnSheets(const APageNumbers: TIntVector;
                        out AFacePageNumbersLeft, AFacePageNumbersRight,
                        ABackPageNumbersLeft, ABackPageNumbersRight: TIntVector);
  procedure PlacePageNumbersOnSheets(const APageNumbers: TIntMatrix;
                        out AFacePageNumbersLeft, AFacePageNumbersRight,
                        ABackPageNumbersLeft, ABackPageNumbersRight: TIntMatrix);

  {LargerMultiple - возвращает ближайшее знаение >= AValue, кратное числу AMultiple}
  function LargerMultiple(const AValue, AMultiple: Integer): Integer;

  {LessMultiple - возвращает ближайшее знаение <= AValue, кратное числу AMultiple}
  function LessMultiple(const AValue, AMultiple: Integer): Integer;

  function IsRangesItersect(const AMin1, AMax1, AMin2, AMax2: Integer): Boolean;

implementation

function LargerMultiple(const AValue, AMultiple: Integer): Integer;
var
  i: Integer;
begin
  for i:= 0 to AMultiple-1 do
  begin
    Result:= AValue + i;
    if (Result mod AMultiple)=0 then break;
  end;
end;

function LessMultiple(const AValue, AMultiple: Integer): Integer;
var
  i: Integer;
begin
  for i:= 0 to AMultiple-1 do
  begin
    Result:= AValue - i;
    if (Result mod AMultiple)=0 then break;
  end;
end;

function IsRangesItersect(const AMin1, AMax1, AMin2, AMax2: Integer): Boolean;
begin
  Result:= ((AMin1<=AMin2) AND (AMax1>=AMax2)) OR
           ((AMin1<=AMin2) AND (AMax1>=AMin2)) OR
           ((AMin1<=AMax2) AND (AMax1>=AMax2)) OR
           ((AMin1>=AMin2) AND (AMax1<=AMax2));
end;

{ TPagesSheet }

constructor TPagesSheet.Create(const AGrid: TsWorksheetGrid);
var
  ColWidths: TIntVector;
begin
  LoadFontFromControl(AGrid, FFontName, FFontSize);
  FGrid:= AGrid;
  ColWidths:=  nil;
  VDim(ColWidths, 7+FIRST_COL, DATA_COL_WIDTH);
  ColWidths[0]:= EMPTY_COL_WIDTH;
  ColWidths[2+FIRST_COL]:= EMPTY_COL_WIDTH;
  ColWidths[5+FIRST_COL]:= EMPTY_COL_WIDTH;
  ColWidths[6+FIRST_COL]:= 2*DATA_COL_WIDTH;
  FWriter:= TSheetWriter.Create(ColWidths, FGrid.Worksheet, FGrid);
end;

destructor TPagesSheet.Destroy;
begin
  if Assigned(FWriter) then FreeAndNil(FWriter);
  inherited Destroy;
end;

procedure TPagesSheet.Draw(const AFacePages1, AFacePages2, ABackPages1, ABackPages2: TIntVector;
  const AColumnName1, AColumnName2, AColumnName3: String);
var
  i, R: Integer;

  procedure SetColor(const ANum: Integer);
  begin
    if ANum=0 then
      FWriter.SetBackground(COLOR_GRAY)
    else
      FWriter.SetBackground(COLOR_GREEN);
  end;

begin
  FWriter.BeginEdit;

  FWriter.SetAlignment(haCenter, vaCenter);
  FWriter.SetFont(FFontName, FFontSize, [], clWindowText);
  R:= 1;
  FWriter.WriteText(R,1+FIRST_COL,R,2+FIRST_COL, AColumnName1, cbtNone, True, True);
  FWriter.WriteText(R,4+FIRST_COL,R,5+FIRST_COL, AColumnName2, cbtNone, True, True);
  FWriter.WriteText(R,7+FIRST_COL, AColumnName3, cbtNone, True, True);

  for i:= 0 to High(AFacePages1) do
  begin
    R:= R + 1;
    FWriter.WriteText(R,1+FIRST_COL,R,5+FIRST_COL, EmptyStr);
    FWriter.SetRowHeight(R, EMPTY_ROW_HEIGHT);
    R:= R + 1;
    SetColor(AFacePages1[i]);
    FWriter.WriteNumber(R, 1+FIRST_COL, AFacePages1[i], cbtOuter);
    SetColor(AFacePages2[i]);
    FWriter.WriteNumber(R, 2+FIRST_COL, AFacePages2[i], cbtOuter);
    SetColor(ABackPages1[i]);
    FWriter.WriteNumber(R, 4+FIRST_COL, ABackPages1[i], cbtOuter);
    SetColor(ABackPages2[i]);
    FWriter.WriteNumber(R, 5+FIRST_COL, ABackPages2[i], cbtOuter);
    FWriter.SetBackgroundClear;
    FWriter.WriteNumber(R, 7+FIRST_COL, i+1, cbtOuter);
    FWriter.SetRowHeight(R, DATA_ROW_HEIGHT);
  end;

  FWriter.SetFrozenRows(1);

  FWriter.EndEdit;
end;

function PageNumberValuesToString(const P1,P2: Integer): String;
begin
  Result:= IntToStr(P1) + ',' + IntToStr(P2);
end;

function PageNumberVectorsToString(const V1, V2: TIntVector): String;
var
  i: Integer;
begin
  Result:= EmptyStr;
  if Length(V1)<>Length(V2) then Exit;
  if VIsNil(V1) then Exit;
  Result:= PageNumberValuesToString(V1[0], V2[0]);
  for i:= 1 to High(V1) do
    Result:= Result + ',' + PageNumberValuesToString(V1[i], V2[i]);
end;

function PageNumbersDocToBooks(const APagesPerBook: Integer;
                               const ADocPageNumbers: TIntVector): TIntMatrix;
var
  i, BooksCount: Integer;
  V: TIntVector;
begin
  PageNumbersDocToBooks:= nil;
  if APagesPerBook=0 then
    MAppend(PageNumbersDocToBooks, ADocPageNumbers)
  else begin
    BooksCount:= Length(ADocPageNumbers) div APagesPerBook;
    for i:=0 to BooksCount-1 do
    begin
      V:= VCut(ADocPageNumbers, i*APagesPerBook, (i+1)*APagesPerBook - 1);
      MAppend(PageNumbersDocToBooks, V);
    end;
  end;
end;



procedure PlacePageNumbersOnSheets(const APageNumbers: TIntMatrix;
                        out AFacePageNumbersLeft, AFacePageNumbersRight,
                        ABackPageNumbersLeft, ABackPageNumbersRight: TIntMatrix);
var
  FacePageNumbersLeft, FacePageNumbersRight,
  BackPageNumbersLeft, BackPageNumbersRight: TIntVector;
  i: Integer;
begin
  AFacePageNumbersLeft:= nil;
  AFacePageNumbersRight:= nil;
  ABackPageNumbersLeft:= nil;
  ABackPageNumbersRight:= nil;
  for i:= 0 to High(APageNumbers) do
  begin
    PlacePageNumbersOnSheets(APageNumbers[i], FacePageNumbersLeft, FacePageNumbersRight,
                             BackPageNumbersLeft, BackPageNumbersRight);
    MAppend(AFacePageNumbersLeft,  FacePageNumbersLeft);
    MAppend(AFacePageNumbersRight, FacePageNumbersRight);
    MAppend(ABackPageNumbersLeft,  BackPageNumbersLeft);
    MAppend(ABackPageNumbersRight, BackPageNumbersRight);
  end;
end;

procedure PlacePageNumbersOnSheets(const APageNumbers: TIntVector;
                        out AFacePageNumbersLeft, AFacePageNumbersRight,
                        ABackPageNumbersLeft, ABackPageNumbersRight: TIntVector);
var
  i,n, SheetsCount: Integer;
begin
  AFacePageNumbersLeft:= nil;
  AFacePageNumbersRight:= nil;
  ABackPageNumbersLeft:= nil;
  ABackPageNumbersRight:= nil;
  SheetsCount:= Length(APageNumbers) div PAGES_PER_SHEET;
  VDim(AFacePageNumbersLeft,  SheetsCount, 0);
  VDim(AFacePageNumbersRight, SheetsCount, 0);
  VDim(ABackPageNumbersLeft,  SheetsCount, 0);
  VDim(ABackPageNumbersRight, SheetsCount, 0);
  for i:= 0 to SheetsCount-1 do
  begin
    n:= 2*i;
    AFacePageNumbersRight[i]:= APageNumbers[n];
    ABackPageNumbersLeft[i]:=  APageNumbers[n+1];
  end;
  for i:= SheetsCount-1 downto 0 do
  begin
    n:=4*SheetsCount - 2*(i+1);
    ABackPageNumbersRight[i]:= APageNumbers[n];
    AFacePageNumbersLeft[i]:=  APageNumbers[n+1];
  end;
end;

end.

