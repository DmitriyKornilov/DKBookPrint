unit UMainForm;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Spin, ComCtrls, fpspreadsheetgrid, DividerBevel, VirtualTrees, DK_VSTTables,
  UUtils, UStrings, DK_Vector, DK_Matrix, DK_Dialogs, DK_StrUtils, Clipbrd,
  Translations, UAboutForm;

type

  { TMainForm }

  TMainForm = class(TForm)
    AddButton: TButton;
    AllSheetsCountLabel: TLabel;
    BackPagesCopyButton: TButton;
    BackPagesPanel: TPanel;
    BookDivideCheckBox: TCheckBox;
    BookPagesCountSpinEdit: TSpinEdit;
    AboutButton: TButton;
    CalcButton: TButton;
    ClearButton: TButton;
    DelButton: TButton;
    DividerBevel2: TDividerBevel;
    DividerBevel3: TDividerBevel;
    DividerBevel4: TDividerBevel;
    FacePagesCopyButton: TButton;
    FacePagesPanel: TPanel;
    Grid1: TsWorksheetGrid;
    Label2: TLabel;
    Label4: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    MainPanel: TPanel;
    Memo1: TMemo;
    Memo2: TMemo;
    PageNumbersPanel: TPanel;
    PagesLabel: TLabel;
    Panel1: TPanel;
    Panel10: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    RUButton: TRadioButton;
    ENButton: TRadioButton;
    DEButton: TRadioButton;
    FRButton: TRadioButton;
    SeparateResultsCheckBox: TCheckBox;
    SheetsCountLabel: TLabel;
    Splitter1: TSplitter;
    TabControl1: TTabControl;
    VT1: TVirtualStringTree;
    procedure AboutButtonClick(Sender: TObject);
    procedure AddButtonClick(Sender: TObject);
    procedure BackPagesCopyButtonClick(Sender: TObject);
    procedure BookDivideCheckBoxChange(Sender: TObject);
    procedure BookPagesCountSpinEditChange(Sender: TObject);
    procedure CalcButtonClick(Sender: TObject);
    procedure ClearButtonClick(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure FacePagesCopyButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure RUButtonClick(Sender: TObject);
    procedure ENButtonClick(Sender: TObject);
    procedure SeparateResultsCheckBoxChange(Sender: TObject);
    procedure TabControl1Change(Sender: TObject);

  private
    VSTParts: TVSTEdit;
    PagesSheet: TPagesSheet;

    PartNames: TStrVector;
    FirstPages, LastPages: TIntVector;

    FacePageNumbersLeft, FacePageNumbersRight,
    BackPageNumbersLeft, BackPageNumbersRight: TIntMatrix;

    procedure SetStrings;

    procedure AddPart;
    procedure DelPart;

    procedure UpdatePartNames;

    procedure SetInputVectors;
    procedure GetInputVectors;
    procedure ClearInputVectors;

    procedure SetDelButtonEnabled;

    procedure Clear;
    procedure ClearTabs;
    function VerifyParts(const AShowInfo: Boolean): Boolean;
    function Calculate(const AShowInfo: Boolean): Boolean;
    procedure DrawBook(AIndex: Integer);
    procedure DrawBook;
    procedure CalcAndDrawResults(const AShowInfo: Boolean);

    procedure Translate;

    procedure SelectCell;
  public

  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

{ TMainForm }

procedure TMainForm.FormCreate(Sender: TObject);
begin
  Grid1.Font.Assign(AddButton.Font);
  PagesSheet:= TPagesSheet.Create(Grid1);

  VSTParts:= TVSTEdit.Create(VT1);
  VSTParts.UnselectOnExit:= False;
  VSTParts.SelectedBGColor:= COLOR_LIGHTBLUE;
  VSTParts.HeaderBGColor:= clBtnFace;
  VSTParts.AutosizeColumnDisable;

  VSTParts.AddColumnRowTitles('', 160);
  VSTParts.ColumnRowTitlesBGColor:= clBtnFace;
  VSTParts.SetColumnRowTitles(PartNames, taLeftJustify);

  VSTParts.AddColumnInteger('', 160);
  VSTParts.AddColumnInteger('', 160);
  VSTParts.FixedColumnsCount:= 1;
  VSTParts.OnSelect:= @SelectCell;

  Translate;
end;

procedure TMainForm.AddButtonClick(Sender: TObject);
begin
  AddPart;
end;

procedure TMainForm.AboutButtonClick(Sender: TObject);
var
  AboutForm: TAboutForm;
begin
  AboutForm:= TAboutForm.Create(MainForm);
  AboutForm.ShowModal;
  FreeAndNil(AboutForm);
end;

procedure TMainForm.FacePagesCopyButtonClick(Sender: TObject);
begin
  Clipboard.AsText:= Memo1.Text;
end;

procedure TMainForm.BackPagesCopyButtonClick(Sender: TObject);
begin
  Clipboard.AsText:= Memo2.Text;
end;

procedure TMainForm.BookDivideCheckBoxChange(Sender: TObject);
begin
  BookPagesCountSpinEdit.Enabled:= BookDivideCheckBox.Checked;
  Label4.Visible:= BookDivideCheckBox.Checked;
  SheetsCountLabel.Visible:= BookDivideCheckBox.Checked;
  SeparateResultsCheckBox.Visible:= BookDivideCheckBox.Checked;

  CalcAndDrawResults(False);
end;

procedure TMainForm.BookPagesCountSpinEditChange(Sender: TObject);
var
  x: Integer;
begin
  x:= BookPagesCountSpinEdit.Value;
  if x<4 then
    BookPagesCountSpinEdit.Value:= 4
  else if x>64 then
    BookPagesCountSpinEdit.Value:= 64
  else
    BookPagesCountSpinEdit.Value:= LargerMultiple(x, 4);
  SheetsCountLabel.Caption:= IntToStr(BookPagesCountSpinEdit.Value div 4);

  CalcAndDrawResults(False);
end;

procedure TMainForm.CalcButtonClick(Sender: TObject);
begin
  CalcAndDrawResults(True);
end;

procedure TMainForm.ClearButtonClick(Sender: TObject);
begin
  Clear;
end;

procedure TMainForm.DelButtonClick(Sender: TObject);
begin
  DelPart;
  SetDelButtonEnabled;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  if Assigned(VSTParts) then FreeAndNil(VSTParts);
  if Assigned(PagesSheet) then FreeAndNil(PagesSheet);
end;

procedure TMainForm.FormShow(Sender: TObject);
begin
  Clear;
end;

procedure TMainForm.RUButtonClick(Sender: TObject);
begin
  Translate;
end;

procedure TMainForm.ENButtonClick(Sender: TObject);
begin
  Translate;
end;

procedure TMainForm.SeparateResultsCheckBoxChange(Sender: TObject);
begin
  CalcAndDrawResults(False);
end;

procedure TMainForm.TabControl1Change(Sender: TObject);
begin
  DrawBook(TabControl1.TabIndex);
end;

procedure TMainForm.SelectCell;
begin
  SetDelButtonEnabled;
end;

procedure TMainForm.SetStrings;
var
  i: Integer;
begin
  VSTParts.RenameColumn(0, STR_PARTS);
  VSTParts.RenameColumn(1, STR_FIRST_PAGE);
  VSTParts.RenameColumn(2, STR_LAST_PAGE);
  UpdatePartNames;
  VSTParts.SetColumnRowTitles(PartNames, taLeftJustify);

  BookDivideCheckBox.Caption:= STR_DIVIDE;
  PagesLabel.Caption:= STR_PAGES;
  SeparateResultsCheckBox.Caption:= STR_RESULTS;
  AddButton.Caption:= STR_PART_ADD;
  DelButton.Caption:= STR_PART_DEL;
  CalcButton.Caption:= STR_CALCULATE;
  ClearButton.Caption:= STR_CLEAR;

  DividerBevel2.Caption:= STR_REQUIRED;
  Label2.Caption:= '- ' + SLower(STR_TOTAL);
  Label4.Caption:= '- ' + SLower(STR_ONEBOOK);
  for i:= 0 to TabControl1.Tabs.Count-1 do
    TabControl1.Tabs.Strings[i]:= STR_BOOK + ' ' + IntToStr(i+1);
  DividerBevel3.Caption:= STR_PAGE_NUMBERS;
  Label7.Caption:= '- ' + SLower(STR_SIDE) + ' 1';
  FacePagesCopyButton.Caption:= STR_COPY;
  Label8.Caption:= '- ' + SLower(STR_SIDE) + ' 2';
  BackPagesCopyButton.Caption:= STR_COPY;
  DividerBevel4.Caption:= STR_ARRANGEMENT;
end;

procedure TMainForm.AddPart;
begin
  GetInputVectors;
  VAppend(PartNames, STR_PART + ' ' + IntToStr(Length(PartNames)+1));
  VAppend(FirstPages, 0);
  VAppend(LastPages, 0);
  SetInputVectors;
end;

procedure TMainForm.DelPart;
var
  Ind: Integer;
begin
  Ind:= VSTParts.SelectedRowIndex;
  GetInputVectors;
  VDel(PartNames, Ind);
  VDel(FirstPages, Ind);
  VDel(LastPages, Ind);
  UpdatePartNames;
  SetInputVectors;
end;

procedure TMainForm.SetInputVectors;
begin
  VSTParts.ValuesClear;
  VSTParts.SetColumnRowTitles(PartNames, taLeftJustify);
  VSTParts.SetColumnInteger(1, FirstPages);
  VSTParts.SetColumnInteger(2, LastPages);
  VSTParts.Draw;
end;

procedure TMainForm.GetInputVectors;
begin
  VSTParts.ColumnAsInteger(FirstPages, 1);
  VSTParts.ColumnAsInteger(LastPages, 2);
end;

procedure TMainForm.ClearInputVectors;
begin
  VDim(PartNames, 1, STR_PART + ' 1');
  VDim(FirstPages, 1, 0);
  VDim(LastPages, 1, 0);
end;

procedure TMainForm.UpdatePartNames;
var
  i: Integer;
begin
  for i:= 0 to High(PartNames) do
    PartNames[i]:= STR_PART + ' ' + IntToStr(i+1);
end;

procedure TMainForm.SetDelButtonEnabled;
begin
  DelButton.Enabled:= VSTParts.IsSelected and (Length(PartNames)>1);
end;

function TMainForm.VerifyParts(const AShowInfo: Boolean): Boolean;
var
  i, j: Integer;
  Min1, Max1, Min2, Max2: Integer;
begin
  Result:= False;

  for i:= 0 to High(PartNames) do
  begin
    if (FirstPages[i]=0) or (LastPages[i]=0) then
    begin
      if AShowInfo then
        ShowInfo(PartNames[i] + ': ' + STR_RANGE_ERROR + '!');
      Exit;
    end;
  end;

  for i:= 0 to High(PartNames)-1 do
  begin
    Min1:= FirstPages[i];
    Max1:= LastPages[i];
    for j:= i+1 to High(PartNames) do
    begin
      Min2:= FirstPages[j];
      Max2:= LastPages[j];
      if IsRangesItersect(Min1, Max1, Min2, Max2) then
      begin
        if AShowInfo then
          ShowInfo(PartNames[i] + ', ' + PartNames[j] + ': ' + STR_RANGE_INTERSECT + '!');
        Exit;
      end;
    end;
  end;

  Result:= True;
end;

function TMainForm.Calculate(const AShowInfo: Boolean): Boolean;
var
  PagesPerBook: Byte;
  i: Integer;
  DocPageNumbers, TmpVector: TIntVector;
  BooksPageNumbers: TIntMatrix;
begin
  Result:= False;

  GetInputVectors;

  if not VerifyParts(AShowInfo) then Exit;

  //кол-во страниц в книге (0 - если не нужно разбивать на книги)
  PagesPerBook:= Ord(BookDivideCheckBox.Checked)*BookPagesCountSpinEdit.Value;
  //получаем вектор номеров страниц, которые нужно печатать
  DocPageNumbers:= nil;
  for i:= 0 to High(FirstPages) do
  begin
    TmpVector:= VRange(FirstPages[i], LastPages[i]);
    DocPageNumbers:= VAdd(DocPageNumbers, TmpVector);
  end;
  //кол-во страниц для печати
  i:= Length(DocPageNumbers);
  //кол-во страниц для расположения на листах А4
  if PagesPerBook=0 then
    i:= LargerMultiple(i, PAGES_PER_SHEET)
  else
    i:= LargerMultiple(i, PagesPerBook);
  //дозаполняем нулями вектор номеров страниц
  VReDim(DocPageNumbers, i, 0);
  //выводим требующееся кол-во листов А4
  AllSheetsCountLabel.Caption:= IntToStr(i div PAGES_PER_SHEET);
  //разбиваем вектор номеров страниц по книгам
  BooksPageNumbers:= PageNumbersDocToBooks(PagesPerBook, DocPageNumbers);
  //распределяем страницы по листам А4
  PlacePageNumbersOnSheets(BooksPageNumbers, FacePageNumbersLeft, FacePageNumbersRight,
                           BackPageNumbersLeft, BackPageNumbersRight);


  Result:= True;
end;

procedure TMainForm.Clear;

begin
  MainPanel.Visible:= False;

  Memo1.Text:= EmptyStr;
  Memo2.Text:= EmptyStr;
  Grid1.Clear;


  ClearTabs;
  ClearInputVectors;
  SetInputVectors;
end;

procedure TMainForm.ClearTabs;
begin
  TabControl1.Tabs.Clear;
  TabControl1.Tabs.Add(STR_BOOK + ' 1');
  TabControl1.TabIndex:= 0;
end;

procedure TMainForm.DrawBook(AIndex: Integer);
var
  V1, V2, V3, V4: TIntVector;
begin
  Grid1.Clear;

  if MIsNil(FacePageNumbersLeft) then Exit;
  if AIndex=-1 then
  begin
    V1:= MToVector(FacePageNumbersLeft);
    V2:= MToVector(FacePageNumbersRight);
    V3:= MToVector(BackPageNumbersLeft);
    V4:= MToVector(BackPageNumbersRight);
  end
  else begin
    V1:= FacePageNumbersLeft[AIndex];
    V2:= FacePageNumbersRight[AIndex];
    V3:= BackPageNumbersLeft[AIndex];
    V4:= BackPageNumbersRight[AIndex];
  end;

  Memo1.Text:= PageNumberVectorsToString(V1, V2);
  Memo2.Text:= PageNumberVectorsToString(V3, V4);


  PagesSheet.Draw(V1, V2, V3, V4, STR_SIDE+' 1', STR_SIDE+' 2', STR_SHEET_NUMBER);


end;

procedure TMainForm.DrawBook;
var
  i: Integer;
begin
  ClearTabs;

  if BookDivideCheckBox.Checked and SeparateResultsCheckBox.Checked then
    for i:= 2 to Length(FacePageNumbersLeft) do
      TabControl1.Tabs.Add(STR_BOOK + ' ' + IntToStr(i));

  if BookDivideCheckBox.Checked and (not SeparateResultsCheckBox.Checked) then
    DrawBook(-1)
  else
    DrawBook(0);
end;

procedure TMainForm.CalcAndDrawResults(const AShowInfo: Boolean);
begin
  if Calculate(AShowInfo) then
  begin
    DrawBook;
    MainPanel.Visible:= True;
  end
  else
    MainPanel.Visible:= False;
end;

procedure TMainForm.Translate;
var
  POFileName: String;
begin
  POFileName := ExtractFilePath(Application.ExeName) +'languages' + DirectorySeparator;

  if RUButton.Checked then
    POFileName:= POFileName+ 'DKBookPrint.ru.po'
  else if ENButton.Checked then
    POFileName:= POFileName+ 'DKBookPrint.en.po'
  else if DEButton.Checked then
    POFileName:= POFileName+ 'DKBookPrint.de.po'
  else if FRButton.Checked then
    POFileName:= POFileName+ 'DKBookPrint.fr.po';
  TranslateUnitResourceStrings('UStrings', POFileName);

  SetStrings;
  CalcAndDrawResults(False);
end;



end.

