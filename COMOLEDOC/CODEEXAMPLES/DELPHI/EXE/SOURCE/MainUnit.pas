unit MainUnit;
{
Please notice:
    This is a demonstration of adding, modifying, deleting object of type
    (IPCsArc  +  IPCsText  +  IPCsLines  +  IPCsSymbol) in
              PC|SCHEMATIC Automation.
    When giving object values, a lot of errorchecking code have not been made.
    This means that out of range and illegal values might occur.
}

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DPSCAD_TLB, StdCtrls, ComCtrls, ExtCtrls, ImgList, Inifiles, Grids,
  TypInfo;
  
type
  {Used for color-comboboxes}
  TPenColorIndex =
   (Co_Black,Co_Red,Co_Green,Co_Yellow,Co_Blue,Co_Magenta,Co_Cyan,Co_White,
    Co_LRed,Co_LGreen,Co_Brown,Co_LBlue,Co_LMagenta,Co_LCyan,Co_NoPrint,Co_Gray,
    Co_Dim, Co_HelpColor,Co_WinColor,co_BlackColor,co_HighLight);

  {Used for color-comboboxes}
  TPenColors = array[Co_Black..co_HighLight] of TColor;

  TTermDataType = (td_ExtComp, td_ExtWireNo, td_ExtCable, td_ExtSignal, td_OutConn,
                   td_InConn, td_InSignal, td_InCable, td_InWireNo, td_InComp, td_JumperLink);

  TPositionValues = record
    PointName   : string;
    XValue      : integer;
    YValue      : integer;
    ZValue      : integer;
    PageVal     : integer;
    LayerVal    : integer;
  end;
  TPointArray  = Array of TPositionValues;

  TSymbolArray = record
    SymHandle        : cardinal;
    SymName          : string;
    InPos            : array of TPositionValues;
    OutPos           : array of TPositionValues;
  end;

  TConnSymbol = record
    SymHandle        : cardinal;
    SymName          : string;
    SymPage          : integer;
    SymLayer         : integer;
    ConnPoints       : array of TPositionValues;
    State            : integer;
  end;

  TConnLines = record
    LineHandle        : cardinal;
    LinePage          : integer;
    LineLayer         : integer;
    LinePoints        : array of TPositionValues;
  end;

  TTermPosition = (TermHandle, OutputWireNo, OutputCabel, OutputName, OutputBridge,
                   TermName, InputBridge, InputName, InputCabel, InputWireNo);

type
  TMainForm = class(TForm)
    PageControl: TPageControl;
    TabSheetCreateConnection: TTabSheet;
    ButtonConnectToAutomation: TButton;
    MemoInfo: TMemo;
    TabSheetGetNumberOfProjects: TTabSheet;
    ButtonGetNumberOfOpenedProjects: TButton;
    TabSheetGetNumberOfPagesInAProject: TTabSheet;
    ButtonGetNumberOfPagesInProject: TButton;
    TabSheetMoveSymbolOnAPage: TTabSheet;
    ButtonFindSymbolsOnPage: TButton;
    ListBoxSymbols: TListBox;
    ButtonMoveSymbol: TButton;
    ComboboxDirection: TCombobox;
    EditMovingUnits: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    TabSheetDimensionofSymbol: TTabSheet;
    ButtonFindDimension: TButton;
    ListBoxDimensions: TListBox;
    ButtonChangeSymbolName: TButton;
    EditSymbolName: TEdit;
    Label3: TLabel;
    ButtonDeleteSymbol: TButton;
    TabSheetNewSymbol: TTabSheet;
    TabSheetNewPage: TTabSheet;
    ButtonSetAsActiveProject: TButton;
    ListBoxNumberOfOpenedProjects: TListBox;
    ButtonSetPageasActive: TButton;
    ListBoxNumberOfPages: TListBox;
    ButtonCreateNewProject: TButton;
    EditProjectName: TEdit;
    Label4: TLabel;
    ButtonCreateNewPage: TButton;
    EditPageName: TEdit;
    LabelPageTitle: TLabel;
    EditPathToProject: TEdit;
    LabelProjectPath: TLabel;
    EditProjectFilename: TEdit;
    Filename: TLabel;
    ComboboxPageType: TCombobox;
    Label5: TLabel;
    Label6: TLabel;
    ComboboxPageFunction: TCombobox;
    EditTabName: TEdit;
    Label7: TLabel;
    ComboboxPageTabName: TCombobox;
    LabelSpecifyName: TLabel;
    ButtonCreateSymbol: TButton;
    EditNewSymbolX: TEdit;
    EditNewSymbolY: TEdit;
    EditNewSymbolZ: TEdit;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    LabelSizeOfSymbol: TLabel;
    EditSymbolWidth: TEdit;
    EditSymbolHeight: TEdit;
    LabelSizeWidth: TLabel;
    LabelSizeHeight: TLabel;
    ComboboxSymbolFilename: TCombobox;
    EditSymbolNameText: TEdit;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    EditSymbolArticleText: TEdit;
    EditSymbolTypeText: TEdit;
    EditSymbolFunctionText: TEdit;
    Label16: TLabel;
    TabSheetCreateNewArcs: TTabSheet;
    TabSheetNewProject: TTabSheet;
    Label17: TLabel;
    ButtonCreateNewArcs: TButton;
    EditNewArcsX: TEdit;
    EditNewArcsY: TEdit;
    EditNewArcsZ: TEdit;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    EditRadius: TEdit;
    Label21: TLabel;
    EditStartAngle: TEdit;
    EditEndAngle: TEdit;
    Label22: TLabel;
    Label23: TLabel;
    CheckBoxFilled: TCheckBox;
    ComboboxColor: TCombobox;
    Label24: TLabel;
    ComboboxPenwidth: TCombobox;
    Label25: TLabel;
    ComboboxElipseFactor: TCombobox;
    Label26: TLabel;
    EditRotation: TEdit;
    Label27: TLabel;
    TabSheetModifyCircle: TTabSheet;
    ButtonfindCircles: TButton;
    ListBoxCircles: TListBox;
    EditArcX: TEdit;
    Label28: TLabel;
    Label29: TLabel;
    Label30: TLabel;
    EditArcY: TEdit;
    Label31: TLabel;
    EditArcZ: TEdit;
    EditArcRadius: TEdit;
    EditArcElipse: TEdit;
    Label32: TLabel;
    Label33: TLabel;
    EditArcRotation: TEdit;
    Label34: TLabel;
    EditArcStartAngle: TEdit;
    Label35: TLabel;
    EditArcEndAngle: TEdit;
    Label36: TLabel;
    CheckBoxArcFilled: TCheckBox;
    ComboboxArcColor: TCombobox;
    Label37: TLabel;
    Label38: TLabel;
    ButtonArcUpdate: TButton;
    ButtonArcDelete: TButton;
    EditArcPenWidth: TEdit;
    TabSheetCreateNewLine: TTabSheet;
    ButtonCreateNewLine: TButton;
    EditLineX: TEdit;
    EditLineY: TEdit;
    EditLineZ: TEdit;
    Label39: TLabel;
    Label40: TLabel;
    Label41: TLabel;
    ListBoxLines: TListBox;
    Label42: TLabel;
    ButtonDeleteThisPoint: TButton;
    ButtonAddThisPoint: TButton;
    ButtonLinesClear: TButton;
    ComboboxLineWidth: TCombobox;
    Label43: TLabel;
    Image1: TImage;
    TrackBarLine: TTrackBar;
    Label44: TLabel;
    LabelNewLineStyle: TLabel;
    Label46: TLabel;
    ComboboxLineDistance: TCombobox;
    ComboboxLineColor: TCombobox;
    Label47: TLabel;
    ButtonAddRandomPoint: TButton;
    CheckBoxCurvedLine: TCheckBox;
    TabSheetModifyLine: TTabSheet;
    ButtonFindLines: TButton;
    ListBoxModifyLines: TListBox;
    ButtonUpdateLine: TButton;
    ButtonDeleteLine: TButton;
    ListBoxLinePoints: TListBox;
    Label48: TLabel;
    EditLineModifyX: TEdit;
    EditLineModifyY: TEdit;
    EditLineModifyZ: TEdit;
    Label49: TLabel;
    Label50: TLabel;
    Label51: TLabel;
    Label52: TLabel;
    ComboboxLineModifyColor: TCombobox;
    Label53: TLabel;
    EditModifyLineWidth: TEdit;
    Label54: TLabel;
    EditModifyLineDistance: TEdit;
    LabelLineDistance: TLabel;
    CheckBoxModifyLineCurved: TCheckBox;
    TrackBarLineModify: TTrackBar;
    Image2: TImage;
    Label55: TLabel;
    LabelModifyLineStyle: TLabel;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    LbelPositionfortext: TLabel;
    EditTextPosX: TEdit;
    EditTextPosY: TEdit;
    EditTextPosZ: TEdit;
    Label56: TLabel;
    Label57: TLabel;
    Label58: TLabel;
    Label59: TLabel;
    ButtonCreateText: TButton;
    EditTextValue: TEdit;
    Label60: TLabel;
    ComboboxTextAngle: TCombobox;
    Label61: TLabel;
    ComboboxTextColor: TCombobox;
    Label62: TLabel;
    ComboboxTextFont: TCombobox;
    Label63: TLabel;
    CheckBoxTextBold: TCheckBox;
    CheckBoxTextItalic: TCheckBox;
    ComboboxTextHeight: TCombobox;
    ComboboxTextWidth: TCombobox;
    Label64: TLabel;
    Label65: TLabel;
    ButtonFindTexts: TButton;
    ListBoxFoundTexts: TListBox;
    EditFoundText: TEdit;
    Label66: TLabel;
    EditTextFoundX: TEdit;
    Label67: TLabel;
    EditTextFoundY: TEdit;
    Label68: TLabel;
    EditTextFoundZ: TEdit;
    Label69: TLabel;
    EditFoundTextAngle: TEdit;
    Label70: TLabel;
    ComboboxFoundTextColor: TCombobox;
    Color: TLabel;
    ComboboxFoundTextFont: TCombobox;
    Label71: TLabel;
    CheckBoxFoundTextBold: TCheckBox;
    CheckBoxFoundTextItalic: TCheckBox;
    EditFoundTextFontHeight: TEdit;
    EditFoundTextFontWidth: TEdit;
    ButtonUpdateText: TButton;
    Label72: TLabel;
    Label73: TLabel;
    ButtonDeleteText: TButton;
    ListBoxSymbolConnections: TListBox;
    Label74: TLabel;
    TabSheet3: TTabSheet;
    StringGridTerminalList: TStringGrid;
    ButtonGetTerminals: TButton;
    ButtonModifyLineAddPoint: TButton;
    EditSymbolFilename: TEdit;
    LabelSymbolFilename: TLabel;
    EditInDirectory: TEdit;
    LabelInDir: TLabel;
    TabSheet4: TTabSheet;
    ButtonGetConnections: TButton;
    StringGridConnections: TStringGrid;
    LabelSymbolType: TLabel;
    GroupBox1: TGroupBox;
    Label75: TLabel;
    Label76: TLabel;
    Label77: TLabel;
    Label78: TLabel;
    LabelSymbolsFoundAll: TLabel;
    LabelBranchSymbols: TLabel;
    LabelNotBranchSymbols: TLabel;
    LabelLinesFound: TLabel;
    Label79: TLabel;
    LabelJoinsFound: TLabel;
    Label80: TLabel;
    MemoConnectionInformation: TMemo;
    Image3: TImage;
    MemoBranchInformation: TMemo;
    EditSymbolType: TEdit;
    TabSheet5: TTabSheet;
    ButtonFindSymbolsALL: TButton;
    ListBoxSymbolsHighlight: TListBox;
    ComboBoxRectType: TComboBox;
    Label45: TLabel;
    TabSheetPageHeader: TTabSheet;
    ComboBoxPageHeader1: TComboBox;
    ComboBoxPageHeader2: TComboBox;
    LabelPageHeader1: TLabel;
    LabelPageHeader2: TLabel;
    ButtonSetOnActivePage: TButton;
    TabSheet6: TTabSheet;
    ButtonGetSymbols: TButton;
    ListBoxSymbolsDatafields: TListBox;
    ListBoxDatafields: TListBox;
    EditDataFieldsX: TEdit;
    ButtonAddSymbolWithDatafield: TButton;
    EditDataFieldsY: TEdit;
    EditSymbolNameDatafiled: TEdit;
    Label81: TLabel;
    Label82: TLabel;
    Label83: TLabel;
    EditSymbolFilenameDatafield: TEdit;
    Label84: TLabel;
    GroupBox2: TGroupBox;
    Label85: TLabel;
    ComboBoxOtherProject: TComboBox;
    ComboBoxOtherPage: TComboBox;
    RadioButtonOtherPro: TRadioButton;
    RadioButtonOtherPage: TRadioButton;
    Label88: TLabel;
    Label89: TLabel;
    EditPretext: TEdit;
    Label90: TLabel;
    RadioButtonDFSymbols: TRadioButton;
    Label91: TLabel;
    ComboBoxDFSymbols: TComboBox;
    EditDFFieldName: TEdit;
    ComboBoxDFCreateNew: TComboBox;
    Label92: TLabel;
    ButtonCreateDatafield: TButton;
    ShapeHorizon: TShape;
    Label86: TLabel;
    ShapeVertical: TShape;
    TabSheetLayers: TTabSheet;
    ButtonFindLayers: TButton;
    ListBoxLayers: TListBox;
    ButtonSetLayerVisible: TButton;
    ButtonSetLayerInvisible: TButton;
    Label87: TLabel;
    Label93: TLabel;
    Label94: TLabel;
    ButtonFindSymbols: TButton;
    Label95: TLabel;
    EditCompGrpNr: TEdit;
    ButtonSetCompGrpAndPos: TButton;
    Label96: TLabel;
    EditSymbolPosNo: TEdit;
    TabSheetLoadSave: TTabSheet;
    ButtonLoadSaveProBrowse: TButton;
    EditLoadSaveFile: TEdit;
    OpenDialog: TOpenDialog;
    ButtonLoadProject: TButton;
    ButtonSaveProject: TButton;
    procedure ButtonLoadProjectClick(Sender: TObject);
    procedure ButtonSaveProjectClick(Sender: TObject);
    procedure ButtonAddRandomPointClick(Sender: TObject);
    procedure ButtonAddThisPointClick(Sender: TObject);
    procedure ButtonArcDeleteClick(Sender: TObject);
    procedure ButtonArcUpdateClick(Sender: TObject);
    procedure ButtonChangeSymbolNameClick(Sender: TObject);
    procedure ButtonConnectToAutomationClick(Sender: TObject);
    procedure ButtonCreateNewArcsClick(Sender: TObject);
    procedure ButtonCreateNewLineClick(Sender: TObject);
    procedure ButtonCreateNewPageClick(Sender: TObject);
    procedure ButtonCreateNewProjectClick(Sender: TObject);
    procedure ButtonCreateSymbolClick(Sender: TObject);
    procedure ButtonCreateTextClick(Sender: TObject);
    procedure ButtonDeleteLineClick(Sender: TObject);
    procedure ButtonDeleteSymbolClick(Sender: TObject);
    procedure ButtonDeleteTextClick(Sender: TObject);
    procedure ButtonDeleteThisPointClick(Sender: TObject);
    procedure ButtonFindCirclesClick(Sender: TObject);
    procedure ButtonFindDimensionClick(Sender: TObject);
    procedure ButtonFindLinesClick(Sender: TObject);
    procedure ButtonFindSymbolsOnPageClick(Sender: TObject);
    procedure ButtonFindTextsClick(Sender: TObject);
    procedure ButtonGetNumberOfOpenedProjectsClick(Sender: TObject);
    procedure ButtonGetNumberOfPagesInProjectClick(Sender: TObject);
    procedure ButtonLinesClearClick(Sender: TObject);
    procedure ButtonMoveSymbolClick(Sender: TObject);
    procedure ButtonSetAsActiveProjectClick(Sender: TObject);
    procedure ButtonSetPageasActiveClick(Sender: TObject);
    procedure ButtonUpdateLineClick(Sender: TObject);
    procedure ButtonUpdateTextClick(Sender: TObject);
    procedure ComboboxColorDrawItem(Control: TWinControl; Index: Integer;
      Rect: TRect; State: TOwnerDrawState);
    procedure ComboboxPageTabNameChange(Sender: TObject);
    procedure ComboboxSymbolFilenameChange(Sender: TObject);
    procedure EditEndAngleExit(Sender: TObject);
    procedure EditRotationExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ListBoxCirclesClick(Sender: TObject);
    procedure ListBoxFoundTextsClick(Sender: TObject);
    procedure ListBoxLinePointsClick(Sender: TObject);
    procedure ListBoxModifyLinesClick(Sender: TObject);
    procedure ListBoxNumberOfOpenedProjectsClick(Sender: TObject);
    procedure ListBoxNumberOfPagesClick(Sender: TObject);
    procedure ListBoxSymbolsClick(Sender: TObject);
    procedure TrackBarLineChange(Sender: TObject);
    procedure TrackBarLineModifyChange(Sender: TObject);
    procedure ButtonGetTerminalsClick(Sender: TObject);
    procedure ButtonModifyLineAddPointClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure ButtonGetConnectionsClick(Sender: TObject);
    procedure ButtonFindSymbolsALLClick(Sender: TObject);
    procedure ListBoxSymbolsHighlightClick(Sender: TObject);
    procedure ComboBoxPageHeader1DropDown(Sender: TObject);
    procedure ComboBoxPageHeader1Change(Sender: TObject);
    procedure ButtonSetOnActivePageClick(Sender: TObject);
    procedure ButtonGetSymbolsClick(Sender: TObject);
    procedure ListBoxSymbolsDatafieldsClick(Sender: TObject);
    procedure ButtonAddSymbolWithDatafieldClick(Sender: TObject);
    procedure TabSheet6Show(Sender: TObject);
    procedure ButtonCreateDatafieldClick(Sender: TObject);
    procedure ComboBoxPredefinedProjectChange(Sender: TObject);
    procedure TabSheetNewPageContextPopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
    procedure ButtonFindLayersClick(Sender: TObject);
    procedure ButtonSetLayerVisibleClick(Sender: TObject);
    procedure ButtonSetLayerInvisibleClick(Sender: TObject);
    procedure ButtonFindSymbolsClick(Sender: TObject);
    procedure ButtonLoadSaveProBrowseClick(Sender: TObject);
    procedure ButtonSetCompGrpAndPosClick(Sender: TObject);
    procedure EditSymbolPosNoChange(Sender: TObject);
    procedure PageControlDrawTab(Control: TCustomTabControl; TabIndex: Integer;
      const Rect: TRect; Active: Boolean);
  private
    { Private declarations }
    procedure CheckAndFollowThisPoint(StrValue : string; ThisPoint : TPositionValues);
    Procedure ClearStringGrid();
    procedure CollectAllLines();
    procedure CollectAllSymbols();
    procedure ConvertingCellsToBridge();
    procedure GetTerminalData();
    procedure FindLineJoins();
//    procedure Test();
    procedure FindAndHightlightSymbols(AllPages : Boolean);
    procedure SetColorCombobox(ABox: TCombobox);
    procedure ShowStringInMemo(ShowString : string);
    procedure UpdateTerminalAdBridge(TerminalHandle : cardinal; IsInput : boolean;
              var BValue : integer);
    procedure CreateDocument();
    function RetriveSymbolType(ThisSymbol: IPCSSymbol) : string;
    function NumberOfTimesThisPoint(TempPoint : TPositionValues) : integer;
    function FixStringLength(FirStr: string; ThisLength : integer) : string;
    function FutureWayOfBranch(ThisPoint : TPositionValues) : TPointArray;
    function GetSymbolsNameAndConnectPointName(ThisPoint : TPositionValues) : string;
    function IsABranchSymbol(ThisPoint : TPositionValues) : boolean;
    function IsASymbolConnectionPoint(ThisPoint : TPositionValues) : boolean;
    function IsSamePoint(Point1,Point2 : TPositionValues) : boolean;
    function CheckConnection() : boolean;
    function CutOut(CutOutString, MainString, EndCut: string) : integer;
    function GetArcFromHandle(ParsedHandle: cardinal) : IPCsArc;
    function GetArcHandleFromListbox(ThisListBox: TListbox) : integer;
    function GetSymbolFromSymbolHandle(ParsedHandle: cardinal) : IPCsSymbol;
    function GetSymbolHandleFromListBox(ThisListBox: TListbox) : integer;
    function GetTextFromHandle(ParsedHandle: cardinal): IPCsText;
    function PenIndexToColor(AIndex: integer): TPenColorIndex;
    function TranslateTrackPosToLineStyle(TrackPos: integer; Reversed: boolean) : integer;
    function ConvertDatafieldBetweenNumberAndName(Way, TheKey : string) : string;
    function NoActiveDocument() : boolean;
    procedure SettingSelectedLayer(ToVisible: boolean);
  public
    { Public declarations }
    AvoidInfiniteLoop : array of TPositionValues;
    PCSComObj         : IPCsApplication;      {object to access Automation}
    PCSComObj2        : IPCsApplication2;
  end;

const
  ProgVer         = ' V 1.09';
  NoneSelectedStr = 'None selected';

  cl_White     = $FFFFFF;
  cl_Black     = $000000;
  cl_WhiteDim  = $BFBFBF;
  cl_BlackDim  = $6F6F6F;
  cl_WhiteNP   = $DFDFDF;
  cl_BlackNP   = $5F5F5F;

  {Used for color-comboboxes}
  PenColors : TPenColors =
         (cl_Black,$0000FF,$00FF00,$00FFFF,$FF0000,$FF00FF,$FFFF00,cl_White,
          $000080,$008000,$0080FF,$800000,$8000FF,$FF8000,cl_WhiteNP,$9F9F9F,
          cl_WhiteDim,$0000FF,{15531775}$FFFFFF,$000000,$0000FF);
var
  MainForm        : TMainForm;
  PCsApplication  : IPCsApplication;
  SymArray        : array of TSymbolArray; // used for terminal list
  SymConnArray    : array of TConnSymbol;  // Used for connectionlist.
  LineConnArray   : array of TConnLines;   // Used for connectionlist.
  SymBranchArray  : array of TConnSymbol;  // Used for connectionlist.
  ColTitleArray : array[0..9] of string =
     ('Handle','Ext. WireNo','Ext. Cabel','Ext. Name','Ext. Bridge','Terminal',
      'Int. Bridge','Int. Name','Int. Cabel','Int. WireNo');

implementation

uses
  AppCall;

{$R *.dfm}

{reverse a string}
function Reverse(S : String): String;
var
   i : Integer;
begin
  result := '';
  for i := Length(S) downto 1 do
    result := result + Copy(S,i,1) ;
end;

{finds last occurence of a substring in a string}
function LastPos(const SubStr: String; const S: String): Integer;
begin
  result := Pos(Reverse(SubStr), Reverse(S)) ;
  if (result <> 0) then
    result := ((Length(S) - Length(SubStr)) + 1) - result + 1;
end;

{used for color-comboboxes}

procedure TMainForm.PageControlDrawTab(Control: TCustomTabControl;
  TabIndex: Integer; const Rect: TRect; Active: Boolean);
var
  TmpRect: TRect;
begin
  if Active then
    begin
      PageControl.Canvas.brush.Color := clWebLightBlue;
      PageControl.Canvas.Font.Style := [fsBold];
    end;
  PageControl.Canvas.FillRect(Rect);
  TmpRect := Rect;
  OffsetRect(TmpRect, 0, 4);
  DrawText(PageControl.Canvas.Handle, PChar(PageControl.Pages[TabIndex].Caption), -1, TmpRect, DT_CENTER or DT_VCENTER);
end;

function TMainForm.PenIndexToColor(AIndex: integer): TPenColorIndex;
begin
  if (AIndex = ORD(co_NoPrint)) then
    inc(AIndex)
  else
    if (AIndex = ORD(co_Gray)) then
      dec(AIndex);
  result := TPenColorIndex(AIndex);
end;

{used for color-comboboxes}
procedure TMainForm.SetColorCombobox(ABox: TCombobox);
var
  i : integer;
begin
  ABox.Items.Clear;
  for i := ORD(co_Black) to ORD(PRED(Co_NoPrint)) do
    ABox.Items.AddObject(CHR(i+ORD('0')),POINTER(i));
  ABox.Items.AddObject(CHR(ORD(co_Gray)+ORD('0')),POINTER(ORD(co_Gray)));
  ABox.Items.AddObject(CHR(ORD(co_NoPrint)+ORD('0')),POINTER(ORD(co_NoPrint)));
end;

{returns number of times a given cordinate are found in all lines}
function TMainForm.NoActiveDocument: boolean;
begin
  result := PCSComObj.Documents.Count = 0;
end;

function TMainForm.NumberOfTimesThisPoint(TempPoint : TPositionValues) : integer;
var
  i, t : integer;
begin
  result := 0;
  for i := 0 to Length(LineConnArray) - 1 do
    for t := 0 to Length(LineConnArray[i].LinePoints) - 1 do
      if IsSamePoint(LineConnArray[i].LinePoints[t],TempPoint) then
        inc(result);
end;

{count number of joins by comparing lines}
procedure TMainForm.FindAndHightlightSymbols(AllPages: Boolean);
var
  i      : integer;
  ThisPages : IPCsPages;
  BuildStr  : string;

  procedure ProcessSymbolsOnThisPage(ThisPage : IPCsPage);
  var
    t : integer;
  begin
    for t := 0 to ThisPage.Symbols.Count -1 do
      begin
        BuildStr := 'PageTab=' + ThisPage.PageNumber;
        BuildStr := FixStringLength(BuildStr,25);
        BuildStr := BuildStr + '  PageTitle=' + ThisPage.Title;
        BuildStr := FixStringLength(BuildStr,60);
        BuildStr := BuildStr + 'Symbol=' + ThisPage.Symbols.Item[t].FullName;
        BuildStr := FixStringLength(BuildStr,80);
        BuildStr := BuildStr + 'Handle=' + IntToStr(ThisPage.Symbols.Item[t].Handle);
        ListBoxSymbolsHighlight.Items.Add(BuildStr);
      end;
  end;

begin
  {checks connectionstatus. if not connected, then it tries to connect}
  Screen.Cursor := crHourGlass;
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;

      ListBoxSymbolsHighlight.Items.BeginUpdate;
      ListBoxSymbolsHighlight.Clear;
      ThisPages := PCSComObj.ActiveDocument.Drawing.Pages;
      {running through all symbols in project}

      if AllPages then
        begin
          for i := 0 to ThisPages.Count - 1 do
            ProcessSymbolsOnThisPage(ThisPages.Item[i]); {running through all symbols on a page}
        end
      else
        ProcessSymbolsOnThisPage(PCSComObj.ActiveDocument.ActivePage);

      ListBoxSymbolsHighlight.Items.EndUpdate;
      ShowStringInMemo('Number of symbols found =  ' + IntToStr(ListBoxSymbolsHighlight.Count));
    end;
  Screen.Cursor := crDefault;
end;

procedure TMainForm.FindLineJoins;
var
  i, t, q, CountJoins     : integer;
  ArrayRememberedPoints   : Array of TPositionValues;
  FoundInArray            : boolean;
begin
  SetLength(ArrayRememberedPoints,0);
  // for each line
  for i := 0 to Length(LineConnArray) - 1 do
    // for each point in line
    for t := 0 to Length(LineConnArray[i].LinePoints) - 1 do
      if NumberOfTimesThisPoint(LineConnArray[i].LinePoints[t]) > 1 then
        begin
          // The point is a new point (as it's the first)
          if (Length(ArrayRememberedPoints) = 0)  then
            begin
              // first point, then just add it.
              SetLength(ArrayRememberedPoints,1);
              ArrayRememberedPoints[0] := LineConnArray[i].LinePoints[t];
            end
          else
            begin
              // find out if were allready found this join before
              FoundInArray := false;
              for q  := 0 to Length(ArrayRememberedPoints) - 1 do
                if IsSamePoint(ArrayRememberedPoints[q],LineConnArray[i].LinePoints[t]) then
                  FoundInArray := true;
              if not FoundInArray then
                begin
                  // this join is not found before, therefore add it to the list of
                  // remembered joins-points
                  SetLength(ArrayRememberedPoints,Length(ArrayRememberedPoints)+1);
                  ArrayRememberedPoints[Length(ArrayRememberedPoints)-1] := LineConnArray[i].LinePoints[t];
                end;
            end;
        end;

  // If a join is in a Symbolsconnectionpoint, then its not a join
  CountJoins := 0;
  if Length(ArrayRememberedPoints) > 0 then
    for i := 0 to Length(ArrayRememberedPoints) - 1 do
      if not IsASymbolConnectionPoint(ArrayRememberedPoints[i]) then
        inc(CountJoins);

  // show number of joins found
  LabelJoinsFound.Caption := IntToStr(CountJoins);
end;

{layout function. just to make a string a given length}
function TMainForm.FixStringLength(FirStr: string; ThisLength : integer) : string;
begin
  while (Length(FirStr) < ThisLength) do
    FirStr := FirStr + ' ';
  result := FirStr;
end;

{all the setup work, for comboboxes etc...}
procedure TMainForm.FormCreate(Sender: TObject);
var
  str,filestr    : string;
  MoreInInifile  : boolean;
  i              : integer;
  Inifile        : TInifile;
begin
  // fill comboboxes...
  ComboBoxPageHeader1.Text := NoneSelectedStr;
  ComboBoxPageHeader2.Text := NoneSelectedStr;
  // Set pageheaders

  // set the default seperator to be '.'

  FormatSettings.DecimalSeparator := '.';
  Self.Caption := Self.Caption + ProgVer;

  {setting up color-Comboboxes}
  SetColorCombobox(ComboboxColor);
  SetColorCombobox(ComboboxArcColor);
  SetColorCombobox(ComboboxLineColor);
  SetColorCombobox(ComboboxLineModifyColor);
  SetColorCombobox(ComboboxTextColor);
  SetColorCombobox(ComboboxFoundTextColor);
  ComboboxTextColor.ItemIndex := 0;
  ComboboxLineModifyColor.ItemIndex := 0;
  ComboboxColor.ItemIndex := 0;
  ComboboxLineColor.ItemIndex := 0;

  // StringGridConnections captions...
  StringGridConnections.Cells[0,0] := 'From';
  StringGridConnections.Cells[1,0] := 'To';
  LabelSymbolsFoundAll.Caption  := '0';
  LabelBranchSymbols.Caption    := '0';
  LabelNotBranchSymbols.Caption := '0';
  LabelLinesFound.Caption       := '0';
  LabelJoinsFound.Caption       := '0';

  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      try
        {setting up fonts, read from PCSCAD.INI}
        filestr := PCSComObj.Preferences.Folders[ft_Home] + 'PCSCAD.INI';
        IniFile := TIniFile.Create(filestr);
        i := 0;
        MoreInInifile := true;
        ComboboxTextFont.Clear;
        while MoreInInifile do
          begin
            str := IniFile.ReadString('TextFonts','FontNo' + IntToStr(i),'');
            if (str <> '') then
              begin
                ComboboxTextFont.Items.Add(str);
                inc(i);
              end;
            MoreInInifile := str <> '';
          end;
        if (ComboboxTextFont.Items.Count > 0) then
          ComboboxTextFont.ItemIndex := 0;
      except
        ShowMessage('Can''t read from ini file: ' + filestr + #13#10 + 'Can''t setup font-comboboxes');
      end;
    end;
  ComboboxFoundTextFont.Items := ComboboxTextFont.Items;
  PageControl.ActivePage := TabSheetCreateConnection;

  // Colnames...
  for i := 0 to Length(ColTitleArray)-1 do
    StringGridTerminalList.Cells[i,0] := ColTitleArray[i];
end;

{dynamicly adjust defaultcolwidt of StringGridTerminallist, according to the
  width of the mainform}
procedure TMainForm.FormResize(Sender: TObject);
begin
  StringGridTerminalList.DefaultColWidth := (MainForm.Width-62) div 10;
end;

{examines a branch-symbol, and find the new direction from that symbol}
function TMainForm.FutureWayOfBranch(ThisPoint: TPositionValues): TPointArray;
var
  i, t : integer;
begin
  // first, find the branchsymbol that has that parsed point.
  // then get the state of that branchsymbol, and get the new
  // direction to continue.
  // Please notice: It possible that more that 1 direction will be returned.
  // Therefore it an array, thats returned. (An array with Points)
  SetLength(Result,0);
  for i := 0 to Length(SymBranchArray) - 1 do
    for t := 0 to Length(SymBranchArray[i].ConnPoints) - 1 do
      if IsSamePoint(SymBranchArray[i].ConnPoints[t],ThisPoint) then
        begin
          // Brach found...
          case SymBranchArray[i].State of
            2,3,4,7,8 : begin
                          if t=0 then
                            begin
                              SetLength(Result,1);
                              Result[0] := SymBranchArray[i].ConnPoints[1];
                            end;
                          if t=1 then
                            begin
                              SetLength(Result,2);
                              Result[0] := SymBranchArray[i].ConnPoints[0];
                              Result[1] := SymBranchArray[i].ConnPoints[2];
                            end;
                          if t = 2 then
                            begin
                              SetLength(Result,1);
                              Result[0] := SymBranchArray[i].ConnPoints[1];
                            end;
                        end;
            5 : begin
                  if ((t=0) or (t=1)) then
                    begin
                      SetLength(Result,1);
                      Result[0] := SymBranchArray[i].ConnPoints[2];
                    end;
                  if t=2 then
                    begin
                      SetLength(Result,2);
                      Result[0] := SymBranchArray[i].ConnPoints[0];
                      Result[1] := SymBranchArray[i].ConnPoints[1];
                    end;
                end;
            6 : begin
                  if t=0 then
                    begin
                      SetLength(Result,2);
                      Result[0] := SymBranchArray[i].ConnPoints[1];
                      Result[1] := SymBranchArray[i].ConnPoints[2];
                    end;
                  if ((t=1) or (t=2)) then
                    begin
                      SetLength(Result,1);
                      Result[0] := SymBranchArray[i].ConnPoints[0];
                    end;
                end;
          else
            begin
              ShowMessage('Branch dont have expected state');
              halt;
            end;
          end;
        end;
end;

{accept a handle, and return the matching circle}
function TMainForm.GetArcFromHandle(ParsedHandle: cardinal): IPCsArc;
var
  i : integer;
begin
  {running through all circles on a page}
  for i := 0 to PCSComObj.ActiveDocument.ActivePage.Arcs.Count-1 do
    if (PCSComObj.ActiveDocument.ActivePage.Arcs[i].Handle = ParsedHandle) then
      begin
        {correct handle found, then return it}
        result := PCSComObj.ActiveDocument.ActivePage.Arcs[i];
        exit;
      end;
end;

{accept a handle, and return the matching Text}
procedure TMainForm.GetTerminalData;
var
  ListOfTerminals                             : IPCsList;
  i, IndexTerm, PosX, PosY, PosZ, p, oldLen   : integer;
  TerminalName                                : IPCsSymbol;
  EmptyString                                 : string;
  Skip                                        : boolean;
begin
  // init array
  SetLength(SymArray,0);
  // create a IPCsList of all terminals in correct active project
  ListOfTerminals := PCSComObj.ActiveDocument.Drawing.Lists.CreateTerminalList(EmptyString);
  // runningt through each terminal
  for i := 0 to ListOfTerminals.Count - 1 do
    begin
      // get symbol from IPCsList and put it in var TerminalName (IPCsSymbol)
      // by using an items GetDataSymbol-method
      ListOfTerminals.Item[i].GetDataSymbol(integer(td_InConn),0,TerminalName,IndexTerm); // terminal name

      // retrieve Symbolname and handle.
      if Terminalname <> nil then
        begin
          // make array ready
          SetLength(SymArray,i+1);
          SetLength(SymArray[i].InPos,0);
          SetLength(SymArray[i].OutPos,0);
          Skip := false;
          SymArray[i].SymName := TerminalName.FullName;
          SymArray[i].SymHandle := TerminalName.Handle;
        end
      else
        skip := true;

      // does terminal have connections?
      if ((not skip) and (TerminalName.Connections.Count > 0)) then
        begin
          // running through connectionpoints of that terminal.
          for p := 0 to TerminalName.Connections.Count - 1 do
            begin
              // get position of that connectionpoint
              TerminalName.Connections[p].GetPosition(PosX,PosY,PosZ);
              // is it an input or an output connection?
              if TerminalName.Connections[p].IOStatus.MainType = mt_Input then
                begin
                  // input
                  OldLen := Length(SymArray[i].InPos);
                  SetLength(SymArray[i].InPos,OldLen+1);
                  // retrieve position
                  with SymArray[i].InPos[OldLen] do
                    begin
                      XValue := PosX;
                      YValue := PosY;
                      ZValue := PosZ;
                    end;
                end
              else
                begin
                  // output
                  OldLen := Length(SymArray[i].OutPos);
                  SetLength(SymArray[i].OutPos,OldLen+1);
                  // retrieve position
                  with SymArray[i].OutPos[OldLen] do
                    begin
                      XValue := PosX;
                      YValue := PosY;
                      ZValue := PosZ;
                    end;
                end;
            end;
        end;
    end;
end;

{convert a handle to the text, that the handle points at}
function TMainForm.GetTextFromHandle(ParsedHandle: cardinal): IPCsText;
var
  i : integer;
begin
  {running through all texts on a page}
  for i := 0 to PCSComObj.ActiveDocument.ActivePage.Texts.Count-1 do
    if (PCSComObj.ActiveDocument.ActivePage.Texts[i].Handle = ParsedHandle) then
      begin
        {correct handle found, then return it}
        result := PCSComObj.ActiveDocument.ActivePage.Texts[i];
        exit;
      end;
end;

{gets handle from string in Listbox}
function TMainForm.GetArcHandleFromListbox(ThisListBox: TListbox): integer;
var
  str : string;
begin
  str := ThisListBox.Items[ThisListBox.ItemIndex];
  {correct handle found, then return it}
  result := StrToInt(copy(str,Pos('=',str)+1,pos('  ',str)-Pos('=',str)-1));
end;

{accept a handle, and return the matching symbols}
function TMainForm.GetSymbolFromSymbolHandle(ParsedHandle: cardinal): IPCsSymbol;
var
  i, t : integer;
  ADrawing : IPCsDrawing;
begin
  {running through all symbols on a page}
  result := nil;
  for i := 0 to PCSComObj.ActiveDocument.ActivePage.Symbols.Count-1 do
    if (PCSComObj.ActiveDocument.ActivePage.Symbols[i].Handle = ParsedHandle) then
      begin
        {correct handle found, then return it}
        result := PCSComObj.ActiveDocument.ActivePage.Symbols[i];
        exit;
      end;
  if result = nil then
    begin
      ADrawing := PCSComObj.ActiveDocument.Drawing;
      for t := 0 to ADrawing.Pages.Count - 1 do
        for i := 0 to ADrawing.Pages.Item[t].Symbols.Count-1 do
          if (ADrawing.Pages.Item[t].Symbols[i].Handle = ParsedHandle) then
            begin
              {correct handle found, then return it}
              result := ADrawing.Pages.Item[t].Symbols[i];
              exit;
            end;
    end;
end;

{get handle from the itemindex'ed string from the passed TListbox}
function TMainForm.GetSymbolHandleFromListBox(ThisListBox: TListbox): integer;
var
  str : string;
begin
  str := ThisListBox.Items.Strings[ThisListBox.ItemIndex];
  {correct handle found, then return it}
  result := StrToInt(trim(Copy(str,pos('Handle=',str)+7,Length(str))));
end;

{gets symbolname and connectionpointname, from a specific cordinate}
function TMainForm.GetSymbolsNameAndConnectPointName(
  ThisPoint: TPositionValues): string;
var
  i,t : integer;
begin
  result := '';
  for i := 0 to Length(SymConnArray) - 1 do
    if Length(SymConnArray[i].ConnPoints) > 0 then
      for t := 0 to Length(SymConnArray[i].ConnPoints) - 1 do
        if IsSamePoint(SymConnArray[i].ConnPoints[t],ThisPoint) then
          begin
            result := SymConnArray[i].SymName + ':' + SymConnArray[i].ConnPoints[t].PointName;
            exit;
          end;
end;

{get handle from listbox, and get values from the selected circle}
procedure TMainForm.ListBoxCirclesClick(Sender: TObject);
var
  AArcs            : IPCsArc;
  PosX, PosY, PosZ : integer;
begin
  ButtonArcUpdate.Enabled := ((ListBoxCircles.Items.Count > 0) and (ListBoxCircles.ItemIndex>-1));
  ButtonArcDelete.Enabled := ButtonArcUpdate.Enabled;
  if ButtonArcUpdate.Enabled then
    begin
      {assign circle, from handle}
      AArcs := GetArcFromHandle(GetArcHandleFromListbox(ListBoxCircles));
      {get position}
      AArcs.GetPosition(PosX,PosY,PosZ);
      EditArcX.Text := IntToStr(PosX);
      EditArcY.Text := IntToStr(PosY);
      EditArcZ.Text := IntToStr(PosZ);
      {get radius}
      EditArcRadius.Text := IntToStr(AArcs.Radius);
      {get elipse factor}
      EditArcElipse.Text := FloatToStr(AArcs.EllipseFactor);
      {get rotation}
      EditArcRotation.Text := IntToStr(AArcs.Rotation);
      {get starting angle}
      EditArcStartAngle.Text := IntToStr(AArcs.StartAngle);
      {get ending angle}
      EditArcEndAngle.Text := IntToStr(AArcs.EndAngle);
      {get filled on/off}
      CheckBoxArcFilled.Checked := AArcs.IsFilled;
      {get color}
      ComboboxArcColor.ItemIndex := AArcs.Color;
      {get pen width}
      EditArcPenWidth.Text := IntToStr(AArcs.PenWidth);
    end;
end;

{get handle from listbox, and get values from the selected text}
procedure TMainForm.ListBoxFoundTextsClick(Sender: TObject);
var
  ItemInListbox : cardinal;
  i : integer;
  AText            : IPCsText;
  PosX, PosY, PosZ : integer;
begin
  if (ListBoxFoundTexts.ItemIndex > -1) then
    begin
      ItemInListbox := CutOut('Handle=',ListBoxFoundTexts.Items[ListBoxFoundTexts.ItemIndex],' ');
      {assign text, from handle}
      AText := GetTextFromHandle(ItemInListbox);
      {get textvalue}
      EditFoundText.Text := AText.Value;
      {get rotation }
      EditFoundTextAngle.Text := IntToStr(AText.Rotation);
      {get position}
      AText.GetPosition(PosX,PosY,PosZ);
      ComboboxFoundTextColor.ItemIndex := AText.Color;
      EditTextFoundX.Text := IntToStr(PosX);
      EditTextFoundY.Text := IntToStr(PosY);
      EditTextFoundZ.Text := IntToStr(PosZ);
      {get fontname}
      i := 0;
      while (I < ComboboxFoundTextFont.Items.Count) do
        begin
          if (ComboboxFoundTextFont.Items[i] = AText.Font.Name) then
            ComboboxFoundTextFont.ItemIndex := i;
          inc(i);
        end;
      {get font options}
      CheckBoxFoundTextBold.Checked := AText.Font.Bold;
      CheckBoxFoundTextItalic.Checked := AText.Font.Italic;
      {get heigth and width}
      EditFoundTextFontHeight.Text := IntToStr(AText.Font.Height);
      EditFoundTextFontWidth.Text := IntToStr(AText.Font.Width);
    end;
  ButtonUpdateText.enabled := ListboxFoundTexts.ItemIndex > -1;
  ButtonDeleteText.enabled := ListboxFoundTexts.ItemIndex > -1;
end;

{get handle from linepoints, and get values from the selected point}
procedure TMainForm.ListBoxLinePointsClick(Sender: TObject);
var
  ALines           : IPCsLines;
  LineNr, PointNr  : integer;
  PosX, PosY, PosZ : integer;
begin
  if (ListBoxLinePoints.ItemIndex > -1) then
    begin
      {assign line, from activepage}
      ALines := PCSComObj.ActiveDocument.ActivePage.Lines;
      {get linenumber}
      LineNr := CutOut('#',ListBoxModifyLines.Items[ListBoxModifyLines.ItemIndex],' ');
      {get pointnumber}
      PointNr := CutOut('#',ListBoxLinePoints.Items[ListBoxLinePoints.ItemIndex],' ');
      {get poistion of point}
      ALines.Item[LineNr].Points[PointNr].GetPosition(PosX,PosY,PosZ);
      EditLineModifyX.Text := IntToStr(PosX);
      EditLineModifyY.Text := IntToStr(PosY);
      EditLineModifyZ.Text := IntToStr(PosZ);
      {get color}
      ComboboxLineModifyColor.ItemIndex := ALines.Item[LineNr].Color;
      {get line width}
      EditModifyLineWidth.Text := IntToStr(ALines.Item[LineNr].PenWidth);
      {get is curved}
      CheckBoxModifyLineCurved.Checked := ALines.Item[LineNr].IsCurvedLine;
      {get line distance}
      EditModifyLineDistance.Text := IntToStr(ALines.Item[LineNr].MultilineDistance);
      {get style of line}
      LabelModifyLineStyle.Caption := IntToStr(TranslateTrackPosToLineStyle(ALines.Item[LineNr].LineStyle,true));
      TrackBarLineModify.Position := StrToInt(LabelModifyLineStyle.Caption);
      ButtonUpdateLine.Enabled := true;
      ButtonDeleteLine.Enabled := true;
      ShowStringInMemo('Point ' + IntToStr(ListBoxLinePoints.ItemIndex) + ' selected');
    end
  else
    begin
      ButtonUpdateLine.Enabled := false;
      ButtonDeleteLine.Enabled := false;
    end;
end;

{get handle from listbox, and get all points in that line}
procedure TMainForm.ListBoxModifyLinesClick(Sender: TObject);
var
  ItemInListbox, i : integer;
  ALines           : IPCsLines;
  PosX, PosY, PosZ : integer;
begin
  if (ListBoxModifyLines.ItemIndex > -1) then
    begin
      {as a line is now selected, enable button for add point}
      ButtonModifyLineAddPoint.Enabled := true;
      {get index of that line that is selected}
      ItemInListbox := CutOut('#',ListBoxModifyLines.Items[ListBoxModifyLines.ItemIndex],' ');
      {assign line, from activepage}
      ALines := PCSComObj.ActiveDocument.ActivePage.Lines;
      ListBoxLinePoints.Items.BeginUpdate;
      ListBoxLinePoints.Clear;
      {running through all lines}
      for i := 0 to Alines.Item[ItemInListbox].Points.Count-1 do
        begin
          {get position of each point}
          ALines.Item[ItemInListbox].Points.Item[i].GetPosition(PosX,PosY,PosZ);
          ListBoxLinePoints.Items.Add('#' + IntToStr(i) + '  (X=' + IntToStr(PosX) + ',Y=' + IntToStr(PosY) + ',Z=' + IntToStr(PosZ) + ')');
        end;
      ListBoxLinePoints.Items.EndUpdate;
      ShowStringInMemo('Line ' + IntToStr(ItemInListbox) + ' selected');
    end;
  ButtonDeleteLine.Enabled := ListBoxModifyLines.ItemIndex > -1;
end;

{button enabled on/off}
procedure TMainForm.ListBoxNumberOfOpenedProjectsClick(Sender: TObject);
begin
  ButtonSetAsActiveProject.Enabled := ((ListBoxNumberOfOpenedProjects.Items.Count > 0) and (ListBoxNumberOfOpenedProjects.ItemIndex>-1));
end;

{button enabled on/off}
procedure TMainForm.ListBoxNumberOfPagesClick(Sender: TObject);
begin
  ButtonSetPageasActive.Enabled := ((ListBoxNumberOfPages.Items.Count > 1) and (ListBoxNumberOfPages.ItemIndex>0));
end;

{returns the symbols type as a string}
function TMainForm.RetriveSymbolType(ThisSymbol: IPCSSymbol) : string;
begin
  case ThisSymbol.SymbolType1 of
    st_Normal       : result := 'st_Normal';
    st_Head         : result := 'st_Head';
    st_Ignore       : result := 'st_Ignore';
    st_Relay        : result := 'st_Relay';
    st_Open         : result := 'st_Open';
    st_Close        : result := 'st_Close';
    st_Switch       : result := 'st_Switch';
    st_MasterRef    : result := 'st_MasterRef';
    st_WithRef      : result := 'st_WithRef';
    st_Reference    : result := 'st_Reference';
    st_Signal       : result := 'st_Signal';
    st_Terminal     : result := 'st_Terminal';
    st_Plc          : result := 'st_Plc';
    st_Data         : result := 'st_Data';
    st_NoneElectric : result := 'st_NoneElectric';
    st_Support      : result := 'st_Support';
    st_Power        : result := 'st_Power';
    st_Cable        : result := 'st_Cable';
    st_WireNumber   : result := 'st_WireNumber';
    st_TestPoint    : result := 'st_TestPoint';
    st_MultiSignal  : result := 'st_MultiSignal';
    st_DynCable     : result := 'st_DynCable';
    st_DynTerm      : result := 'st_DynTerm';
    st_PowerSource  : result := 'st_PowerSource';
    st_PowerDistrb  : result := 'st_PowerDistrb';
    st_PowerLoad    : result := 'st_PowerLoad';
    st_PowerCable   : result := 'st_PowerCable';
    st_CableDuct    : result := 'st_CableDuct';
    st_Insertion    : result := 'st_Insertion';
    st_PlcRef       : result := 'st_PlcRef';
    st_Measure      : result := 'st_Measure';
    st_LineConnect  : result := 'st_LineConnect';
    st_Branch       : result := 'st_Branch';
    st_BusBar       : result := 'st_BusBar';
    st_PowerFuse    : result := 'st_PowerFuse';
  end;
end;

{get handle from listbox and get values from the selected symbol}
procedure TMainForm.ListBoxSymbolsClick(Sender: TObject);
var
  ThisHandle                  : cardinal;
  Symbol                      : IPCsSymbol;
  i, PosX, PosY, PosZ, LPos   : integer;
  tempStr                     : string;
begin
  ButtonMoveSymbol.Enabled := ((ListBoxSymbols.Items.Count > 0) and (ListBoxSymbols.ItemIndex>-1));
  ButtonChangeSymbolName.Enabled := ButtonMoveSymbol.Enabled;
  ButtonDeleteSymbol.Enabled := ButtonMoveSymbol.Enabled;
  ButtonSetCompGrpAndPos.Enabled := ButtonMoveSymbol.Enabled;
  if ButtonChangeSymbolName.Enabled then
    begin
      ThisHandle := GetSymbolHandleFromListBox(ListBoxSymbols);
      {assign symbol, from handle}
      Symbol := GetSymbolFromSymbolHandle(ThisHandle);
      {get symbol name}
      EditSymbolName.Text := Symbol.FullName;
      EditSymbolName.Hint := EditSymbolName.Text;

      LPos := LastPos('\',Symbol.LibType.Name);
      EditSymbolFilename.Text := Copy(Symbol.LibType.Name,LPos+1,Length(Symbol.LibType.Name));
      EditSymbolFilename.Hint := EditSymbolFilename.Text;

      EditInDirectory.Text := Copy(Symbol.LibType.Name,0,LPos);
      Label74.Caption := 'Symbol connections ' + inttostr(Symbol.Connections.Count);
      ListBoxSymbolConnections.Items.BeginUpdate;
      ListBoxSymbolConnections.Clear;

      EditSymbolType.Text := RetriveSymbolType(Symbol);
      EditSymbolType.Hint := EditSymbolType.Text;

      EditSymbolPosNo.Text := IntToStr(Symbol.ComponentPosNo);
      EditCompGrpNr.Text := IntToStr(Symbol.ComponentGroupNo);
      {running through every connection that symbol has}
      for i := 0 to Symbol.Connections.Count - 1 do
        begin
          TempStr := 'Con.item ' + inttostr(i);
          TempStr := FixStringLength(TempStr,13);
          TempStr := TempStr + 'Handle=' + IntToStr(Symbol.Connections.Item[i].Handle);
          TempStr := FixStringLength(TempStr,30);

          {the connection name}
          TempStr := TempStr + 'Pinname=' + Symbol.Connections.Item[i].PinName;
          TempStr := FixStringLength(TempStr,48);
          TempStr := TempStr + 'IOStatus=';
          {is it an input or ourput, or none of these}
          case Symbol.Connections.Item[i].IOStatus.MainType of
            mt_None   : TempStr := TempStr + 'none';
            mt_Output : TempStr := TempStr + 'output';
            mt_Input  : TempStr := TempStr + 'input';
          end;

          TempStr := FixStringLength(TempStr,64);
          {The coordinate for that connection point}
          Symbol.Connections.Item[i].GetPosition(PosX,PosY,PosZ);
          TempStr := TempStr + '   Coords:X=' + inttostr(PosX) + ' Y=' + inttostr(PosY) + ' Z=' + inttostr(PosZ);
          TempStr := FixStringLength(TempStr,84);

          {add it to the listbox}
          ListBoxSymbolConnections.Items.Add(TempStr);
        end;
      ListBoxSymbolConnections.Items.EndUpdate;
    end;
end;

{Shows datafield that are on a specific symbol}
procedure TMainForm.ListBoxSymbolsDatafieldsClick(Sender: TObject);
var
  ItemInListbox : cardinal;
  ASymbol : IPCsSymbol;
  i : integer;
begin
  ListBoxDatafields.Items.BeginUpdate;
  ListBoxDatafields.Clear;
  {a valid item is clicked}
  if ListBoxSymbolsDatafields.ItemIndex > -1 then
    begin
      ItemInListbox := CutOut('Handle=',ListBoxSymbolsDatafields.Items[ListBoxSymbolsDatafields.ItemIndex],' ');
      {getting the symbol, specified in the listbox}
      ASymbol := GetSymbolFromSymbolHandle(ItemInListbox);
      if ASymbol = nil then
        begin
          ListBoxSymbolsDatafields.Clear;
          ButtonGetSymbols.Click;
          exit;
        end;
      for i := 0 to ASymbol.Datafields.Count - 1 do
        if ASymbol.Datafields.Item[i].FieldName <> '' then
          {none predefined value}
          ListBoxDatafields.Items.Add(ASymbol.Datafields.Item[i].FieldName + ' = ' + ASymbol.Datafields.Item[i].DisplayText)
        else
          {predefined value}
          ListBoxDatafields.Items.Add(ConvertDatafieldBetweenNumberAndName('NoToNa',IntToStr(ASymbol.Datafields.Item[i].DataNo)) + ' = ' + ASymbol.Datafields.Item[i].DisplayText);
     end;
  ListBoxDatafields.Items.EndUpdate;
end;

{highlights a specific symbol}
procedure TMainForm.ListBoxSymbolsHighlightClick(Sender: TObject);
var
  AreaLow, AreaHigh : TPCsPoint3D;
  ASymbol           : IPCsSymbol;
begin
  if (ListBoxSymbolsHighlight.ItemIndex > -1) then
    begin
      ASymbol := GetSymbolFromSymbolHandle(CutOut('Handle=',ListBoxSymbolsHighlight.Items[ListBoxSymbolsHighlight.ItemIndex],' '));
      AreaLow  := ASymbol.Position;
      AreaHigh := ASymbol.Position;
      dec(AreaLow.X);
      inc(AreaLow.Y);
      inc(AreaHigh.X);
      dec(AreaHigh.Y);
      PCSComObj.ActiveDocument.ActivePage := ASymbol.Page;
      PCSComObj.ActiveDocument.SelectInWindow(AreaLow, AreaHigh, ASymbol.Page,ASymbol.Layer,'');
      PCSComObj.Redraw;
      ShowStringInMemo('Highligthing symbol = ' + ASymbol.FullName);
    end;
end;

{examines a point, to see if its the same as a Symbol(branch)'s point}
function TMainForm.IsABranchSymbol(ThisPoint: TPositionValues): boolean;
var
  i, t : integer;
begin
  result := false;
  for i := 0 to Length(SymBranchArray) - 1 do
    for t := 0 to Length(SymBranchArray[i].ConnPoints) - 1 do
      if IsSamePoint(SymBranchArray[i].ConnPoints[t],ThisPoint) then
        begin
          result := true;
          exit;
        end;
end;

{examines a point, to see if its the same as a Symbols connection point}
function TMainForm.IsASymbolConnectionPoint(
  ThisPoint: TPositionValues): boolean;
var
  i, t : integer;
begin
  // a bad way of finding same coordinates, as this is a dumb and
  // very slow way of doing it. Arrays should be sorted, etc...
  result := false;
  for i := 0 to Length(SymConnArray) - 1 do
    if Length(SymConnArray[i].ConnPoints) > 0 then
      for t := 0 to Length(SymConnArray[i].ConnPoints) - 1 do
        if IsSamePoint(SymConnArray[i].ConnPoints[t],ThisPoint) then
            begin
              // exactly the same point found. Then set to true and quit...
              result := true;
              exit;
            end;
end;

{examines two points, to see if the are the same}
function TMainForm.IsSamePoint(Point1, Point2: TPositionValues): boolean;
begin
  // as it's not possible to say "if (point1 = point2) then",
  // this function is needed...
  result := false;
  if Point1.XValue = Point2.XValue then
    if Point1.YValue = Point2.YValue then
      if Point1.ZValue = Point2.ZValue then
        if Point1.PageVal = Point2.PageVal then
          if Point1.LayerVal = Point2.LayerVal then
            result := true;
end;

{if not connection to Automation, then do try to connect}
procedure TMainForm.CheckAndFollowThisPoint(StrValue: string;
  ThisPoint: TPositionValues);
var
  i, t, p, ThisLen, ThisRow : integer;
  FirstStr, SecondStr       : string;
  TempArray                 : TPointArray;
  UsingPoint                : TPositionValues;
  AvoidInfinity             : integer;
begin
  AvoidInfinity := 0;
  for i := 0 to Length(AvoidInfiniteLoop) - 1 do
    if IsSamePoint(Thispoint, AvoidInfiniteLoop[i]) then
      inc(AvoidInfinity);
  if (AvoidInfinity > 1) then
    exit;

  // for each line
  for i := 0 to Length(LineConnArray) - 1 do
    // for each point on line
    for t := 0 to Length(LineConnArray[i].LinePoints) - 1 do
      // is the same point as ThisPoint (parsed variable to procedure)
      if IsSamePoint(LineConnArray[i].LinePoints[t],ThisPoint) then
        begin
          // string to stringgrid
          FirstStr := StrValue;
          if t = 0 then
            // found at first point, then examine last point
            UsingPoint := LineConnArray[i].LinePoints[(Length(LineConnArray[i].LinePoints)-1)]
          else
            // found at last point, then examine first point
            UsingPoint := LineConnArray[i].LinePoints[0];
          // examine point...
          if IsASymbolConnectionPoint(UsingPoint) then
            begin
              // at a symbol connectionpoint, there add this to stringgrid
              SecondStr := GetSymbolsNameAndConnectPointName(UsingPoint);
              StringGridConnections.RowCount := StringGridConnections.RowCount+1;
              ThisRow := StringGridConnections.RowCount -2;
              StringGridConnections.Cells[0,ThisRow] := FirstStr;
              StringGridConnections.Cells[1,ThisRow] := SecondStr;
              // the line ends here...
              exit;
            end
          else
            if IsABranchSymbol(UsingPoint) then
              begin
                // if Point is a branch, then find the new way(s)
                TempArray := FutureWayOfBranch(UsingPoint);
                // for each new way, recursivly call CheckAndFollowThisPoint,
                // with that value
                for p := 0 to Length(TempArray) - 1 do
                  begin
                    ThisLen := Length(AvoidInfiniteLoop);
                    SetLength(AvoidInfiniteLoop,ThisLen+1);
                    AvoidInfiniteLoop[ThisLen] := TempArray[p];
                    CheckAndFollowThisPoint(StrValue, TempArray[p]);
                  end;
              end
            else
              // this will only be called if two different colors of lines connects..
              begin
                ThisLen := Length(AvoidInfiniteLoop);
                SetLength(AvoidInfiniteLoop,ThisLen+1);
                AvoidInfiniteLoop[ThisLen] := UsingPoint;
                CheckAndFollowThisPoint(StrValue, UsingPoint);
              end;
        end;
end;

{connects to Automation, if not allready connected}
function TMainForm.CheckConnection() : boolean;
begin
  if (PCSComObj = nil) then
    begin
      ShowStringInMemo('Trying to connect to Automation');
      try
        {tries to connection to Automation}
        PCSComObj := coApplication.Create();
        ShowStringInMemo('Connect to Automation succesfull');
        result := true;
      except
        {connection to Automation failed}
        On E:Exception do
          begin
            ShowMessage('Can''t connect to Automation' + #13#10 + E.ClassName+' error raised, with message : '+E.Message);
            ShowStringInMemo('Connecting to Automation failed');
            result := false;
          end;
      end;
    end
  else
    result := true;
end;

{clears stringgrid}
procedure TMainForm.ClearStringGrid;
var
  x,y : integer;
begin
  StringGridConnections.RowCount := 2;
  for x := 0 to 1 do
    for y := 1 to StringGridConnections.RowCount - 1 do
      StringGridConnections.Cells[x,y] := '';
end;

{labels on/off}
procedure TMainForm.ComboboxSymbolFilenameChange(Sender: TObject);
var
  bool : boolean;
begin
  bool := not (ComboboxSymbolFilename.ItemIndex > 0);
  LabelSizeOfSymbol.Visible := bool;
  LabelSizeWidth.Visible    := bool;
  LabelSizeHeight.Visible   := bool;
  EditSymbolWidth.Visible   := bool;
  EditSymbolHeight.Visible  := bool;
end;

{splits a string into an array (by seperator = Delimiter)}
procedure Split(const Delimiter: Char; Input: string; const Strings: TStrings);
begin
  // splits an string into an array or strings, delimited by "Delimiter"
  Assert(Assigned(Strings)) ;
  Strings.Clear;
  Strings.Delimiter     := Delimiter;
  Strings.DelimitedText := Input;
end;

{count number of occurences of one (sub)string in another string}
function CountPos(const subtext: string; Text: string): Integer;
begin
  if (Length(subtext) = 0) or (Length(Text) = 0) or (Pos(subtext, Text) = 0) then
    Result := 0
  else
    Result := (Length(Text) - Length(StringReplace(Text, subtext, '', [rfReplaceAll]))) div
      Length(subtext);
end;

procedure TMainForm.ButtonLoadProjectClick(Sender: TObject);
begin
  if FileExists(EditLoadSaveFile.Text) then
    try
      PCSComObj.Documents.Open(EditLoadSaveFile.Text)
    except
      ShowMessage('Error: An error occured while tring to load file:' + sLineBreak +
        EditLoadSaveFile.Text)
    end
  else
    ShowMessage('File don''t exist:' + sLineBreak + EditLoadSaveFile.Text);
end;

procedure TMainForm.ButtonSaveProjectClick(Sender: TObject);
var
  MessageAnswer : integer;
  s : string;
begin
  if (PCSComObj.Documents.Count < 1) then
    begin
      ShowMessage('Nothing to save (no document open)');
      exit;
    end;

  s := EditLoadSaveFile.Text;

  if (trim(s) = '') then
    begin
      MessageAnswer := MessageDlg('Saving using projects existing name', mtConfirmation, mbYesNo, 0);
      if (MessageAnswer = mrYes) then
        EditLoadSaveFile.Text := PCSComObj.ActiveDocument.FullName
      else
        exit;
    end;

  s := EditLoadSaveFile.Text;

  if FileExists(s) then
    begin
      MessageAnswer := MessageDlg('File allready exist. Do you want to overwrite it', mtConfirmation, mbYesNo, 0);
      if MessageAnswer = mrYes then
        begin
          try
            if not DeleteFile(s) then
              begin
                ShowMessage('Could not delete file:' + sLineBreak + s);
                exit;
              end;
          except
             on E : Exception do
             begin
               ShowMessage(E.ClassName + sLineBreak +E.Message);
               exit;
             end;
          end;
        end
      else
        exit;
    end
  else;

  PCSComObj.ActiveDocument.Save;
  ShowMessage('Project saved:' + sLineBreak + PCSComObj.ActiveDocument.FullName);
end;

{convert value in StringGridTerminalList (InputBridge+OutputBridge), to
only 1 number, to group them together more userfriendly}
function TMainForm.ConvertDatafieldBetweenNumberAndName(Way,
  TheKey: string): string;
var
  i : integer;
begin
  result := '';
  if Way = 'NoToNa' then
    begin
      i := StrToInt(TheKey);
      case i of
        // SystemData
        1 : result := 'FirstSystemData';
        9 : result := 'LastSystemData';
        // ProjData
        10 : result := 'FirstProjData';
        14 : result := 'ReferenceIDsField';
        15 : result := 'LogoField';
        17 : result := 'RemarksField';
        18 : result := 'ProjectNameIndexed';
        19 : result := 'LastProjectData';
        // PageData
        20 : result := 'FirstPageData';
        36 : result := 'PictureField';
        38 : result := 'PageNameIndexed';
        39 : result := 'LastPageData';
        // IndexData
        40 : result := 'FirstIndexData';
        51 : result := 'IndexNameIndexed';
        52 : result := 'IndexComments';
        59 : result := 'LastIndexData';
        // PartData
        60 : result := 'FirstPartData';
        66 : result := 'FigurField';
        84 : result := 'EAN13Field';
        99 : result := 'LastPartData';
        // TermData
        100 : result := 'FirstTermData';
        132 : result := 'JumperLinkField';
        149 : result := 'LastTermData';
        //CableData
        150 : result := 'FirstCableData';
        199 : result := 'LastCableData';
        // PLCData
        200 : result := 'FirstPLCData';
        249 : result := 'LastPLCData';
        // NetData
        250 : result := 'FirstNetData';
        261 : result := 'NetWireNoField';
        299 : result := 'LasttNetData';
        // SymbolData
        300 : result := 'FirstSymbolData';
        301 : result := 'SymbolDBDataField';
        309 : result := 'LastSymbolData';
        // PageDataExt
        310 : result := 'FirstPageDataExt';
        311 : result := 'PageCommentsData';
        329 : result := 'LastPageDataExt';
        // ConnsData
        330 : result := 'FirstConnsData';
        379 : result := 'LastConnsData';
      end;
    end
  else
    begin
        // SystemData
        if TheKey = 'FirstSystemData' then result := '1';
        if TheKey = 'LastSystemData' then result := '9';
        // ProjData
        if TheKey = 'FirstProjData' then result := '10';
        if TheKey = 'ReferenceIDsField' then result := '14';
        if TheKey = 'LogoField' then result := '15';
        if TheKey = 'RemarksField' then result := '17';
        if TheKey = 'ProjectNameIndexed' then result := '18';
        if TheKey = 'LastProjectData' then result := '19';
        // PageData
        if TheKey = 'FirstPageData' then result := '20';
        if TheKey = 'PictureField' then result := '36';
        if TheKey = 'PageNameIndexed' then result := '38';
        if TheKey = 'LastPageData' then result := '39';
        // IndexData
        if TheKey = 'FirstIndexData' then result := '40';
        if TheKey = 'IndexNameIndexed' then result := '51';
        if TheKey = 'IndexComments' then result := '52';
        if TheKey = 'LastIndexData' then result := '59';
        // PartData
        if TheKey = 'FirstPartData' then result := '60';
        if TheKey = 'FigurField' then result := '66';
        if TheKey = 'EAN13Field' then result := '84';
        if TheKey = 'LastPartData' then result := '99';
        // TermData
        if TheKey = 'FirstTermData' then result := '100';
        if TheKey = 'JumperLinkField' then result := '132';
        if TheKey = 'LastTermData' then result := '149';
        //CableData
        if TheKey = 'FirstCableData' then result := '150';
        if TheKey = 'LastCableData' then result := '199';
        // PLCData
        if TheKey = 'FirstPLCData' then result := '200';
        if TheKey = 'LastPLCData' then result := '249';
        // NetData
        if TheKey = 'FirstNetData' then result := '250';
        if TheKey = 'NetWireNoField' then result := '261';
        if TheKey = 'LasttNetData' then result := '299';
        // SymbolData
        if TheKey = 'FirstSymbolData' then result := '300';
        if TheKey = 'SymbolDBDataField' then result := '301';
        if TheKey = 'LastSymbolData' then result := '309';
        // PageDataExt
        if TheKey = 'FirstPageDataExt' then result := '310';
        if TheKey = 'PageCommentsData' then result := '311';
        if TheKey = 'LastPageDataExt' then result := '329';
        // ConnsData
        if TheKey = 'FirstConnsData' then result := '330';
        if TheKey = 'LastConnsData' then result := '379';
    end;
end;

{getting which terminals that a connected with bridges}
procedure TMainForm.ConvertingCellsToBridge;
type
  TBridge = record
    CellCordX : integer;
    CellCordY : integer;
    CellValue : string;
  end;
var
  BridgeArray                  : array of TBridge;
  i, q, p, r, ArLen, CellCol   : integer;
  TStrList                     : TStringList;
begin
  ArLen := 0;
  // array that will contain cellposition in stringgrid, plus cell value
  SetLength(BridgeArray,Arlen);

  // getting information into array.
  for r := 1 to 2 do
    begin
      case r of
        1 : CellCol := integer(InputBridge);
        2 : CellCol := integer(OutputBridge);
      else
        begin
          ShowMessage('Error in ConvertingCellsToBridge;');
          exit;
        end;
      end;
      for i := 1 to StringGridTerminalList.RowCount - 1 do
        if StringGridTerminalList.Cells[CellCol,i] <> '' then
          begin
            inc(ArLen);
            SetLength(BridgeArray,ArLen);
            with BridgeArray[ArLen-1] do
              begin
                CellCordX := CellCol;
                CellCordY := i;
                CellValue := StringGridTerminalList.Cells[CellCol,i];
              end;
          end;
    end;

  // finding out with numbers that are connected.
  i := 0;
  while i < Length(BridgeArray) do
    begin
      if BridgeArray[i].CellValue <> '' then
        begin
          TStrList := TStringList.Create();
          // get numbers from BridgeArray[i].MyValue into TStrList
          Split('-',BridgeArray[i].CellValue,TStrList);
          // replacing numbers with keynumber (first number)
          for q := 2 to TStrList.Count - 1 do
            for p := 0 to Length(BridgeArray) - 1 do
              if (('-'+TStrList[q]+'-') <> ('--')) then
                BridgeArray[p].CellValue := StringReplace(BridgeArray[p].CellValue,'-'+TStrList[q]+'-','-'+TStrList[1]+'-',[rfReplaceAll]);

          // only keep keynumber, from the allready processed value.
          BridgeArray[i].CellValue := copy(BridgeArray[i].CellValue,2,Length(BridgeArray[i].CellValue));
          BridgeArray[i].CellValue := '-' + copy(BridgeArray[i].CellValue,1,pos('-',BridgeArray[i].CellValue));
          FreeAndNil(TStrList);
        end;
      // get ready to next item
      inc(i);
      // it all item processed, than run through them to see if the all only contain
      // a key value (only one number)
      if not (i < Length(BridgeArray)) then
        for q := 0 to Length(BridgeArray)-1 do
          if (CountPos('-',BridgeArray[q].CellValue) > 2) then
            i := 0; // more the one number found, the run through them all again.
    end;

  // fill in correct laskenumber, by replacing with values from array,
  // and removing '-' character
  for i := 0 to Length(BridgeArray) - 1 do
    StringGridTerminalList.Cells[BridgeArray[i].CellCordX,BridgeArray[i].CellCordY] := StringReplace(BridgeArray[i].CellValue,'-','',[rfReplaceAll]);
  SetLength(BridgeArray,0);
end;

procedure TMainForm.CreateDocument;
begin
  PCSComObj.Documents.Add('','ProjectName');
end;

{cut out a part of a string. Accept a string as StartPos=CutOutString
 and another string as EndPos=EndCut}
function TMainForm.CutOut(CutOutString, MainString, EndCut: string): integer;
var
  PosCutEnd, PosCutOutString : integer;
begin
  PosCutOutString := Pos(CutOutString, MainString);
  if (PosCutOutString > -1) then
    begin
      inc(PosCutOutString,Length(CutOutString));
      MainString := copy(MainString,PosCutOutString,Length(MainString));
      if (pos(EndCut,MainString) > 0) then
        begin
          PosCutEnd := Pos(EndCut,MainString);
          try
            result := StrToInt(Copy(mainstring,0,PosCutEnd-1));
          except
            result := -1;
          end;
        end
      else
        begin
          try
            result := StrToInt(MainString);
          except
            result := -1;
          end;
        end;
    end
  else
    result := -1;
  if (result = -1) then
    ShowMessage('Could not cut out from "' + CutOutString + '"' + #13#10 +
                'to "' + EndCut + '"' + #13#10 +
                'from this main string "' + MainString + '"');
end;

{keep value inside valid range (EndAngle)}
procedure TMainForm.EditEndAngleExit(Sender: TObject);
begin
  try
    if (StrToInt(EditEndAngle.Text) > 360) then
      begin
        ShowStringInMemo('EndAngle must not be greater than 360');
        EditEndAngle.Text := '360';
      end;
    if (StrToInt(EditEndAngle.Text) < 0) then
      begin
        ShowStringInMemo('EndAngle must not be smaller than 0');
        EditEndAngle.Text := '0';
      end;
  except
    ShowStringInMemo('Not an integer in EndAngel');
  end;
end;

{keep value inside valid range (Rotation)}
procedure TMainForm.EditRotationExit(Sender: TObject);
begin
  try
    if (StrToInt(EditRotation.Text) > 360) then
      begin
        ShowStringInMemo('Rotation must not be greater than 360');
        EditRotation.Text := '360';
      end;
    if (StrToInt(EditRotation.Text) < 0) then
      begin
        ShowStringInMemo('Rotation must not be smaller than 0');
        EditRotation.Text := '0';
      end;
  except
    ShowStringInMemo('Not an integer in Rotation');
  end;
end;

procedure TMainForm.EditSymbolPosNoChange(Sender: TObject);
begin

end;

{draw color-rectangles in comboboxes}
procedure TMainForm.ComboboxColorDrawItem(Control: TWinControl; Index: Integer;
  Rect: TRect; State: TOwnerDrawState);
var
  ARect : TRect;
begin
  with (Control as TCombobox),Canvas do
    begin
      FillRect(Rect);
      if (integer(Items.Objects[Index]) <> ORD(Co_NoPrint)) then
        begin
          ARect.Left := Rect.Left+4;
          ARect.Top := Rect.Top+2;
          ARect.Right := Rect.Right-4;
          ARect.Bottom := Rect.Bottom-2;
          Brush.Color := PenColors[TPenColorIndex(Index)];
          FillRect(ARect);
          Brush.Color := clBlack;
          FrameRect(ARect);
        end
      else
        TextOut(Rect.Left+4,Rect.Top,'NP');
    end;
end;

{visual feedback}
procedure TMainForm.ComboBoxPageHeader1Change(Sender: TObject);
begin
  ComboBoxPageHeader2.Visible := ComboBoxPageHeader1.ItemIndex > 0;
  LabelPageHeader2.Visible := ComboBoxPageHeader1.ItemIndex > 0;
end;

{Fill comboboxes with pageheaders}
procedure TMainForm.ComboBoxPageHeader1DropDown(Sender: TObject);
var
  Headers        : OleVariant;
  i              : integer;
  CmbPageheader  : TCombobox;
begin
  {this eventhandler, goes for both comboboxes (Pageheader1 + 2}
  CmbPageheader := Sender as TComboBox;
  {To get pageheaders, currently availeble to Automation, is only
  possible through our old OLE-interface. This will be made availeble
  in the new interface}
  (PCSComObj as IApp).GetPageHeaders(Headers);
  CmbPageheader.Clear;
  CmbPageheader.Items.Add(NoneSelectedStr);
  {running through the availeble headers, and adding them to the combobox}
  for i:=VarArrayLowBound(Headers, 1) to VarArrayHighBound(Headers, 1) do
     CmbPageheader.Items.Add(Headers[i]);
end;

{if pagetab should have specific name}
procedure TMainForm.ComboboxPageTabNameChange(Sender: TObject);
begin
  EditTabName.Visible := ComboboxPageTabName.ItemIndex = 1;
  LabelSpecifyName.Visible := ComboboxPageTabName.ItemIndex = 1;
end;

procedure TMainForm.ComboBoxPredefinedProjectChange(Sender: TObject);
begin

end;

{update the memo at the bottom of the screen}
procedure TMainForm.ShowStringInMemo(ShowString: string);
begin
  MemoInfo.Lines.Add(TimeToStr(now) + ': ' + ShowString + '.');
  MemoInfo.Lines.Add('');
end;

{shows the position of the TrackBar just above it}
procedure TMainForm.TabSheet6Show(Sender: TObject);
var
  i : integer;
  ADrawing : IPCsDrawing;
begin
  // builing the comboboxes...
  ComboBoxOtherProject.Clear;
  ComboBoxOtherPage.Clear;
  ComboBoxDFSymbols.Clear;

  if NoActiveDocument then
    CreateDocument;

  ADrawing := PCSComObj.ActiveDocument.Drawing;
  ComboBoxDFCreateNew.Clear;

  if VarIsArray(ADrawing.ProjectDataDefs) then {array is created}
    for i := 0 to VarArrayHighBound(ADrawing.ProjectDataDefs,1) do
      ComboBoxOtherProject.Items.Add(ADrawing.ProjectDataDefs[i]);
  if ComboBoxOtherProject.Items.Count > 0 then
    begin
      ComboBoxOtherProject.ItemIndex := 0;
      ComboBoxDFCreateNew.Items.Add('Project Data Defs');
      ComboBoxDFCreateNew.ItemIndex := 0;
    end;

  if VarIsArray(ADrawing.PageDataDefs) then {array is created}
    for i := 0 to VarArrayHighBound(ADrawing.PageDataDefs,1) do
      ComboBoxOtherPage.Items.Add(ADrawing.PageDataDefs[i]);
  if ComboBoxOtherPage.Items.Count > 0 then
    begin
      ComboBoxOtherPage.ItemIndex := 0;
      ComboBoxDFCreateNew.Items.Add('Page Data Defs');
      ComboBoxDFCreateNew.ItemIndex := 0;
    end;

  if VarIsArray(ADrawing.SymbolDataDefs) then {array is created}
  for i := 0 to VarArrayHighBound(ADrawing.SymbolDataDefs,1) do
    ComboBoxDFSymbols.Items.Add(ADrawing.SymbolDataDefs[i]);
  if ComboBoxDFSymbols.Items.Count > 0 then
    begin
      ComboBoxDFSymbols.ItemIndex := 0;
      ComboBoxDFCreateNew.Items.Add('Symbol Data Defs');
      ComboBoxDFCreateNew.ItemIndex := 0;
    end;

end;

procedure TMainForm.TabSheetNewPageContextPopup(Sender: TObject;
  MousePos: TPoint; var Handled: Boolean);
begin

end;

{update the caption with the number of linestyle selected}
procedure TMainForm.TrackBarLineChange(Sender: TObject);
begin
  LabelNewLineStyle.Caption := IntToStr(TrackBarLine.Position);
end;

{shows the position of the TrackBar just above it}
procedure TMainForm.TrackBarLineModifyChange(Sender: TObject);
begin
  LabelModifyLineStyle.Caption := IntToStr(TrackBarLineModify.Position);
end;

{unfortunately, the type PsLineStyle, is not used at nummeric order,
  therefore this specific setup is needed}
function TMainForm.TranslateTrackPosToLineStyle(TrackPos: integer; Reversed : boolean): integer;
begin
  result := 0;
  if Reversed then
    begin
      case TrackPos of
        ls_Solid              : result := 0;
        ls_Dash               : result := 1;
        ls_DashDot            : result := 2;
        ls_DashDotDot         : result := 3;
        ls_DashDotDotDot      : result := 4;
        ls_DashDashDot        : result := 5;
        ls_DashDashDotDot     : result := 6;
        ls_DashDashDotDotDot  : result := 7;
        ls_Dot                : result := 8;
        ls_FinDot             : result := 9;
        ls_DualColor          : result := 10;
        ls_SolidBox           : result := 11;
        ls_EmptyBox           : result := 12;
        ls_FDiagonal05Box     : result := 13;
        ls_FDiagonal10Box     : result := 14;
        ls_BDiagonal05Box     : result := 15;
        ls_BDiagonal10Box     : result := 16;
        ls_DiagCross05Box     : result := 17;
        ls_VerticalBox        : result := 18;
        ls_HorizontalBox      : result := 19;
        ls_ZHashedBox         : result := 20;
        ls_RoundBox           : result := 21;
        ls_SolidFDiag50       : result := 22;
        ls_SolidCross50       : result := 23;
        ls_DashDotChannel     : result := 24;
        ls_DashDotChannel25   : result := 25;
        ls_DashDotChannel50   : result := 26;
        ls_ZFigur             : result := 27;
        ls_BDiagonal05        : result := 28;
        ls_BDiagonal10        : result := 29;
        ls_VertHoriBox        : result := 30;
        ls_DashedBox          : result := 47;
      end;
    end
  else
    begin
      case TrackPos of
        0  : result := ls_Solid;
        1  : result := ls_Dash;
        2  : result := ls_DashDot;
        3  : result := ls_DashDotDot;
        4  : result := ls_DashDotDotDot;
        5  : result := ls_DashDashDot;
        6  : result := ls_DashDashDotDot;
        7  : result := ls_DashDashDotDotDot;
        8  : result := ls_Dot;
        9  : result := ls_FinDot;
        10 : result := ls_DualColor;
        11 : result := ls_SolidBox;
        12 : result := ls_EmptyBox;
        13 : result := ls_FDiagonal05Box;
        14 : result := ls_FDiagonal10Box;
        15 : result := ls_BDiagonal05Box;
        16 : result := ls_BDiagonal10Box;
        17 : result := ls_DiagCross05Box;
        18 : result := ls_VerticalBox;
        19 : result := ls_HorizontalBox;
        20 : result := ls_ZHashedBox;
        21 : result := ls_RoundBox;
        22 : result := ls_SolidFDiag50;
        23 : result := ls_SolidCross50;
        24 : result := ls_DashDotChannel;
        25 : result := ls_DashDotChannel25;
        26 : result := ls_DashDotChannel50;
        27 : result := ls_ZFigur;
        28 : result := ls_BDiagonal05;
        29 : result := ls_BDiagonal10;
        30 : result := ls_VertHoriBox;
        31 : result := ls_DashedBox;
      end;
    end;
end;

{updates a row in Stringgrid, with bridge value}
procedure TMainForm.UpdateTerminalAdBridge(TerminalHandle: cardinal;
  IsInput: boolean; var BValue: integer);
var
  i : integer;
begin
  // running through symbolarray.
  for i := 1 to StringGridTerminalList.RowCount - 1 do
    // if value in stringgrid (StringGridTerminalList) = terminalhandle
    if (StrToInt(StringGridTerminalList.Cells[integer(TermHandle),i]) = Integer(TerminalHandle)) then
      begin
        // this is the terminal to update...
        if IsInput then
          begin
            // Add BValue to the cell
            if StringGridTerminalList.Cells[integer(InputBridge),i] = '' then
              StringGridTerminalList.Cells[integer(InputBridge),i] := '-' + IntToStr(BValue) + '-'
            else
              StringGridTerminalList.Cells[integer(InputBridge),i] := StringGridTerminalList.Cells[integer(InputBridge),i] + IntToStr(BValue) + '-';
          end
        else
          begin
            // Add BValue to the cell
            if StringGridTerminalList.Cells[integer(OutputBridge),i] = '' then
              StringGridTerminalList.Cells[integer(OutputBridge),i] := '-' + IntToStr(BValue) + '-'
            else
              StringGridTerminalList.Cells[integer(OutputBridge),i] := StringGridTerminalList.Cells[integer(OutputBridge),i] + IntToStr(BValue) + '-';
          end;
      end;
end;

{geting lineinformation into LineConnArray}
procedure TMainForm.CollectAllLines;
var
  i, t, CurPage     : integer;
  ALines            : IPCSLines;
begin
  CurPage := PCSComObj.ActiveDocument.ActivePage.Index;
  ALines := PCSComObj.ActiveDocument.ActivePage.Lines;
  SetLength(LineConnArray,0);
  SetLength(LineConnArray,ALines.Count);
  for i := 0 to ALines.Count - 1 do
    begin
      LineConnArray[i].LineHandle := ALines.Item[i].Handle;
      LineConnArray[i].LinePage   := CurPage;
      LineConnArray[i].LineLayer  := ALines.Item[i].Layer.Index;

      // Setting array for connectionpoints...
      SetLength(LineConnArray[i].LinePoints,ALines.Item[i].Points.Count);
      for t := 0 to ALines.Item[i].Points.Count - 1 do
        with LineConnArray[i].LinePoints[t] do
          begin
            PageVal := CurPage;
            LayerVal := ALines.Item[i].Layer.Index;;
            XValue := ALines.Item[i].Points.Item[t].Position.X;
            YValue := ALines.Item[i].Points.Item[t].Position.Y;
            ZValue := ALines.Item[i].Points.Item[t].Position.Z;
          end;
    end;
  LabelLinesFound.Caption := IntToStr(ALines.Count);
end;

{getting symbolinformation into SymConnArray}
procedure TMainForm.CollectAllSymbols;
var
  i, t, ArraySymCount, ArrayBranchCount, CurPage     : integer;
  ASymbols                                           : IPCSSymbols;
begin
  CurPage := PCSComObj.ActiveDocument.ActivePage.Index;
  ASymbols := PCSComObj.ActiveDocument.ActivePage.Symbols;
  SetLength(SymConnArray,0);
  SetLength(SymBranchArray,0);
  ArraySymCount     := 0;
  ArrayBranchCount  := 0;
  for i := 0 to ASymbols.Count - 1 do
    if ASymbols.Item[i].SymbolType1 <> st_Branch then
      begin
        SetLength(SymConnArray,ArraySymCount+1);
        SymConnArray[ArraySymCount].SymHandle := ASymbols.Item[i].Handle;
        SymConnArray[ArraySymCount].SymPage   := CurPage;
        SymConnArray[ArraySymCount].SymName   := ASymbols.Item[i].FullName;
        SymConnArray[ArraySymCount].SymLayer  := ASymbols.Item[i].Layer.Index;
        SetLength(SymConnArray[ArraySymCount].ConnPoints,ASymbols.Item[i].Connections.Count);
        for t := 0 to ASymbols.Item[i].Connections.Count - 1 do
          with SymConnArray[ArraySymCount].ConnPoints[t] do
            begin
              PageVal   := CurPage;
              LayerVal  := ASymbols.Item[i].Layer.Index;
              PointName := ASymbols.Item[i].Connections.Item[t].PinName;
              XValue    := ASymbols.Item[i].Connections.Item[t].Position.X;
              YValue    := ASymbols.Item[i].Connections.Item[t].Position.Y;
              ZValue    := ASymbols.Item[i].Connections.Item[t].Position.Z;
            end;
        inc(ArraySymCount);
      end
    else
      begin
        SetLength(SymBranchArray,ArrayBranchCount+1);
        SymBranchArray[ArrayBranchCount].SymHandle := ASymbols.Item[i].Handle;
        SymBranchArray[ArrayBranchCount].SymPage   := CurPage;
        SymBranchArray[ArrayBranchCount].SymName   := ASymbols.Item[i].FullName;
        SymBranchArray[ArrayBranchCount].SymLayer  := ASymbols.Item[i].Layer.Index;
        SymBranchArray[ArrayBranchCount].State     := ASymbols.Item[i].State;
        // Setting array for connectionpoints...
        SetLength(SymBranchArray[ArrayBranchCount].ConnPoints,ASymbols.Item[i].Connections.Count);
        for t := 0 to ASymbols.Item[i].Connections.Count - 1 do
          with SymBranchArray[ArrayBranchCount].ConnPoints[t] do
            begin
              PageVal   := CurPage;
              LayerVal  := ASymbols.Item[i].Layer.Index;
              PointName := ASymbols.Item[i].Connections.Item[t].PinName;
              XValue    := ASymbols.Item[i].Connections.Item[t].Position.X;
              YValue    := ASymbols.Item[i].Connections.Item[t].Position.Y;
              ZValue    := ASymbols.Item[i].Connections.Item[t].Position.Z;
            end;
        inc(ArrayBranchCount);
      end;
  LabelSymbolsFoundAll.Caption := IntToStr(ArraySymCount+ArrayBranchCount);
  LabelBranchSymbols.Caption := IntToStr(ArrayBranchCount);
  LabelNotBranchSymbols.Caption := IntToStr(ArraySymCount);
end;

{gets and list number of opened projects}
procedure TMainForm.ButtonGetConnectionsClick(Sender: TObject);
var
  i,t             : integer;
  SymAndPinName   : string;
begin
  ButtonGetConnections.Enabled := false;
  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;

      // clear stringgrid
      ClearStringGrid();

      // get all symbols into array of record (both Normal symbols and branch's)
      CollectAllSymbols();

      // get all lines into array of record
      CollectAllLines();

      // find joins, that this program does handle yet...
      FindLineJoins();
      {Go through each symbol. Look at that symbols connectionspoint. If a line
      is connected to a connectionpoint, that follow that line, to what its
      connected to}
      SetLength(AvoidInfiniteLoop,0);
      for i := 0 to Length(SymConnArray) - 1 do
        begin
          // all the connectionpoints of each symbol
          if Length(SymConnArray[i].ConnPoints) > 0 then
            // for each connection point
            for t := 0 to Length(SymConnArray[i].ConnPoints) - 1 do
              begin
                // symbolName + pinname
                SymAndPinName := SymConnArray[i].SymName + ':' + SymConnArray[i].ConnPoints[t].PointName;

                // follow the SymConnArray[i].ConnPoints[t]
                CheckAndFollowThisPoint(SymAndPinName,SymConnArray[i].ConnPoints[t]);
              end;
        end; // End of symbol-loop
    end;
  // delete the last empty row in grid
  if (StringGridConnections.Cells[0,(StringGridConnections.RowCount-1)] = '') then
    StringGridConnections.RowCount := StringGridConnections.RowCount -1;

  // if Joins were found, then show message, that the connectionlist is VERY likely wrong...
  if LabelJoinsFound.Caption <> '0'  then
    ShowMessage('Result is most likely "NOT CORRECT", as Joins were found...');

  ButtonGetConnections.Enabled := true;
end;

{showing numbers of opened projects}
procedure TMainForm.ButtonGetNumberOfOpenedProjectsClick(Sender: TObject);
var
  i : integer;
begin
  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;

      ListBoxNumberOfOpenedProjects.Items.BeginUpdate;
      ListBoxNumberOfOpenedProjects.Clear;
      {running through each document, and gets its name/path}
      for i:= 0 to PCSComObj.Documents.Count-1 do
        ListBoxNumberOfOpenedProjects.Items.Add('Document #' + IntToStr(i) + ' = ' + PCSComObj.Documents[i].FullName);
      ListBoxNumberOfOpenedProjects.Items.EndUpdate;
      ShowStringInMemo('Number of opened projects found = ' + IntToStr((PCSComObj.Documents.Count)));
    end;
  ButtonSetAsActiveProject.Enabled := false;
end;

{gets and list number of pages in a project}
procedure TMainForm.ButtonGetNumberOfPagesInProjectClick(Sender: TObject);
var
  i : integer;
  FirStr, SecStr, ThiStr : string;
begin
  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;

      ListBoxNumberOfPages.Items.BeginUpdate;
      ListBoxNumberOfPages.Clear;
      {name of the project}
      ListBoxNumberOfPages.items.Add('Project=' + PCSComObj.ActiveDocument.FullName);
      {Running through each project, and show number of pages}
      for i:= 0 to PCSComObj.ActiveDocument.Drawing.Pages.Count-1 do
        begin
          FirStr := FixStringLength('Page = ' + IntToStr(i),15);
          {get title on each page}
          SecStr := FixStringLength('Title = ' + PCSComObj.ActiveDocument.Drawing.Pages[i].Title,30);
          ThiStr := 'PageNo = ' + PCSComObj.ActiveDocument.Drawing.Pages[i].PageNumber;
          ListBoxNumberOfPages.Items.Add(FirStr+SecStr + ThiStr);
        end;
      ListBoxNumberOfPages.Items.EndUpdate;
      ShowStringInMemo('Number of pages in project = ' + IntToStr(PCSComObj.ActiveDocument.Drawing.Pages.Count));
    end;
  ButtonSetPageasActive.Enabled := false;
end;

{getting symbols into listbox}
procedure TMainForm.ButtonGetSymbolsClick(Sender: TObject);
var
  ASymbols : IPCsSymbols;
  i : integer;
begin
  {make sure the is a document open}
  if NoActiveDocument then
    CreateDocument;

  ListBoxSymbolsDatafields.Items.BeginUpdate;
  ListBoxSymbolsDatafields.Clear;
  {getting all symbols}
  ASymbols := PCSComObj.ActiveDocument.ActivePage.Symbols;
  for i := 0 to ASymbols.Count - 1 do
    ListBoxSymbolsDatafields.Items.Add(ASymbols[i].FullName + '     Handle=' + IntToStr(ASymbols[i].Handle));
  ListBoxSymbolsDatafields.Items.EndUpdate;
end;

{creates a terminal list, with connection information}
procedure TMainForm.ButtonGetTerminalsClick(Sender: TObject);
var
  x, y, i, f, u                                             : integer;
  ListOfTerminals                                           : IPCsList;
  EmptyStr                                                  : string;
  TerminalName, ConnectedToInput, ConnectedToOutput,
  CablesToInput,CablesToOutput, WireToInput, WireToOutput   : IPCsSymbol;
  IndexTerm, IndexTermInput, IndexTermOutput,
  IndexCableInput, IndexCableOutput, IndexWireInput,
  IndexWireOutput, TopOfLoop, ThisPoint, BridgeVal,
  PosX,PosY,PosZ, MisCount                                  : integer;
  BridgeAdded                                               : boolean;
begin
{checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;

      ButtonGetTerminals.Enabled := false;

      MisCount := 0;
      // Clear StringGridTerminalList
      with StringGridTerminalList do
        for x := 0 to ColCount - 1 do
          for y := 1 to RowCount - 1 do
            Cells[x,y] := '';
      StringGridTerminalList.RowCount := 2;

      // Build array, with the positions of connectionpoint of each terminal
      // array is variable = SymArray, and is of type TSymbolArray
      GetTerminalData();
      // To get the lines the can exist between terminals, a IPCsList is needed.
      ListOfTerminals := PCSComObj.ActiveDocument.Drawing.Lists.CreateTerminalList(EmptyStr);

      {
      PCSComObj.ActiveDocument.Drawing.Lists.CreateCableList(EmptyStr);
      PCSComObj.ActiveDocument.Drawing.Lists.CreateTerminalList(EmptyStr);
      PCSComObj.ActiveDocument.Drawing.Lists.CreatePartList(EmptyStr);
      PCSComObj.ActiveDocument.Drawing.Lists.CreateComponentList(EmptyStr);

      TNetDataType   = (nd_Signal, nd_FromComp, nd_ToComp, nd_WireNo);
      TConnsDataType = (cd_Conn);
      TPLCDataType   = (pd_PLC, pd_PLCRef, pd_WireNo, pd_Signal, pd_ToComp, pd_Path);
      TTermDataType  = (td_ExtComp, td_ExtWireNo, td_ExtCable, td_ExtSignal, td_OutConn,
                        td_InConn, td_InSignal, td_InCable, td_InWireNo, td_InComp, td_JumperLink);
      TCableDataType = (cd_ToComp, cd_WireNo, cd_Cable, cd_FrCable, cd_Signal, cd_FromComp);
      }

      // make correct number of rows in StringGridTerminalList
      StringGridTerminalList.RowCount := ListOfTerminals.Count + 1;

      // information to memo
      ShowStringinMemo('Length of array with terminal = ' + inttostr(Length(SymArray)));
      // running through list of terminals. Please notice, that at terminal can be represented
      // more than onces, if its connected to more than 1 other terminal...
      for i := 0 to ListOfTerminals.Count - 1 do
        begin
          with ListOfTerminals.Item[i] do
            begin
              // fills different Symbol-variables, with information, that shows a symbols connection
              // to other symbols

              // td_InConn, is used for getting name of Terminal, and gives symbol to TerminalName
              GetDataSymbol(integer(td_InConn),0,TerminalName,IndexTerm);
              // td_InComp, which termnial is it connected to (input), and gives symbol to ConnectedToInput
              GetDataSymbol(integer(td_InComp),0,ConnectedToInput,IndexTermInput);
              // td_ExtComp, which termnial is it connected to (output), and gives symbol to ConnectedtoOutput
              GetDataSymbol(integer(td_ExtComp),0,ConnectedToOutput,IndexTermOutput);
              // td_InCable, what is the name of the cable (input), and gives symbol to CablesToInput
              GetDataSymbol(integer(td_InCable),0,CablesToInput,IndexCableInput);
              // td_InCable, what is the name of the cable (output), and gives symbol to CablesToOutput
              GetDataSymbol(integer(td_ExtCable),0,CablesToOutput,IndexCableOutput);
              // td_InWireNo, what is the name of the WireNo (input)
              GetDataSymbol(integer(td_InWireNo),0,WireToInput,IndexWireInput);
              // td_ExtWireNo, what is the name of the WireNo (ouput)
              GetDataSymbol(integer(td_ExtWireNo),0,WireToOutput,IndexWireOutput);
            end;

          // now, insert all these information in StringGrid (StringGridTerminalList)
          if TerminalName = nil then
            inc(MisCount)   // used for keep track of what cell to write in
          else
            begin
              // terminal handle:
              StringGridTerminalList.Cells[integer(TermHandle),i+1-MisCount] := IntToStr(TerminalName.Handle);
              // terminal name + pinname
              StringGridTerminalList.Cells[integer(TermName),i+1-MisCount] := TerminalName.FullName + ':' + TerminalName.Connections[0].PinName;

              // input name
              if ConnectedToInput <> nil then
                begin
                  // show which terminal its connected to + pinname
                  StringGridTerminalList.Cells[integer(InputName),i+1-MisCount] := ConnectedToInput.FullName + ':' + ConnectedToInput.Connections[IndexTermInput].PinName;
                  // if it's a cable, then show the cable name + pinname
                  if CablesToInput <> nil then
                    StringGridTerminalList.Cells[integer(InputCabel),i+1-MisCount] := CablesToInput.FullName + ':' + CablesToInput.Connections[0].PinName;
                  // if it's a WireNo then show the WireNo-name
                  if WireToInput <> nil then
                    StringGridTerminalList.Cells[integer(InputWireNo),i+1-MisCount] := WireToInput.FullName;
                end;

              // output name
              if ConnectedToOutput <> nil then
                begin
                  // show which terminal its connected to + pinname
                  StringGridTerminalList.Cells[integer(OutputName),i+1-MisCount] := ConnectedToOutput.FullName + ':' + ConnectedToOutput.Connections[IndexTermOutput].PinName;
                  // if it's a cable, then show the cable name + pinname
                  if CablesToOutput <> nil then
                    StringGridTerminalList.Cells[integer(Outputcabel),i+1-MisCount] := CablesToOutput.FullName + ':' + CablesToOutput.Connections[0].PinName;
                  // if it's a WireNo then show the WireNo-name
                  if WireToOutput <> nil then
                    StringGridTerminalList.Cells[integer(OutputWireNo),i+1-MisCount] := WireToOutput.FullName;
                end;
            end;
        end;

      // another type of connection is a bridge (Jumpering link) (laske)
      // to get this type, it's nessecary to run through all lines,
      // and check if the line is a JumperingLink
      // All terminals that have the same bridge value, is bridges together.
      if (Length(SymArray) > 0) then
        begin
          // only continue, if there where at least 1 terminal.
          BridgeVal := 1;
          for i:= 0 to PCSComObj.ActiveDocument.ActivePage.Lines.Count-1 do
            if PCSComObj.ActiveDocument.ActivePage.Lines.Item[i].IsJumperingLink then
              begin
                // number of points in that line
                TopOfLoop := PCSComObj.ActiveDocument.ActivePage.Lines.Item[i].Points.Count;
                ThisPoint := 0;
                // init BridgeAdded
                BridgeAdded := false;
                // for each point in that line
                while (ThisPoint < TopOfLoop) do
                  begin
                    // get the position of that point in that line.
                    PCSComObj.ActiveDocument.ActivePage.Lines.Item[i].Points.Item[ThisPoint].GetPosition(PosX,PosY,PosZ);
                    // running through array of terminals, to see if a matching position in
                    // a connection point, exist...
                    for f := 0 to Length(SymArray) - 1 do
                      begin
                        // input
                        for u := 0 to Length(SymArray[f].InPos) - 1 do
                          // same position for a point in line, as for a terminals connectionpoint?
                          if ((PosX=SymArray[f].InPos[u].XValue) and (PosY=SymArray[f].InPos[u].YValue) and (PosZ=SymArray[f].InPos[u].ZValue)) then
                            begin
                              // same position found, then add the information, that its a bridge

                              // update the StringGridTerminalList, (find correct terminal,
                              // and update information, with number BridgeVal)
                              UpdateTerminalAdBridge(SymArray[f].SymHandle,true,BridgeVal);
                              BridgeAdded := true;
                            end;
                        // output
                        for u := 0 to Length(SymArray[f].OutPos) - 1 do
                          // same position for a point in line, as for a terminals connectionpoint?
                          if ((PosX=SymArray[f].OutPos[u].XValue) and (PosY=SymArray[f].OutPos[u].YValue) and (PosZ=SymArray[f].OutPos[u].ZValue)) then
                            begin
                              // same position found, then add the information, that its a bridge

                              // update the StringGridTerminalList, (find correct terminal,
                              // and update information, with number BridgeVal)
                              UpdateTerminalAdBridge(SymArray[f].SymHandle,false,BridgeVal);
                              BridgeAdded := true;
                            end;
                      end;
                    inc(ThisPoint);
                  end;

                // if bridge allready added, the increment BridgeVal
                if BridgeAdded then
                  inc(BridgeVal);
              end;
        end;
      ShowStringinMemo('Finished creating terminal list');
    end;

  // clear empty rows, (if any)
  i := StringGridTerminalList.RowCount-1;
  while StringGridTerminalList.Cells[0,i] = '' do
    begin
      StringGridTerminalList.RowCount := i;
      dec(i);
    end;

  // finding out witch terminals that are bridges together...
  ConvertingCellsToBridge();

  ButtonGetTerminals.Enabled := true;
  StringGridTerminalList.SetFocus;
end;

{userfeedback}
procedure TMainForm.ButtonLinesClearClick(Sender: TObject);
begin
  ListBoxLines.Clear;
  ButtonCreateNewLine.Enabled := false;
end;

{move a specific symbol}
procedure TMainForm.ButtonModifyLineAddPointClick(Sender: TObject);
var
  ItemInListbox : integer;
  ALine        : IPCsLine;
begin
  {get index of that line that is selected}
  ItemInListbox := CutOut('#',ListBoxModifyLines.Items[ListBoxModifyLines.ItemIndex],' ');
  {assign line, from activepage}
  ALine := PCSComObj.ActiveDocument.ActivePage.Lines.Item[ItemInListbox];
  {add new point (after all the other points in that line), with value = 0,0,0}
  ALine.Points.Add;
  {show the new point in line}
  ListBoxModifyLinesClick(sender);
  {repaint diagram}
  PCSComObj.Redraw;
  {disable update button}
  ButtonUpdateLine.Enabled := false;
end;

{move a symbol's position}
procedure TMainForm.ButtonMoveSymbolClick(Sender: TObject);
var
  FindHandle       : cardinal;
  MoveUnit         : integer;
  PosX, PosY, PosZ : integer;
  Symbol           : IPCsSymbol;
begin
  {get handle of symbol}
  FindHandle := GetSymbolHandleFromListBox(ListBoxSymbols);
  MoveUnit := StrToInt(EditMovingUnits.Text);
  {assign symbol, from handle}
  Symbol := GetSymbolFromSymbolHandle(FindHandle);
  {get position of symbol}
  Symbol.GetPosition(PosX,PosY,PosZ);
  {modify values of the position}
  case ComboboxDirection.ItemIndex of
    0 : Dec(PosX,MoveUnit);
    1 : Inc(PosX,MoveUnit);
    2 : Inc(PosY,MoveUnit);
    3 : Dec(PosY,MoveUnit);
  end;
  {set new position of symbol}
  Symbol.SetPosition(PosX,PosY,PosZ);
  {repaint diagram}
  PCSComObj.Redraw;
  ShowStringInMemo('Symbol handle= ' + IntToStr(FindHandle) + ' is moved ' + EditMovingunits.Text + 'units. Direction=' + ComboboxDirection.Text);
end;

{list all symbols on a page}
procedure TMainForm.ButtonFindSymbolsALLClick(Sender: TObject);
begin
  FindAndHightlightSymbols(TRUE);
end;

{getting all symbols on page}
procedure TMainForm.ButtonFindSymbolsOnPageClick(Sender: TObject);
var
  i                      : integer;
  PosX, PosY, PosZ       : integer;
  FirStr, SecStr, ThiStr : string;
begin
  {checks connectionstatus. if not connected, then it tries to connect}
  Screen.Cursor := crHourGlass;
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;
      ButtonMoveSymbol.Enabled := false;
      ButtonChangeSymbolName.Enabled := false;
      ButtonDeleteSymbol.Enabled := false;
      ButtonSetCompGrpAndPos.Enabled := false;
      EditSymbolName.Text := '';
      EditSymbolFileName.Text := '';
      EditInDirectory.Text := '';
      EditSymbolType.Text := '';
      ListBoxSymbols.Items.BeginUpdate;
      ListBoxSymbols.Clear;
      EditCompGrpNr.Text := '';
      EditSymbolPosNo.Text := '';
      ListBoxSymbolConnections.Clear;
      {running through all symbols}
      for i:= 0 to PCSComObj.ActiveDocument.ActivePage.Symbols.Count-1 do
        begin
          {get the position of the symbol}
          PCSComObj.ActiveDocument.ActivePage.Symbols[i].GetPosition(PosX,PosY,PosZ);
          {name of the symbol}
          FirStr := FixStringLength('Symbol=' + PCSComObj.ActiveDocument.ActivePage.Symbols[i].FullName,40);
          SecStr := FixStringLength('Cords(X=' + IntToStr(PosX) + ',Y=' + IntToStr(PosY) + ',Z=' + IntToStr(PosZ) + ')',40);
          {get handle to the symbol}
          ThiStr := 'Handle=' + IntToStr(PCSComObj.ActiveDocument.ActivePage.Symbols[i].Handle);
          ListBoxSymbols.Items.Add(FirStr+SecStr+ThiStr);
        end;
      ListBoxSymbols.Items.EndUpdate;
      ShowStringInMemo('Number of symbols found: ' + IntToStr(PCSComObj.ActiveDocument.ActivePage.Symbols.Count));
    end;
  Screen.Cursor := crDefault;;
end;

{finds the dimension of all symbols}
procedure TMainForm.ButtonFindDimensionClick(Sender: TObject);
const
  RectTypes : array [0..3] of integer = (rt_Raw,rt_WithTexts,rt_Relative,rt_WithReference);
var
  i, LongestName          : integer;
  Rect                    : TPCsRect;
  FirStr, SecStr, ThiStr  : string;
begin
  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;

      ListBoxDimensions.Items.BeginUpdate;
      ListBoxDimensions.Clear;
      {running through all symbols on a page}
      LongestName := 0;
      for i:= 0 to PCSComObj.ActiveDocument.ActivePage.Symbols.Count-1 do
        if Length(PCSComObj.ActiveDocument.ActivePage.Symbols[i].FullName) > longestname then
          longestName := Length(PCSComObj.ActiveDocument.ActivePage.Symbols[i].FullName);

      for i:= 0 to PCSComObj.ActiveDocument.ActivePage.Symbols.Count-1 do
        begin
          {get dimension of the symbol}
          Rect := PCSComObj.ActiveDocument.ActivePage.Symbols[i].GetRect(RectTypes[ComboBoxRectType.ItemIndex]);
          {name of the symbol}
          FirStr := FixStringLength('Symbol=' + PCSComObj.ActiveDocument.ActivePage.Symbols[i].FullName,LongestName+9);
          SecStr := FixStringLength('Cords(Left=' + IntToStr(Rect.Left) + ',Top=' + IntToStr(Rect.Top) + ',Right=' + IntToStr(Rect.Right) + ',Bottom=' + IntToStr(Rect.Bottom) + ')',60);
          {get handle to the symbol}
          ThiStr := ' Handle=' + IntToStr(PCSComObj.ActiveDocument.ActivePage.Symbols[i].Handle);
          ListBoxDimensions.Items.Add(FirStr+SecStr+ThiStr);
        end;
      ListBoxDimensions.Items.EndUpdate;
      ShowStringInMemo('Number of symbols found =  ' + IntToStr(PCSComObj.ActiveDocument.ActivePage.Symbols.Count));
    end;
end;

{find lines on page}
procedure TMainForm.ButtonFindLayersClick(Sender: TObject);
var
  MyLayers : IPCsLayers;
  i, TheActiveLayer : integer;
  FirStr, SecStr, ThiStr : string;
begin
  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      if NoActiveDocument then
        CreateDocument;
      ListBoxLayers.Items.BeginUpdate;
      ListBoxLayers.Clear;
      MyLayers := PCSComObj.ActiveDocument.Drawing.Layers;
      // only set buttons visible if there are some layers
      ButtonSetLayerVisible.Enabled := (MyLayers.Count > 0);
      ButtonSetLayerInvisible.Enabled := (MyLayers.Count > 0);
      // getting the active layer index
      TheActiveLayer := PCSComObj.ActiveDocument.Drawing.ActiveLayer.Index;
      for i := 0 to MyLayers.Count - 1 do
        if MyLayers.Item[i].IsUsed then
          begin
            // if loop reach the active layer, then show shoe *
            if (TheActiveLayer = MyLayers.Item[i].Index) then
              FirStr := '* '
            else
              FirStr := '  ';
            // adding the name to the output
            FirStr := FirStr + 'Name=' + FixStringLength(MyLayers.Item[i].Name,35);
            if MyLayers.Item[i].Visible then
              SecStr := 'Visible = true'
            else
              SecStr := 'Visible = false';
            SecStr := FixStringLength(SecStr,20);
            // adding index
            ThiStr := 'Index=' + IntToStr(MyLayers.Item[i].Index);
            // posting in listbox.
            ListBoxLayers.Items.Add(FirStr + SecStr + ThiStr);
          end;
      ListBoxLayers.Items.EndUpdate;
    end;
end;

procedure TMainForm.ButtonFindLinesClick(Sender: TObject);
var
  i, p, TopOfLoop              : integer;
  OutStr, TempStr, MorePoints  : string;
  PosX, PosY, PosZ             : integer;
begin
  ButtonUpdateLine.Enabled := false;
  ButtonDeleteLine.Enabled := false;
  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;

      ListBoxModifyLines.Items.BeginUpdate;
      ListBoxLinePoints.Items.BeginUpdate;

      ListBoxModifyLines.Clear;
      ListBoxLinePoints.Clear;
      {running through all lines on page}
      for i:= 0 to PCSComObj.ActiveDocument.ActivePage.Lines.Count-1 do
        begin
          {get number of point in actual line}
          TopOfLoop := PCSComObj.ActiveDocument.ActivePage.Lines.Item[i].Points.Count;
          if (TopOfLoop > 3) then
            begin
              TopOfLoop := 3; {no room to display more in the listbox}
              MorePoints := '    (.....)'
            end
          else
            MorePoints := '';
          p := 0;
          OutStr := 'Line #' + IntToStr(i);
          {get position from each point in that line}
          while (p < TopOfLoop) do
            begin
              {get position}
              PCSComObj.ActiveDocument.ActivePage.Lines.Item[i].Points.Item[p].GetPosition(PosX,PosY,PosZ);
              TempStr := FixStringLength('   (' + IntToStr(PosX) + ',' + IntToStr(PosY) + ',' + IntToStr(PosZ) + ')',25);
              OutStr := OutStr + TempStr;
              inc(p);
            end;
          if PCSComObj.ActiveDocument.ActivePage.Lines.Item[i].IsJumperingLink then
            OutStr := OutStr + ' (Laske)';
          if PCSComObj.ActiveDocument.ActivePage.Lines.Item[i].IsElectrical then
            OutStr := OutStr + ' (Elektrisk)';
          ListBoxModifyLines.Items.Add(OutStr + MorePoints);
        end;

      ListBoxLinePoints.Items.EndUpdate;
      ListBoxModifyLines.Items.EndUpdate;

      ShowStringInMemo('Lines found= ' + IntToStr(PCSComObj.ActiveDocument.ActivePage.Lines.Count));
    end;
end;

{set the selected project, to be active}
procedure TMainForm.ButtonSetAsActiveProjectClick(Sender: TObject);
var
  str    : string;
  Docnum : integer;
begin
  str := ListBoxNumberOfOpenedProjects.Items.Strings[ListBoxNumberOfOpenedProjects.ItemIndex];
  {get document number}
  Docnum := StrToInt(trim(Copy(str,Pos('#',str)+1,pos(' = ',str)-Pos('#',str))));
  {activate that project}
  PCSComObj.Documents[Docnum].Activate;

  {when activating a project, that project become document[0].
   therefore this call is made to update documentnumbers}
  ButtonGetNumberOfOpenedProjectsClick(Sender);
  ButtonSetAsActiveProject.Enabled := false;
  ShowStringInMemo('Project number ' + IntToStr(Docnum) + ' is set as active');
end;

{Set Pageheader on active page}
procedure TMainForm.ButtonSetOnActivePageClick(Sender: TObject);
begin
  {make sure the is a document open}
  if NoActiveDocument then
    CreateDocument;

  if (ComboBoxPageHeader1.ItemIndex > 0) then
    {Set the first Page header}
    (PCSComObj.ActiveDocument.ActivePage as IPCsPage2).LoadPageHeader2(0,ComboBoxPageHeader1.Items.Strings[ComboBoxPageHeader1.ItemIndex]);
  if (ComboBoxPageHeader2.Visible and (ComboBoxPageHeader2.ItemIndex > 0)) then
    {Set the second Page header}
    (PCSComObj.ActiveDocument.ActivePage as IPCsPage2).LoadPageHeader2(1,ComboBoxPageHeader2.Items.Strings[ComboBoxPageHeader2.ItemIndex]);
  if (ComboBoxPageHeader1.ItemIndex > 0) then;
    begin
      PCSComObj.Redraw;
      ShowStringInMemo('Pageheader 1 set to = ' + ComboBoxPageHeader1.Items.Strings[ComboBoxPageHeader1.ItemIndex]);
    end;
  if (ComboBoxPageHeader2.Visible and (ComboBoxPageHeader2.ItemIndex > 0)) then
    ShowStringInMemo('Pageheader 2 set to = ' + ComboBoxPageHeader2.Items.Strings[ComboBoxPageHeader2.ItemIndex]);
end;

{set selected page as active}
procedure TMainForm.ButtonSetPageasActiveClick(Sender: TObject);
var
  str    : string;
  Docnum : integer;
begin
  str := ListBoxNumberOfPages.Items.Strings[ListBoxNumberOfPages.ItemIndex];
  {get document number, for selected page}
  Docnum := StrToInt(trim(Copy(str,Pos('=',str)+1,pos('Title',str)-Pos('=',str)-1)));
  {set activepage to variable just found}
  PCSComObj.ActiveDocument.ActivePage := PCSComObj.ActiveDocument.Drawing.Pages[Docnum];
  ShowStringInMemo('Page number ' + IntToStr(Docnum) + ' is set as active');
end;

{update the selected point with new values.}
procedure TMainForm.ButtonUpdateLineClick(Sender: TObject);
var
  LineNr, PointNr  : integer;
  ALines           : IPCsLines;
  PosX, PosY, PosZ : integer;
begin
  ButtonUpdateLine.Enabled := false;

  try
    {get values from textboxes}
    PosX := StrToInt(EditLineModifyX.Text);
    PosY := StrToInt(EditLineModifyY.Text);
    PosZ := StrToInt(EditLineModifyZ.Text);
  except
    ShowStringInMemo('Not integervalues in all coordinate textboxes');
    exit;
  end;

  try
    {get linenumber}
    LineNr := CutOut('#',ListBoxModifyLines.Items[ListBoxModifyLines.ItemIndex],' ');
    {get pointnumber}
    PointNr := CutOut('#',ListBoxLinePoints.Items[ListBoxLinePoints.ItemIndex],' ');
    {assign line, from activepage}
    ALines := PCSComObj.ActiveDocument.ActivePage.Lines;
    {with that specific line}
    With (ALines.Item[LineNr]) do
      begin
        {sets new position at specific point, with values from textboxes}
        Points.Item[PointNr].SetPosition(PosX,PosY,PosZ);
        {sets color}
        Color := integer(PenIndexToColor(ComboboxLineModifyColor.ItemIndex));
        {sets line width}
        PenWidth := StrToInt(EditModifyLineWidth.Text);
        {sets line distance}
        MultilineDistance := StrToInt(EditModifyLineDistance.Text);
        {make line as curves}
        IsCurvedLine := CheckBoxModifyLineCurved.Checked;
        {style of line}
        LineStyle := TranslateTrackPosToLineStyle(TrackBarLineModify.Position,false);
      end;

    {repaint diagram}
    PCSComObj.Redraw;
    ShowStringInMemo('Line updated');
    {show the new point in line}
    ListBoxModifyLinesClick(sender);
  except
    ShowStringInMemo('Item not selected in listbox(es)');
  end;
end;

{updates a text}
procedure TMainForm.ButtonUpdateTextClick(Sender: TObject);
var
  MyHandle : cardinal;
  AText    : IPCsText;
begin
  {get handle to selected text and update it with new values}
  MyHandle := CutOut('Handle=',ListBoxFoundTexts.Items[ListBoxFoundTexts.ItemIndex],' ');
  {assign text, from handle}
  AText := GetTextFromHandle(MyHandle);
  {set textvalue}
  AText.Value := EditFoundText.Text;
  {set color}
  AText.Color := integer(PenIndexToColor(ComboboxfoundTextColor.ItemIndex));
  {set rotation}
  AText.Rotation := StrToInt(EditFoundTextAngle.Text);
  {set font options}
  AText.Font.Name := ComboboxFoundTextFont.Items[ComboboxFoundTextFont.ItemIndex];
  AText.Font.Bold := CheckboxFoundTextBold.Checked;
  AText.Font.Italic := CheckboxFoundTextItalic.Checked;
  FormatSettings.DecimalSeparator := ',';
  {height and width are integers}
  AText.Font.Height := StrToInt(EditFoundTextFontHeight.Text);
  AText.Font.Width := StrToInt(EditFoundTextFontWidth.Text);
  {set position}
  AText.SetPosition(StrToInt(EditTextFoundX.Text),StrToInt(EditTextFoundY.Text),StrToInt(EditTextFoundZ.Text));
  FormatSettings.DecimalSeparator := '.';
  ShowStringInMemo('A text has been updated');
  {repaint diagram}
  PCSComObj.Redraw;
end;

{creates a new circle}
procedure TMainForm.ButtonCreateNewArcsClick(Sender: TObject);
var
  ADocument : IPCsDocument;
  APage     : IPCsPage;
  AArcs     : IPCsArc;
  APt       : TPCsPoint3D;
  TempFloat : extended;
begin
  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;
    
      {active project}
      ADocument := PCSComObj.ActiveDocument;
      {active page}
      APage := ADocument.ActivePage;

      {assigning valus to TPCsPoint3D}
      APt.X := StrToInt(EditNewArcsX.Text);
      APt.Y := StrToInt(EditNewArcsY.Text);
      APt.Z := StrToInt(EditNewArcsZ.Text);
      {create circle}
      AArcs := APage.Arcs.Add(Apt,StrToInt(EditRadius.Text),StrToInt(EditStartAngle.Text)*10,StrToInt(EditEndAngle.Text)*10);
      {set is filled}
      AArcs.IsFilled := CheckBoxFilled.Checked;
      {set color}
      AArcs.Color := integer(PenIndexToColor(ComboboxColor.ItemIndex));
      {set elipsefactor}
      AArcs.EllipseFactor := StrToFloat(ComboboxElipseFactor.Text);

      {pensize is an integer. The numbers in the Combobox are exactly the same as in
      Automation. These numbers must be multiplied by 1000}
      TempFloat := StrToFloat(ComboboxPenWidth.Items[ComboboxPenWidth.ItemIndex]);
      AArcs.PenWidth := trunc(TempFloat * 1000);
      {set rotation}
      AArcs.Rotation := trunc(StrToFloat(EditRotation.Text)*10);

      {repaint diagram}
      PCSComObj.Redraw;
      ShowStringInMemo('Circle Created');
    end;
end;

{creates a new line}
procedure TMainForm.ButtonCreateNewLineClick(Sender: TObject);
var
  ADocument : IPCsDocument;
  APage     : IPCsPage;
  ALine     : IPCsLine;
  Point3D   : TPCsPoint3d;
  i         : integer;
  str       : string;
begin
  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;

      {active project}
      ADocument := PCSComObj.ActiveDocument;
      {active page}
      APage := ADocument.ActivePage;
      {create line}
      ALine := APage.Lines.Add;

      {adds points to that line}
      for i := 0 to ListBoxLines.Count - 1 do
        begin
          str := ListBoxLines.Items[i];
          {assigning valus to TPCsPoint3D}
          Point3D.X := CutOut('X=',str,' ');
          Point3D.y := CutOut('Y=',str,' ');
          Point3D.Z := CutOut('Z=',str,' ');
          ALine.AddPoint(Point3D);
        end;
      {set linestyle}
      Aline.LineStyle := TranslateTrackPosToLineStyle(TrackBarLine.Position,false);
      {set line distance}
      Aline.MultilineDistance := Trunc(StrToFloat(ComboboxLineDistance.Text) * 1000);
      {set color}
      Aline.Color := integer(PenIndexToColor(ComboboxLineColor.ItemIndex));
      {set line width}
      Aline.PenWidth := trunc(StrToFloat(ComboboxLineWidth.Text) * 1000);
      {set is curved}
      Aline.IsCurvedLine := CheckBoxCurvedLine.Checked;
      ShowStringInMemo('A line has been added');
      {repaint diagram}
      PCSComObj.Redraw;
    end;
end;

{creates a new page}
procedure TMainForm.ButtonCreateNewPageClick(Sender: TObject);
var
  APage : IPCsPage;
begin
  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;

      {create page}
      Apage := PCSComObj.ActiveDocument.Drawing.Pages.Add(pf_Normal,'');
      {set pagefunction}
//      APage.PageFunction := psPageFunction(ComboboxPageFunction.Items[ComboboxPageFunction.ItemIndex]);
      APage.PageFunction := ComboboxPageFunction.ItemIndex;
      {set title}
      APage.Title := EditPageName.Text;
      {set pagetype}
      APage.PageType := psPageType(ComboboxPageType.Items[ComboboxPageType.ItemIndex]);
      if (ComboboxPageTabName.ItemIndex = 1) then
        {set specific pagetab-name}
        Apage.PageNumber := EditTabName.Text
      else
        {make pagename, the current pagenumber}
        Apage.PageNumber := IntToStr(PCSComObj.ActiveDocument.Drawing.Pages.Count);
      {repaint diagram}
      PCSComObj.Redraw;
      ShowStringInMemo('A new page, with title = ' + EditPageName.Text + ' has been created');
    end;
end;

{creates a new symbol}
procedure TMainForm.ButtonCreateSymbolClick(Sender: TObject);
var
  ADocument                     : IPCsDocument;
  APage                         : IPCsPage;
  ASymbol                       : IPCsSymbol;
  APt                           : TPCsPoint3D;
  ALibType                      : IPCsLibType;
  FilenameStr                   : string;
  ALine                         : IPCsLine;
  Xorg, YOrg, ZOrg, XMove,YMove : integer;
begin
  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;
      {active project}
      ADocument := PCSComObj.ActiveDocument;
      {active page}
      APage := ADocument.ActivePage;
      {assigning valus to TPCsPoint3D}
      APt.X := StrToInt(EditNewSymbolX.Text);
      APt.Y := StrToInt(EditNewSymbolY.Text);
      APt.Z := StrToInt(EditNewSymbolZ.Text);

      if (ComboboxSymbolFilename.ItemIndex > 0) then
        {every symbol is valid here, as long as its a file og type .SYM}
        ASymbol := Apage.PlaceSymbol(Apt,ComboboxSymbolFilename.Items[ComboboxSymbolFilename.ItemIndex])
      else
        begin
          {if no symbol is selected, then draw a symbol yourself...
          here a square is drawn.}
          {the format for Automation if no file specifies the dimension of symbol}
          FilenameStr := '#x' + EditSymbolWidth.Text + 'mmy' + EditSymbolHeight.Text + 'mm';
          {but just type another name if you want...}
          FilenameStr := 'SymbolFileName';
          {create Libtype}
          ALibType := Adocument.Drawing.LibTypes.Add(FilenameStr);
          {create line}
          ALine := ALibType.Lines.Add();
          XOrg := StrToInt(EditNewSymbolX.Text);
          YOrg := StrToInt(EditNewSymbolY.Text);
          ZOrg := StrToInt(EditNewSymbolZ.Text);
          XMove := StrToInt(EditSymbolwidth.Text);
          YMove := StrToInt(EditSymbolHeight.Text);
          {adding point, which make a square}
          ALine.AddPointXYZ(Xorg,YOrg,ZOrg);
          ALine.AddPointXYZ(Xorg+XMove,YOrg,Zorg);
          ALine.AddPointXYZ(Xorg+XMove,YOrg+YMove,Zorg);
          ALine.AddPointXYZ(Xorg,YOrg+YMove,Zorg);
          ALine.AddPointXYZ(Xorg,YOrg,ZOrg);
          {create a symbol, from LibType}
          ASymbol := Apage.Symbols.Add(ALibType);
          {setting position}
          ASymbol.Position := Apt;
        end;
      {set name}
      ASymbol.SymbolTexts.NameText.Value := EditSymbolNameText.Text;
      {set articletext}
      ASymbol.SymbolTexts.ArticleText.Value := EditSymbolArticleText.Text;
      {set typetext}
      ASymbol.SymbolTexts.TypeText.Value := EditSymbolTypeText.Text;
      {set functiontext}
      ASymbol.SymbolTexts.FunctionText.Value := EditSymbolFunctionText.Text;
      {repaint diagram}
      PCSComObj.Redraw;
      ShowStringInMemo('Symbol: "' + ASymbol.SymbolTexts.NameText.Value + '" Created');
    end;
end;

{creates a new text}
procedure TMainForm.ButtonCreateTextClick(Sender: TObject);
var
  AText   : IPCsText;
  Point3D : TPCsPoint3d;
begin
  if (trim(EditTextValue.Text) <> '') then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;

      {assigning valus to TPCsPoint3D}
      Point3D.X := StrToInt(EditTextPosX.Text);
      Point3D.Y := StrToInt(EditTextPosY.Text);
      Point3D.Z := StrToInt(EditTextPosZ.Text);
      {create text}
      AText := PCSComObj.ActiveDocument.ActivePage.Texts.Add(Point3D,EditTextValue.Text);
      {set color}
      AText.Color := integer(PenIndexToColor(ComboboxTextColor.ItemIndex));
      {set rotation}
      AText.Rotation := StrToInt(ComboboxTextAngle.Text) * 10;
      {set font + font options}
      AText.Font.Name := ComboboxTextFont.Items[ComboboxTextFont.ItemIndex];
      AText.Font.Bold := CheckboxTextBold.Checked;
      AText.Font.Italic := CheckboxTextItalic.Checked;
      FormatSettings.DecimalSeparator := ',';
      {set height}
      AText.Font.Height := Trunc(StrToFloat(ComboboxTextHeight.Items[ComboboxTextHeight.ItemIndex])*10);
      ShowMessage(IntToStr(AText.Font.Height));
      {width should not be set, it scale should fit}
      if (ComboboxTextWidth.Items[ComboboxTextWidth.ItemIndex] <> 'AUTO') then
        AText.Font.Width := Trunc(StrToFloat(ComboboxTextWidth.Items[ComboboxTextWidth.ItemIndex])*10);
      FormatSettings.DecimalSeparator := '.';
      ShowStringInMemo('A text has been added');
      {repaint diagram}
      PCSComObj.Redraw;
    end
  else
    ShowStringInMemo('No text entered');
end;

{shows circles on a pages}
procedure TMainForm.ButtonFindCirclesClick(Sender: TObject);
var
  FirStr, SecStr, ThiStr, FouStr, FivStr : string;
  i, PosX, PosY, PosZ                    : integer;
begin
  ButtonArcUpdate.Enabled := false;
  ButtonArcDelete.Enabled := false;
  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;

      ListBoxcircles.Items.BeginUpdate;
      ListBoxcircles.Clear;
      {runninf through all circles on a page}
      for i:= 0 to PCSComObj.ActiveDocument.ActivePage.Arcs.Count-1 do
        begin
          {get position}
          PCSComObj.ActiveDocument.ActivePage.Arcs.Item[i].GetPosition(PosX,PosY,PosZ);
          {get handle}
          FirStr := FixStringLength('Circle handle=' + IntToStr(PCSComObj.ActiveDocument.ActivePage.Arcs[i].Handle),25);
          SecStr := FixStringLength('Cords(X=' + IntToStr(PosX) + ',Y=' + IntToStr(PosY) + ',Z=' + IntToStr(PosZ) + ')',30);
          {get radius}
          ThiStr := FixStringLength('Radius=' + IntToStr(PCSComObj.ActiveDocument.ActivePage.Arcs.Item[i].Radius),16);
          {get elipsefactor}
          FouStr := FixStringLength('Elipse=' + FloatToStr(PCSComObj.ActiveDocument.ActivePage.Arcs.Item[i].EllipseFactor),11);
          {get rotation}
          FivStr := 'Rotation=' + FloatToStr(PCSComObj.ActiveDocument.ActivePage.Arcs.Item[i].Rotation);
          ListBoxcircles.Items.Add(FirStr+SecStr+ThiStr+FouStr+FivStr);
        end;
      ListBoxcircles.Items.EndUpdate;
      ShowStringInMemo('Circles found= ' + IntToStr(PCSComObj.ActiveDocument.ActivePage.Arcs.Count));
    end;
end;

{adds a random point}
procedure TMainForm.ButtonAddSymbolWithDatafieldClick(Sender: TObject);
var
  MyPoint : TPCsPoint3D;
  ALibtype : IPCsLibType;
  ADatafield : IPCsDatafield;
  ASymbol : IPCsSymbol;
  AStr : string;
begin
  {setting libtype}
  ALibtype := PCSComObj.ActiveDocument.Drawing.LibTypes.Add(EditSymbolFilenameDatafield.Text);
  MyPoint.X := 10000;
  MyPoint.Y := 10000;
  MyPoint.Z := 0;
  {adding a datafield}
  ADatafield := ALibtype.Datafields.Add(MyPoint,'');
  {setting value to that datafield}
  with ADatafield do
    begin
      Pretext := EditPretext.Text;
      if RadioButtonOtherPro.Checked then
        begin
          AStr := ComboBoxOtherProject.Items.Strings[ComboBoxOtherProject.ItemIndex];
          {as this is a none-predefined value, then the procedure SetDataNo is
          used together with FieldName}
          SetDataNo(StrToInt(ConvertDatafieldBetweenNumberAndName('NaToNo','FirstProjData')),PCSComObj.ActiveDocument.DataLink);
          FieldName := AStr;
        end;
      if RadioButtonOtherPage.Checked then
        begin
          AStr := ComboBoxOtherPage.Items.Strings[ComboBoxOtherPage.ItemIndex];
          SetDataNo(StrToInt(ConvertDatafieldBetweenNumberAndName('NaToNo','FirstPageData')),PCSComObj.ActiveDocument.DataLink);
          {as this is a none-predefined value, then the procedure SetDataNo is
          used together with FieldName}
          FieldName := AStr;
        end;
      if RadioButtonDFSymbols.Checked then
        begin
          AStr := ComboBoxDFSymbols.Items.Strings[ComboBoxDFSymbols.ItemIndex];
          SetDataNo(StrToInt(ConvertDatafieldBetweenNumberAndName('NaToNo','FirstSymbolData')),PCSComObj.ActiveDocument.DataLink);
          {as this is a none-predefined value, then the procedure SetDataNo is
          used together with FieldName}
          FieldName := AStr;
        end;
    end;
  {adding symbol to currentpage, using the libtype (symbol + datafield)}
  ASymbol := PCSComObj.ActiveDocument.ActivePage.Symbols.Add(ALibtype);
  {setting symbol name}
  ASymbol.FullName := EditSymbolNameDatafiled.Text;
  {setting position}
  ASymbol.SetPosition(StrToInt(EditDataFieldsX.Text), StrToInt(EditDataFieldsY.Text),0);
  {redraw}
  PCSComObj.Redraw;
end;

{creating a new datafield}
procedure TMainForm.ButtonCreateDatafieldClick(Sender: TObject);
const
  FieldAdded = 'Field added to ADrawing.';
var
  MyNewArray : Variant;
  ADrawing : IPCsDrawing;
  s : string;
  i : integer;
begin
  {getting the drawing}
  ADrawing := PCSComObj.ActiveDocument.Drawing;

  {setting integer according to data defs selected}
  s := ComboBoxDFCreateNew.Text;
  i := 0;
  if s = 'Project Data Defs' then
    i := 1;
  if s = 'Page Data Defs' then
    i := 2;
  if s = 'Symbol Data Defs' then
    i := 3;
  if i = 0 then
    exit;

  {the 3 arrays (ProjectDataDefs, PageDataDefs, SymbolDataDefs), can't be
  VarArrayReDim'et, so we're making a new array called MyNewArray}
  case i of
    1 : MyNewArray := ADrawing.ProjectDataDefs;
    2 : MyNewArray := ADrawing.PageDataDefs;
    3 : MyNewArray := ADrawing.SymbolDataDefs;
  end;

  {expanded that array by 1}
  VarArrayRedim(MyNewArray,VarArrayHighBound(MyNewArray,1)+1);
  {setting the new extra field to that specified by the user}
  MyNewArray[VarArrayHighBound(MyNewArray,1)] := EditDFFieldName.Text;

  {setting the corrext array, to MyNewArray}
  case i of
    1 : ADrawing.ProjectDataDefs := MyNewArray;
    2 : ADrawing.PageDataDefs := MyNewArray;
    3 : ADrawing.SymbolDataDefs := MyNewArray;
  end;

  {show information to user}
  case i of
    1 : ShowMessage(FieldAdded + 'ProjectDataDefs');
    2 : ShowMessage(FieldAdded + 'PageDataDefs');
    3 : ShowMessage(FieldAdded + 'SymbolDataDefs');
  end;

  {call a onshow-event, that make comboboxes, getting that new value in list}
  self.TabSheet6.OnShow(sender);
end;

procedure TMainForm.SettingSelectedLayer(ToVisible: boolean);
var
  str : string;
  LastPosOfEqual, SelectedIndex : integer;
begin
  // if item selected
  if ListBoxLayers.ItemIndex > -1 then
    begin
      str := ListBoxLayers.Items.Strings[ListBoxLayers.ItemIndex];
      LastPosOfEqual := LastPos('=',str) + 1;
      SelectedIndex := strToInt(trim(copy(Str,LastPosOfEqual,Length(str))));
      // selectedIndex is now value in the selected row in the listbox
      PCSComObj.ActiveDocument.Drawing.Layers[SelectedIndex].Visible := ToVisible;
      PCSComObj.Redraw;
      // updating LayerListbox
      Self.ButtonFindLayers.Click();
    end;
end;

procedure TMainForm.ButtonSetLayerVisibleClick(Sender: TObject);
begin
  SettingSelectedLayer(true {set visible});
end;

procedure TMainForm.ButtonSetLayerInvisibleClick(Sender: TObject);
begin
  SettingSelectedLayer(false {set invisible});
end;

procedure TMainForm.ButtonFindSymbolsClick(Sender: TObject);
begin
  FindAndHightlightSymbols(False);
end;

procedure TMainForm.ButtonSetCompGrpAndPosClick(Sender: TObject);
var
  ThisHandle : cardinal;
  ASymbol    : IPCsSymbol;
  NewValue : integer;
  RedrawIt : Boolean;
begin
  {get handle to symbol}
  ThisHandle := GetSymbolHandleFromListBox(ListBoxSymbols);
  {assign symbol, from handle}
  ASymbol := GetSymbolFromSymbolHandle(ThisHandle);
  {change the name of the symbol}
  RedrawIt := false;
  if TryStrToInt(EditCompGrpNr.Text, NewValue) then
    begin
      ASymbol.ComponentGroupNo := NewValue;
      {repaint diagram}
      RedrawIt := True;
      ShowStringInMemo('Symbol handle= ' + IntToStr(ThisHandle) + ' has CompGrpNo changed to ' + IntToStr(ASymbol.ComponentGroupNo));
    end
  else
    ShowStringInMemo('Not a valid CompGrpNo. It must be a number...');
  if TryStrToInt(EditSymbolPosNo.Text, NewValue) then
    begin
      ASymbol.ComponentPosNo := NewValue;
      RedrawIt := True;
      ShowStringInMemo('Symbol handle= ' + IntToStr(ThisHandle) + ' has CompPosNo changed to ' + IntToStr(ASymbol.ComponentPosNo));
    end
  else
    ShowStringInMemo('Not a valid CompPosNo. It must be a number...');

  if RedrawIt then
    PCSComObj.Redraw;
end;

procedure TMainForm.ButtonAddRandomPointClick(Sender: TObject);
begin
  Randomize;
  {the specific numbers are used, to keep drawing inside diagram area}
  ListBoxLines.Items.Add('X=' + IntToStr(Random(387000)+25000) + ' Y=' + IntToStr(Random(275000)+11000) + ' Z=' + EditLineZ.Text);
  ButtonCreateNewLine.Enabled := ListBoxLines.Count > 1;
end;

{find texts on a page}
procedure TMainForm.ButtonFindTextsClick(Sender: TObject);
var
  AText                   : IPCsTexts;
  i                       : integer;
  FirStr, SecStr, ThiStr  : string;
begin
  ListBoxFoundTexts.Items.BeginUpdate;
  ListBoxFoundTexts.Clear;
  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      {make sure the is a document open}
      if NoActiveDocument then
        CreateDocument;

      {assign text, from activepage}
      AText := PCSComObj.ActiveDocument.ActivePage.Texts;
      {running through texts on page}
      for i := 0 to AText.Count-1 do
        begin
          FirStr := FixStringLength('Text #' + IntToStr(i),15);
          {get textvalue}
          SecStr := FixStringLength('Value=' + AText.Item[i].Value,55);
          {get handle}
          ThiStr := 'Handle=' + IntToStr(AText.Item[i].Handle);
          ListboxFoundTexts.Items.Add(FirStr + SecStr + ThiStr);
        end;
      ButtonUpdateText.enabled := ListboxFoundTexts.ItemIndex > -1;
      ButtonDeleteText.enabled := ListboxFoundTexts.ItemIndex > -1;
      ShowStringInMemo('Number of texts found = ' + IntToStr(AText.Count));
    end;
  ListBoxFoundTexts.Items.EndUpdate;
end;

{add a point to listbox. (not to the actual line... this happens later)}
procedure TMainForm.ButtonAddThisPointClick(Sender: TObject);
begin
  if ((EditLineX.Text <> '') and (EditLineY.Text <> '') and (EditLineZ.Text <> '')) then
    begin
      ListBoxLines.Items.Add('X=' + EditLineX.Text + ' Y=' + EditLineY.Text + ' Z=' + EditLineZ.Text);
      EditLineX.Text := '0';
      EditLineY.Text := '0';
      EditLineZ.Text := '0';
    end;
  ButtonCreateNewLine.Enabled := ListBoxLines.Count > 1;
end;

{deletes a circle}
procedure TMainForm.ButtonArcDeleteClick(Sender: TObject);
var
  AArc : IPCsArc;
begin
  {assign circle, from handle}
  AArc := GetArcFromHandle(GetArcHandleFromListbox(listBoxCircles));
  ShowStringInMemo('Circles with handle= ' + IntToStr(GetArcHandleFromListbox(listBoxCircles)) + ' deleted');
  {delete circle}
  AArc.Delete;
  {repaint diagram}
  PCSComObj.Redraw;
end;

{update circle on a page}
procedure TMainForm.ButtonArcUpdateClick(Sender: TObject);
var
  AArc      : IPCsArc;
begin
  {assign circle, from handle}
  AArc := GetArcFromHandle(GetArcHandleFromListbox(listBoxCircles));
  {set position}
  AArc.SetPosition(StrToInt(EditArcX.Text),StrToInt(EditArcY.Text),StrToInt(EditArcZ.Text));
  {set radius}
  AArc.Radius := StrToInt(EditArcRadius.Text);
  {set ellipsefactor}
  AArc.EllipseFactor := StrToFloat(EditArcElipse.Text);
  {set rotation}
  AArc.Rotation := StrToInt(EditArcRotation.Text);
  {set stantangle}
  AArc.StartAngle := StrToInt(EditArcStartAngle.Text);
  {set endangle}
  AArc.EndAngle := StrToInt(EditArcEndAngle.Text);
  {set isfilled on/off}
  AArc.IsFilled := CheckBoxArcFilled.Checked;
  {set color}
  AArc.Color := integer(PenIndexToColor(ComboboxArcColor.ItemIndex));
  {set pen width}
  AArc.PenWidth := trunc(StrToFloat(EditArcPenWidth.Text));
  ShowStringInMemo('Circles with handle= ' + IntToStr(GetArcHandleFromListbox(listBoxCircles)) + ' modified');
  {repaint diagram}
  PCSComObj.Redraw;
end;

{change the name of a symbol}
procedure TMainForm.ButtonChangeSymbolNameClick(Sender: TObject);
var
  ThisHandle : cardinal;
  ASymbol    : IPCsSymbol;
begin
  {get handle to symbol}
  ThisHandle := GetSymbolHandleFromListBox(ListBoxSymbols);
  {assign symbol, from handle}
  ASymbol := GetSymbolFromSymbolHandle(ThisHandle);
  {change the name of the symbol}
  ASymbol.FullName := EditSymbolName.Text;
  {repaint diagram}
  PCSComObj.Redraw;
  ShowStringInMemo('Symbol handle= ' + IntToStr(ThisHandle) + ' has changed name to ' + EditSymbolName.Text);
end;

{connect to Automation}
procedure TMainForm.ButtonConnectToAutomationClick(Sender: TObject);
begin
  CheckConnection();
end;

{creates a new project}
procedure TMainForm.ButtonCreateNewProjectClick(Sender: TObject);
var
  ADocument : IPCsDocument;
  str       : string;
begin
  {checks connectionstatus. if not connected, then it tries to connect}
  if CheckConnection() then
    begin
      {filepath to project}
      str := EditPathToProject.Text+'\'+EditProjectFilename.Text;
      {create new project}
      ADocument := PCSComObj.Documents.Add('',str);
      {saves the project}
      ADocument.SaveAs(str);
      {set projectname}
      ADocument.Drawing.Title := EditProjectName.Text;
      PCSComObj.Redraw;
      ShowStringInMemo('A new project, with title = ' + EditProjectName.Text + ' has been created');
    end;
end;

{deletes a line}
procedure TMainForm.ButtonDeleteLineClick(Sender: TObject);
var
  ALines   : IPCsLines;
  MyHandle : cardinal;
begin
  {getting handle to line}
  MyHandle := CutOut('#',ListBoxModifyLines.Items[ListBoxModifyLines.ItemIndex],' ');
  {assign line, from activepages}
  ALines := PCSComObj.ActiveDocument.ActivePage.Lines;
  {delete circle, from handle}
  Alines.Item[MyHandle].Delete;
  ShowStringInMemo('Line # ' + IntToStr(MyHandle) + ' has been deleted');
  {repaint diagram}
  PCSComObj.Redraw;
  {update Listbox with lines}
  ButtonFindLinesClick(Sender);
  {disable add point, as no line is selected}
  ButtonModifyLineAddPoint.Enabled := false;
end;

{deletes a symbol}
procedure TMainForm.ButtonDeleteSymbolClick(Sender: TObject);
var
  ThisHandle : cardinal;
  Symbol     : IPCsSymbol;
  RemIndex   : integer;
begin
  {getting handle to symbol}
  ThisHandle := GetSymbolHandleFromListBox(ListBoxSymbols);
  {assign symbol, from handle}
  Symbol := GetSymbolFromSymbolHandle(ThisHandle);
  {delete the symbol}
  Symbol.Delete;
  ShowStringInMemo('Symbol handle= ' + IntToStr(ThisHandle) + ' has been deleted');
  RemIndex := ListBoxSymbols.ItemIndex;
  ButtonFindSymbolsOnPage.Click;
  if (ListBoxSymbols.Count > RemIndex) then
    begin
      ListBoxSymbols.ItemIndex := RemIndex;
      ListBoxSymbolsClick(sender);
    end;
  {repaint diagram}
  PCSComObj.Redraw;
end;

{deletes a text}
procedure TMainForm.ButtonDeleteTextClick(Sender: TObject);
var
  AText : IPCsText;
begin
  if (ListBoxFoundTexts.ItemIndex > -1) then
    begin
      {assign text, from handle}
      AText := GetTextFromHandle(CutOut('Handle=',ListBoxFoundTexts.Items[ListBoxFoundTexts.ItemIndex],' '));
      {delete the text}
      AText.Delete;
      {repaint diagram}
      PCSComObj.Redraw;
      ButtonFindTextsClick(Sender);
    end;
end;

{deletes a point from the listbox}
procedure TMainForm.ButtonDeleteThisPointClick(Sender: TObject);
begin
  {deletes a point from the listbox. (not from an actual line)}
  if ((ListBoxLines.Items.Count > 0) and (ListBoxlines.ItemIndex > -1)) then
    ListBoxLines.Items.Delete(ListBoxLines.ItemIndex);
  ButtonCreateNewLine.Enabled := ListBoxLines.Count > 1;
end;

procedure TMainForm.ButtonLoadSaveProBrowseClick(Sender: TObject);
begin
  OpenDialog.Filter := '*.pro';
  OpenDialog.FileName := EditLoadSaveFile.Text;
  if OpenDialog.Execute then
    EditLoadSaveFile.Text := OpenDialog.FileName;
end;

//procedure TMainForm.Test();
//var
//  {typical used types...}
//  AApllication : IPCsApplication;
//  ADocument    : IPCsDocument;
//  ADocuments   : IPCsDocuments;
//  ADrawing     : IPCsDrawing;
//  APage        : IPCsPage;
//  APages       : IPCsPages;
//  ASymbol      : IPCsSymbol;
//  ASymbols     : IPCsSymbols;
//  ALine        : IPCsLine;
//  ALines       : IPCsLines;
//  AText        : IPCsText;
//  ATexts       : IPCsTexts;
//  AArc         : IPCsArc;
//  AArcs        : IPCsArcs;
//  ALibType     : IPCsLibType;
//  Point3D      : TPCsPoint3D;
//  LibTypeConn  : IPCsLibTypeConnections;
//    AList  : IPCsList;
//    ALists : IPCsLists;
//begin
//
//end;

end.

