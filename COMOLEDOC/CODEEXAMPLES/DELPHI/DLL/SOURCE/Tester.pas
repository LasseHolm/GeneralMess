unit Tester;

interface

uses
  Dpscad_TLB,Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls,
  Forms, Dialogs, StdCtrls, ExtCtrls, ComCtrls, StrUtils;

type
  TFormTester = class(TForm)
    ListBoxAllSymbols: TListBox;
    ListBoxAllLines: TListBox;
    LabelSymbols: TLabel;
    LabelLines: TLabel;
    ButtonGetSymbolsAndLines: TButton;
    procedure FormCreate(Sender: TObject);
    procedure ButtonGetSymbolsAndLinesClick(Sender: TObject);
  private
    PCSComObj     : IPCsApplication;
    procedure SetThisLength(var ThisStr : string; ThisLength : integer);
    { Private declarations }
  public
    constructor create(a : TComponent;PCSOLE: IPCsApplication);
    { Public declarations }
  end;

var
  FormTester  : TFormTester;

const
  ProgVersion = ' V1.00';
  ProgName    = 'DLL-test';

implementation

{$R *.dfm}
procedure TFormTester.SetThisLength(var ThisStr : string; ThisLength : integer);
begin
  while (Length(ThisStr) < ThisLength) do
    ThisStr := ThisStr + ' ';
end;

procedure TFormTester.ButtonGetSymbolsAndLinesClick(Sender: TObject);
var
  i, Sym, Lin, X, Y, Z, LinePoints  : integer;
  LinkSymbols                       : IPCsSymbols;
  LinkLines                         : IPCsLines;
  TempStr                           : string;
  RememberCursor                    : TCursor;
begin
  ButtonGetSymbolsAndLines.Enabled := false;
  RememberCursor := Screen.Cursor;
  Screen.Cursor := crHourGlass;
  Forms.Application.ProcessMessages;

  ListBoxAllSymbols.Clear;
  ListBoxAllLines.Clear;

  if (PCSComObj <> nil) then
    for i := 0 to PCSComObj.ActiveDocument.Drawing.Pages.Count -1 do
      begin
        LinkSymbols := PCSComObj.ActiveDocument.Drawing.Pages[i].Symbols;
        ListBoxAllSymbols.Items.BeginUpdate;
        for Sym := 0 to LinkSymbols.Count - 1 do
          begin
            TempStr := 'Page=' + IntToStr(i);
            SetThisLength(TempStr,10);
            TempStr := TempStr + 'Symbol=' + inttostr(Sym);
            SetThisLength(TempStr,22);
            TempStr := TempStr + LeftStr(LinkSymbols.Item[Sym].FullName,38);
            SetThisLength(TempStr,60);
            LinkSymbols.Item[Sym].GetPosition(x,y,z);
            TempStr := TempStr + ' Pos=(' + IntToStr(X) + ',' + IntToStr(Y) + ',' + IntToStr(Z) + ')';
            ListBoxAllSymbols.Items.Add(TempStr);
          end;
        ListBoxAllSymbols.Items.EndUpdate;
        LinkLines := PCSComObj.ActiveDocument.Drawing.Pages[i].Lines;
        ListBoxAllLines.Items.BeginUpdate;
        for Lin := 0 to LinkLines.Count - 1 do
          begin
            TempStr := 'Page=' + IntToStr(i);
            SetThisLength(TempStr,10);
            TempStr := TempStr + 'Line=' + inttostr(Lin);
            SetThisLength(TempStr,20);
            for Linepoints := 0 to LinkLines.Item[Lin].Points.Count - 1 do
              begin
                LinkLines.Item[Lin].Points.Item[LinePoints].GetPosition(X,Y,Z);
                TempStr := TempStr + ' Point' + inttostr(LinePoints) + '=(' + IntToStr(X) + ',' + IntToStr(Y) + ',' + IntToStr(Z) + ') ';
              end;
            ListBoxAllLines.Items.Add(TempStr);
          end;
        ListBoxAllLines.Items.EndUpdate;
      end;
  Screen.Cursor := RememberCursor;
  ButtonGetSymbolsAndLines.Enabled := true;
  Forms.Application.ProcessMessages;
end;

constructor TFormTester.create(a: TComponent; PCSOLE: IPCsApplication);
begin
  inherited create(a);
  PCSComObj := PCSOLE;
end;

procedure TFormTester.FormCreate(Sender: TObject);
begin
  // Make sure PCSObj is valid
  Self.Caption := ProgName + ProgVersion;
end;

end.
