unit TestPlugin;

interface

uses
  PluginModule,
  SysUtils,
  Windows,
  DPSCAD_TLB,
  Forms,
  Dialogs,
  Tester;

type
  TTestPlugin = class(TInterfacedObject, IPlugin)
    private
      FSilentMode     : boolean;
      FParamStr       : ansistring;
      FLastError      : shortstring;
      FPCsApplication : IPCsApplication;
    public
      constructor Create;
      destructor Destroy; override;
      function Get_ParamString : shortstring;
      procedure Set_ParamString(const Value : shortstring);
      function Get_OnProgress: TPluginProgressEvent;
      procedure Set_OnProgress(const Value: TPluginProgressEvent);
      function Get_UIInteractive : boolean;
      procedure Set_UIInteractive(const Value : boolean);
      function Execute : boolean;
      function GetLastError : shortstring;
      property ParamString : shortstring read Get_ParamString write Set_ParamString;
      procedure Set_MainWindowHandle(AWnd : HWnd);
      procedure Set_PCsApplication(PCsApplication : IPCsApplication);
      procedure SetDirectories(ASysDir,AHomeDir : shortstring);
      property UIInteractive : boolean read Get_UIInteractive write Set_UIInteractive;
      property OnProgress : TPluginProgressEvent read Get_OnProgress write Set_OnProgress;
  end;

function LoadPCschematicPlugin2 : IPlugIn; stdcall;

implementation

function LoadPCschematicPlugin2 : IPlugIn; stdcall;
begin
  RESULT := TTestPlugin.Create;
end;

{ TProjectGeneratorPlugin }

constructor TTestPlugin.Create;
begin
  inherited;
  FSilentMode := FALSE;
  FLastError := '';            
end;

destructor TTestPlugin.Destroy;
begin
  inherited;
end;

function TTestPlugin.Execute : boolean;
var
  tempForm : TFormTester;
begin
  try
    FLastError := '';
    tempForm := TFormTester.create(nil, FPCsApplication);
    TempForm.showModal;
    FreeAndNil(TempForm);
    RESULT := TRUE;
  except
    on x:exception do
    begin
      FLastError := x.Message;
      raise;
    end;
  end;
end;

function TTestPlugin.GetLastError : shortstring;
begin
  RESULT := FLastError;
end;

function TTestPlugin.Get_OnProgress: TPluginProgressEvent;
begin
  RESULT := NIL;
end;

procedure TTestPlugin.Set_OnProgress(
  const Value: TPluginProgressEvent);
begin
end;

function TTestPlugin.Get_ParamString: shortstring;
begin
  RESULT := FParamStr;
end;

procedure TTestPlugin.Set_ParamString(
  const Value: shortstring);
begin
  { Format for paramstring: }
  { MyParam1=value;MyParam2=value }
  { If the module is called from the PCschematic PCschematic tools menu with }
  { commandline parameters, the paramstring will be containing a parameter }
  { called 'commandlineparameters' with the parameters as the value }
  FParamStr := Value;
end;

procedure TTestPlugin.Set_UIInteractive(const Value: boolean);
begin
  FSilentMode := not Value;
end;

function TTestPlugin.Get_UIInteractive: boolean;
begin
  RESULT := not FSilentMode;
end;

procedure TTestPlugin.Set_PCsApplication(PCsApplication: IPCsApplication);
begin
  FPCsApplication := PCsApplication;
end;

procedure TTestPlugin.SetDirectories(ASysDir,AHomeDir : shortstring);
begin
  { ASysDir and AHomedir contains the system directory and the home }
  { directory for the PCschematic system }
end;

procedure TTestPlugin.Set_MainWindowHandle(AWnd : HWnd);
begin
  if AWnd <> 0 then
    Application.Handle := AWnd;
end;

var
  dllSavedApplicationHandle : HWnd = 0;
initialization
  dllSavedApplicationHandle := Application.Handle;
finalization
  Application.Handle := dllSavedApplicationHandle;
end.
