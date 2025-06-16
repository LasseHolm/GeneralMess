unit PluginModule;
{ Plugin DLL to PCschematic }
{ The plugin DLLs and PCschematic shares this unit }
{ PCschematic calls LoadPluginModule which returns an IPlugInModuleobject object}
{ When the object is released the DLL is unloaded from memory  }

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  DPSCAD_TLB, Windows, ActiveX, Classes, ComObj, SysUtils;

const
  { The plugin DLL must export a TLoadPCschematicPlugin function with this name: }
  c_PCschematicPluginDLLProc = 'LoadPCschematicPlugin';

type
  TPluginProgressEvent = procedure (MicroPercentDone: integer; var Abort : boolean; const Msg: shortstring) of object;
  IPlugIn = interface
    ['{B3CB132D-83BB-4A6C-A7A3-7B3A85694ED5}']
    function Get_ParamString : shortstring;
    procedure Set_ParamString(const Value : shortstring);
    function Get_OnProgress: TPluginProgressEvent;
    procedure Set_OnProgress(const Value: TPluginProgressEvent);
    function Get_UIInteractive : boolean;
    procedure Set_UIInteractive(const Value : boolean);
    procedure Set_MainWindowHandle(AWnd : HWnd);
    procedure Set_PCsApplication(PCsApplication : IPCsApplication);
    procedure SetDirectories(ASysDir, AHomeDir : shortstring);
    function GetLastError : shortstring;
        
    { Run the plugin }
    function Execute : boolean;

    { Enables PCschematic to send custom data to the plugin: }
    property ParamString : shortstring read Get_ParamString write Set_ParamString;

    property MainWindowHandle : HWnd write Set_MainWindowHandle;
    { Enable PCschematic to set the Application.Handle of the plugin }
    { in order to make dialogboxes i the plugin to appear as part of PCschematic }

   { Enable PCschematic to send the Ole application object to the client: }
    property PCsApplication : IPCsApplication write Set_PCsApplication;

    { PCschematic (or another client) can set this to FALSE to disallow user interaction ("silent mode"): }
    property UIInteractive : boolean read Get_UIInteractive write Set_UIInteractive;

    { Enables PCschematic (or another client) to show the progress of a batch process:  }
    property OnProgress : TPluginProgressEvent read Get_OnProgress write Set_OnProgress;
  end;
  TPluginModule = class(TObject)
    { Wrapper around IPlugIn and the plugin DLL }
    { managing loading and releasing the DLL }
    private
      FPlugIn : IPlugin;
      FhLib   : cardinal;
    public
      constructor Create(APlugin : IPlugIn; hLib : cardinal);
      destructor Destroy; override;
      property Plugin : IPlugin read FPlugin;
  end;

  TLoadPCschematicPlugin = function : IPlugIn; stdcall;

function LoadPluginModule(ADLLFilename : ansistring) : TPlugInModule;


implementation

function LoadPluginModule(ADLLFilename : ansistring) : TPlugInModule;
{ Returns NIL if the DLL is not a PCschematic plugin }
{ Raises an exception on Win32 errors like file not found etc. }
var
  hLib       : cardinal;
  LoadModule : TLoadPCschematicPlugin;
  PlugIn     : IPlugIn;
  P          : pointer;
BEGIN
  Result := NIL;
  hLib   := SafeLoadLibrary(PChar(ADLLFilename));
  IF hLib = 0 THEN
    RaiseLastOSError;
  p := GetProcAddress(hLib, c_PCschematicPluginDLLProc);
  if p<>NIL then
  begin
    @LoadModule := p;
    PlugIn := LoadModule;
    RESULT := TPluginModule.Create(Plugin, hLib);
  end;
END;

{ TPluginModule }

constructor TPluginModule.Create(APlugin: IPlugin; hLib : cardinal);
begin
  inherited Create;
  FPlugIn := APlugin;
  FhLib := hLib;
end;

destructor TPluginModule.Destroy;
begin
  FPlugIn := NIL;
  FreeLibrary(FhLib);

  inherited;
end;

initialization

end.