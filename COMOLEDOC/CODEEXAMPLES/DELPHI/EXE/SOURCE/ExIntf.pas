unit ExIntf;
{ Sidst rev. 14/2-2005 af LH }

{ Interprocess-kommunikation mellem PCschematic og moduler                      }
{ Unitten bruges både af PCschematic og modulerne                               }
{ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -                   }

{ Hovedformål:                                                                  }
{    At give OLE-clienter mulighed for at connecte til                          }
{    en given server (CreateOLEObjectOnServer)                                  }
{ Serverens ansvar:                                                             }
{    - Skal registrere sig selv (fx. i Project-filen) med RegisterPCSInstance   }
{    - MainFormen skal sende WM_CreateOLEApp beskeder videre til                }
{      _WMCreateOLEAppHandler                                                    }
{    - I OLE Application-objektets Initialize henh. Destroy skal kaldes         }
{      FRegID:=RegisterAsRunning(Self)                                          }
{           henh.                                                               }
{       UnregisterAsRunning(FRegID).                                            }
{      (FRegID skal huskes af objektet)                                         }

{ Sekundært formål:                                                             }
{    TPCSSharedData giver adgang til data om alle processer der er registreret  }
{    med RegisterPCSInstance }

{ LOG for rettelser/ændringer: }
{   2/2 Version til BETA-test }
{   4/2 RemoveExistingInstances kaldes i TPCSSharedData.AddInstance }
{       således at hvis serverfilnavnet allerede er registeret slettes først }
{       (Gør at servernedbrud ikke betyder at man skal genstarte for at  }
{       ole-klienter virker ordenligt) }
{   9/7 Klienter kunne ikke køre på en netinstallation hvor serverens katalog var mappet }
{       Nu bruges UNC navne }
{  10/9 Kun klienter: Der er indsat gentagne forsøg på at connecte til serveren }
{       med timeout, således at hvis serveren ikke er klar forsøges igen        }
{  10/9 Kun klienter: Tilføjet funktionen IsPCSServerRunning, som returnerer    }
{       TRUE hvis den angivne server kører og har registreret sig selv          }
interface

uses
  Windows, Messages, ComObj, DPSCAD_TLB, Dialogs, SysUtils;

const
  WM_CreateOLEApp = WM_USER + 1975;
  MAX_PATHCHARS   = 255;

type
  EPCSComClient = class(Exception);

  IPCsAppObject = IDispatch;

  TProgramInstance = record
    ProcessID        : cardinal;              // Programmets processID
    MainWindowHandle : HWnd;                  // Programmets MainWindow
    Filename         : string[MAX_PATHCHARS]; // Filsti på programmet
    Reserved         : longword;              // (Fremtidssikring... Gamle klienter skal stadig kunne bruge gamle servere)
  end;

  TPCSSharedData = class
    private
      class function GetInstanceCount : integer;
      class function GetInstance(AIndex : integer) : TProgramInstance;
      class procedure AddInstance(AProgramInstance : TProgramInstance);
      class procedure RemoveExistingInstances(AServerFilepath : string);
      class procedure RemoveInstance(AProcessID : cardinal);
    public
      property InstanceCount : integer read GetInstanceCount;
      property Instances[AIndex : integer] : TProgramInstance read GetInstance;
      class function FindProgram(AFilepath : string) : TProgramInstance;
  end;

  { Kopiereret fra _TLB.pas filer som er genereret ud fra .TLB filer }
  IELautomationApplication = interface(IApp)
    ['{6CF6C0C6-D8FB-11D5-9030-0020AF16D6BF}']
  end;
  IElectronicsApplication = interface(IApp)
    ['{6CF6C0BF-D8FB-11D5-9030-0020AF16D6BF}']
  end;



{ Server procedurer:  }
procedure _WMCreateOLEAppHandler(var Msg : TMessage; AFactory, AOldFactory : TAutoObjectFactory);
function RegisterAsRunning(AObj : TAutoObject) : Longword;
procedure UnRegisterAsRunning(var ARegisterID : integer);
procedure RegisterPCSInstance;
procedure RevokeAllRunningAutomationObjs;

{ Client procedurer:  }
function CreateOLEObjectOnServer(AServerFilename : string; AInterface : TGuid) : IPCsApplication;
function GetRunningObject(APath : string; AInterface : TGUID) : IPCsApplication;
function PCSSharedData : TPCSSharedData;
function ConnectToPCSServer : IPCsApplication;
function IsPCSServerRunning(AServerFilename : string) : boolean;
procedure KillAllServers;

implementation

uses
  Classes, Forms, ActiveX, Inifiles, Controls;

type
  TSharedData = record
    Count : integer;
    ProgramInstances : array[0..14] of TProgramInstance;
  end;

var
  hMapObject       : THandle = 0;
  SharedData       : ^TSharedData;
  PCSSharedDataObj : TPCSSharedData = NIL;
  ObjList          : TList = NIL;

procedure RevokeAllRunningAutomationObjs;
{ Fjern alle registrerede (RegisterAsRunning) objekter }
{ fra Running Object Table  }
var
//  i : integer;
//  F : TGUID;
  RegID : integer;
begin
  if ObjList <> NIL then
  begin
    while (ObjList<>NIL) and (ObjList.Count > 0) do
    begin
      RegID := integer( ObjList[0] );
      UnregisterAsRunning(RegID);
    end;
    if ObjList<>NIL then
      ObjList.Free;
    ObjList := NIL;
  end;
end;

function RegisterAsRunning(AObj : TAutoObject) : Longword;
// Registrerer et COM objekt i Running Object Table
// under serverens filnavn
// Returnerer en ID som skal anvendes ved af-registreringen
var
  ROT : IRunningObjectTable;
  Reg : longint;
  Moniker : IMoniker;
  Filepath : PWideChar;
//  F : TGUID;
  ObjI : IUnknown;
begin
  Filepath := StringToOleStr( ExpandUNCFileName(ParamStr(0)) );
  OleCheck( CreateFileMoniker(Filepath, Moniker) );
  OleCheck( GetRunningObjectTable(0, ROT) );
  ObjI := AObj as IUnknown;
  OleCheck( ROT.Register(0, ObjI, Moniker, Reg) );
  ObjI := NIL;

  if ObjList=NIL then
    ObjList := TList.Create;

  ObjList.Add( Pointer(Reg) );

  RESULT := Reg;
end;

procedure UnRegisterAsRunning(var ARegisterID : integer);
var
  ROT : IRunningObjectTable;
  i : integer;
//  Moniker : IMoniker;
//  Filepath : PWideChar;
  Found : boolean;
begin
  try
    Found := FALSE;
    for i:=ObjList.Count-1 downto 0 do
    begin
      if integer(ObjList[i])=ARegisterID then
      begin
        ObjList.Delete(i);
        if ObjList.Count = 0 then
        begin
          ObjList.Free;
          ObjList := NIL;
        end;
        Found := TRUE;
      end;
    end;

    if Found then
    begin
      if ARegisterID <> 0 then
      begin
        OleCheck( GetRunningObjectTable(0, ROT) );
        OleCheck( ROT.Revoke(ARegisterID) );

        ARegisterID := 0;
      end;
    end;

  except
    // Ignorer fejl
  end;
end;

function GetRunningObject(APath : string; AInterface : TGUID) : IPCsApplication;
// Returnerer et COM objekt som er registreret i Running Object Table
// under den angivne fil-sti
var
  ROT : IRunningObjectTable;
  EnumMoniker : IEnumMoniker;
  Moniker : IMoniker;
  BindCtx : IBindCtx;
  StrPtr : pwidechar;
  DisplayName : ansistring;
  Obj : IPCsAppObject;
begin
  RESULT := NIL;
  Screen.Cursor := crHourglass;
  OleCheck( GetRunningObjectTable(0, ROT) );
  OleCheck(ROT.EnumRunning(EnumMoniker));
  OleCheck(EnumMoniker.Reset);
  try
    while EnumMoniker.Next(1, Moniker,NIL) = S_OK do
    begin
      OleCheck( CreateBindCtx(0, BindCtx) );
      StrPtr := NIL;
      StrPtr := CoTaskMemAlloc(100);

      OleCheck( Moniker.GetDisplayName(BindCtx, NIL, StrPtr) );
      DisplayName := StrPtr;
      CoTaskMemFree(StrPtr);

      if Uppercase(DisplayName) = Uppercase(APath) then
      begin
        OleCheck(BindMoniker(Moniker,0,IPCsAppObject,Obj));
        if Supports(Obj, AInterface) then
        begin
          RESULT := Obj as IPCsApplication;
          Exit;
        end;
      end;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

function GetHomeDir : TFileName;
var
  i : Integer;
  HomeDir : string;
BEGIN
  HomeDir := ExtractFilepath(ParamStr(0));
  FOR i := 1 to ParamCount do
    If POS('/N',UPPERCASE(Paramstr(i))) = 1 THEN
      BEGIN
        Homedir := UpperCase(Copy(Paramstr(i),3,Length(Paramstr(i))));
        IF POS(':',Homedir) = 1 THEN
          HomeDir := Copy(Homedir,2,Length(Homedir));
        IF HomeDir = '' THEN
          GETDIR(0,Homedir);
        Break;
      END;
  IF HomeDir[Length(HomeDir)] <> '\' THEN
    HomeDir := HomeDir + '\';
  RESULT := HomeDir;
END;

function GetParamOption(Option : string) : string;
var
   OptionStr : string;
   i : integer;
begin
   OptionStr := '';
   for i:=1 to ParamCount do
   begin
      if ((Uppercase(Copy(ParamStr(i),1,Length(Option))) = Uppercase(Option))) or
         ((Uppercase(Copy(ParamStr(i),1,Length(Option)+1)) = Uppercase(Option)+':')) then
      begin
         OptionStr := Uppercase(Copy(ParamStr(i),Length(Option)+1,1000));
         if OptionStr[1] = ':' then
            Delete(OptionStr,1,1);
         Break;
      end;
   end;
   RESULT := OptionStr;
end;

function GetParamSwitch(Switch : string) : boolean;
var
   SwitchFound : boolean;
   i : integer;
begin
   SwitchFound := FALSE;
   for i:=1 to ParamCount do
   begin
      if Uppercase(ParamStr(i)) = Uppercase(Switch) then
      begin
         SwitchFound := TRUE;
         Break;
      end;
   end;
   RESULT := SwitchFound;
end;

function Test : boolean;
begin
  RESULT := GetParamSwitch('/TEST');
end;

function ConnectToPCSServerOld : IApp;
{ Bør ikke anvendes af nye klienter }
var
  Inifile     : TInifile;
  ProgID,
  MyPath,
  Path,
  SvrFilename : ansistring;
  Intf        : TGuid;
  Obj         : IUnknown;
begin
  Path := GetHomeDir;
  Inifile := TInifile.Create(Path + 'PCSCAD.INI');
  ProgID := Inifile.ReadString('SystemData','ProgramOLEServerType', '');
  if ProgID = '' then
    ProgID := Inifile.ReadString('SystemData','ProgramType','');
  SvrFilename := Inifile.ReadString('SystemData','OleServerFilename','');
  Inifile.Free;
  MyPath := GetParamOption('/SERVERPATH');
  if MyPath = '' then
    MyPath := Uppercase(ExtractFilePath(ParamStr(0)));

  MyPath := ExpandUNCFileName(MyPath);

  if SvrFilename = '' then
  begin
    if (ProgID = 'ELCAD') or (ProgID = '') then
    begin
      if FileExists(MyPath + 'DPSCAD.EXE') then
        SvrFilename := 'DPSCAD.EXE'
      else
        SvrFilename := 'PCSELCAD.EXE';
      Intf := IELautomationApplication;
    end
    else
    begin
      if ProgID = 'ECAD' then
      begin
        SvrFilename := 'PCSECAD.EXE';
        Intf := IElectronicsApplication;
      end;
    end;
  end;

  if SvrFilename = '' then
    raise EPCSComClient.Create('Wrong PCschematic COM server');

  Obj := CreateOLEObjectOnServer(MyPath + SvrFilename, Intf);
  if Obj <> NIL then
    RESULT := IApp(Obj)
  else
    RESULT := NIL;

  if RESULT = NIL then
    raise EPCSComClient.Create('Error connecting to PCschematic COM server: '+ MyPath + SvrFilename);
end;

function ConnectToPCSServer : IPCsApplication;
{ Connect til en PCs ole-server som befinder sig i samme katalog som klienten }
{ Hvilke server der skal startes afhænger af programtypen i PCSCAD.INI }
var
  Inifile     : TInifile;
  ProgID,
  MyPath,
  Path,
  SvrFilename : ansistring;
  Obj         : IPCsApplication;
begin
  Path := GetHomeDir;
  Inifile := TInifile.Create(Path + 'PCSCAD.INI');
  ProgID := Inifile.ReadString('SystemData','ProgramOLEServerType', '');
  if ProgID = '' then
    ProgID := Inifile.ReadString('SystemData','ProgramType','');
  SvrFilename := Inifile.ReadString('SystemData','OleServerFilename','');
  Inifile.Free;
  MyPath := GetParamOption('/SERVERPATH');
  if MyPath = '' then
    MyPath := Uppercase(ExtractFilePath(ParamStr(0)));

  MyPath := ExpandUNCFileName(MyPath);

  if SvrFilename = '' then
  begin
    if (ProgID = 'ELCAD') or (ProgID = '') then
    begin
      if FileExists(MyPath + 'DPSCAD.EXE') then
        SvrFilename := 'DPSCAD.EXE'
      else
        SvrFilename := 'PCSELCAD.EXE';
    end
    else
    begin
      if ProgID = 'ECAD' then
      begin
        SvrFilename := 'PCSECAD.EXE';
      end;
    end;
  end;

  if SvrFilename = '' then
    raise EPCSComClient.Create('Wrong PCschematic COM server');

  Obj := CreateOLEObjectOnServer(MyPath + SvrFilename, IPCsApplication);
  if Obj <> NIL then
    RESULT := Obj
  else
    RESULT := NIL;

  if RESULT = NIL then
    raise EPCSComClient.Create('Error connecting to PCschematic COM server: '+ MyPath + SvrFilename);
end;

function IsPCSServerRunning(AServerFilename : string) : boolean;
var
  ProgramInstance : TProgramInstance;
begin
  ProgramInstance := PCSSharedData.FindProgram(AServerFilename);
  RESULT := (ProgramInstance.ProcessID <> 0);
end;

procedure RemoveThisServerFromROT;
var
  TempInst : IDispatch;
begin
  // Find evt. objekter som er registreret under dette serverfilnavn
  // og nedlæg dem:
  repeat
    try
      TempInst := GetRunningObject(ExpandUNCFileName(ParamStr(0)), IELautomationApplication);
    except
      { Ignorer fejl }
    end;
  until TempInst = NIL;
  repeat
    try
      TempInst := GetRunningObject(ExpandUNCFileName(ParamStr(0)), IElectronicsApplication);
    except
      { Ignorer fejl }
    end;
  until TempInst = NIL;
  repeat
    try
      TempInst := GetRunningObject(ExpandUNCFileName(ParamStr(0)), IPCsApplication);
    except
      { Ignorer fejl }
    end;
  until TempInst = NIL;

  TempInst := NIL;
end;

procedure KillAllServers;
{ Fjern alle registreringer af servere. Kan anvendes efter server-nedbrud, men }
{ kun til TEST ! }
var
  Instance : TProgramInstance;
begin
  CoInitialize(NIL);
  while PCSSharedData.GetInstanceCount > 0 do
  begin
    Instance := PCSSharedData.GetInstance(0);
    PCSSharedData.RemoveInstance(Instance.ProcessID);
  end;

  RemoveThisServerFromROT;

  RevokeAllRunningAutomationObjs;

  CoUnInitialize;
end;

procedure AllocSharedData;
var
  Init: boolean;
begin
  if hMapObject = 0 then
  begin
    // Create a named file mapping object.
    hMapObject := CreateFileMapping(
      THandle($FFFFFFFF),  // use paging file
      NIL,                 // no security attributes
      PAGE_READWRITE,      // read/write access
      0,                   // size: high 32-bits
      SizeOf(TSharedData), // size: low 32-bits
      'PCschematicOLESharedDataBlock'); // name of map object, THIS MUST BE UNIQUE!
    // The first process to attach initializes memory.
    Init := GetLastError <> ERROR_ALREADY_EXISTS;
    if hMapObject = 0 then exit;
    // Get a pointer to the file-mapped shared memory.
    SharedData := MapViewOfFile(
      hMapObject,     // object to map view of
      FILE_MAP_WRITE, // read/write access
      0,              // high offset:  map from
      0,              // low offset:   beginning
      0);             // default: map entire file
    if SharedData = NIL then exit;
    // Initialize memory if this is the first process.
    if Init then
    begin
      FillChar(SharedData^, SizeOf(TSharedData), 0);
    end;
  end;
end;

procedure FreeSharedData;
begin
  if hMapObject <> 0 then
    try
      try
        if not UnMapViewOfFile(SharedData) then
          MessageBox(0, 'Unable to free shared mem','error',0);
        if not CloseHandle(hMapObject) then
          MessageBox(0,'Unable to free shared mem handle','error',0);
      except
      end;
    finally
      SharedData := NIL;
      hMapObject := 0;
    end;
end;

procedure ArrayPullDown(AIndex : integer; AArray : array of TProgramInstance);
var
  i : integer;
begin
  for i:=AIndex to Length(AArray)-2 do
    AArray[i] := AArray[i+1];

  FillChar( AArray[ Length(AArray)-1 ], SizeOf(TProgramInstance), 0);
end;

function GetInstanceShortFilepath(AFullPath : string) : ansistring;
{ Da filnavnet i TProgramInstances er begrænset tol MAX_PATHCHARS karakterer }
{ sørger jeg for at det er den SIDSTE del af filnavnet der regnes med }
{ På den måde vil det være meget sjældent at man vil mærke til denne begrænsning }
var
  Filepath : ansistring;
begin
  Filepath := AFullPath;

  if Length(Filepath) > MAX_PATHCHARS then
    Filepath := Copy(Filepath, Length(Filepath)-MAX_PATHCHARS-1, MAX_PATHCHARS);

  RESULT := Filepath;
end;

procedure RegisterPCSInstance;
// Registrer data om denne proces i shared hukommelse, således at bl.a. CreateOLEObjectOnServer
// kan finde en given instance af en server som den kan connecte til
var
  NewInstance : TProgramInstance;
//  Buffer : array[0..254] of char;
begin
  if Application.MainForm = NIL then
    raise Exception.Create('MainForm not created yet. RegisterPCSInstance must be called after MainForm creation');

  NewInstance.MainWindowHandle := Application.MainForm.Handle;
  NewInstance.Filename := GetInstanceShortFilepath(ExpandUNCFileName(ParamStr(0)));
  NewInstance.ProcessID := GetCurrentProcessId;
  PCSSharedData.AddInstance( NewInstance );
end;

procedure CreateCOMObjectOnServerOld(AServerWindow : HWnd);
begin
  OleCheck(SendMessage(AServerWindow, WM_CreateOLEApp, 0,0));
end;

procedure CreateCOMObjectOnServer(AServerWindow : HWnd);
begin
  OleCheck(SendMessage(AServerWindow, WM_CreateOLEApp, 0,1));
end;

function CreateOLEObjectOnServerOld(AServerFilename : string; AInterface : TGUID) : IPCsAppObject;
{ Bør ikke anvendes af nye klienter }
var
  ProgramInstance   : TProgramInstance;
  Obj{, Temp}         : IPCsAppObject;
  TimeOut : integer;
  CursorBefore      : TCursor;
begin
  Obj := NIL;
  CursorBefore := Screen.Cursor;
  Screen.Cursor := crHourGlass;
  try
    TimeOut := 25; {5 sec}
    repeat
      ProgramInstance := PCSSharedData.FindProgram(AServerFilename);

      if ProgramInstance.ProcessID <> 0 then
      begin
        Obj := GetRunningObject(AServerFilename, AInterface);
        if Obj = NIL then
        begin
          CreateCOMObjectOnServer(ProgramInstance.MainWindowHandle);

          Obj := GetRunningObject(AServerFilename, AInterface);
        end;
      end
      else
        Obj := NIL;
      Dec(TimeOut);
      if Obj = NIL then
      begin
        Sleep(200);
      end;
    until (Obj <> NIL) or (TimeOut = 0);
  finally
    Screen.Cursor := CursorBefore;
  end;

  if ProgramInstance.ProcessID = 0 then
    raise EPCSComClient.Create('PCschematic COM server is not running'+#13+AServerFilename);

  RESULT := Obj;
end;

function CreateOLEObjectOnServer(AServerFilename : string; AInterface : TGUID) : IPCsApplication;
{ Opret et nyt OLE objekt på en kørende server med filnavnet AServerFilename }
{ (Hvis to servere har samme sti kan der naturligvis ikke skelnes)           }
var
  ProgramInstance   : TProgramInstance;
  Obj{, Temp}         : IPCsApplication;
  TimeOut{, TimeOut2} : integer;
  CursorBefore      : TCursor;
begin
  Obj := NIL;
  CursorBefore := Screen.Cursor;
  Screen.Cursor := crHourGlass;
  try
    TimeOut := 25; {5 sec}
    repeat
      ProgramInstance := PCSSharedData.FindProgram(AServerFilename);

      if ProgramInstance.ProcessID <> 0 then
      begin
        Obj := GetRunningObject(AServerFilename, AInterface) as IPCsApplication;
        if Obj = NIL then
        begin
          CreateCOMObjectOnServer(ProgramInstance.MainWindowHandle);

          Obj := GetRunningObject(AServerFilename, AInterface) as IPCsApplication;
        end;
      end
      else
        Obj := NIL;
      Dec(TimeOut);
      if Obj = NIL then
      begin
        Sleep(200);
      end;
    until (Obj <> NIL) or (TimeOut = 0);
  finally
    Screen.Cursor := CursorBefore;
  end;

  if ProgramInstance.ProcessID = 0 then
    raise EPCSComClient.Create('PCschematic COM server is not running'+#13+AServerFilename);

  RESULT := Obj;
end;

procedure _WMCreateOLEAppHandler(var Msg : TMessage; AFactory, AOldFactory : TAutoObjectFactory);
{ Handler til WM_CreateOLEApp message. Opretter et nyt objekt i serveren, således at  }
{ en klient har mulighed for at connecte til netop en bestemt server, ved at bede }
{ serveren om at oprette et objekt, som så bliver registreret med RegisterAsRunning }
var
  Obj     : IPCsAppObject;
//  Factory : TAutoObjectFactory;
begin
  Msg.Result := 0;
  if Msg.LParam = 0 then
  begin
    if (AOldFactory<>NIL) then
    begin
      Obj := AOldFactory.CreateComObject(NIL) as IPCsAppObject;
      Msg.Result := Integer(Pointer(Obj));
    end;
  end
  else if Msg.LParam = 1 then
  begin
    if (AFactory<>NIL) then
    begin
      Obj := AFactory.CreateComObject(NIL) as IPCsAppObject;
      Msg.Result := Integer(Pointer(Obj));
    end;
  end;
end;

(*** TPCSSharedData: ****************************************************************************************)

class procedure TPCSSharedData.AddInstance(AProgramInstance : TProgramInstance);
begin
  AllocSharedData;

  with SharedData^ do
  begin
    RemoveExistingInstances(AProgramInstance.Filename);

    Inc(Count);

    if Count > Length(ProgramInstances) then
    begin
      Count := Length(ProgramInstances);
      ArrayPullDown(0, ProgramInstances);
    end;

    ProgramInstances[Count-1] := AProgramInstance;
  end;
end;

class procedure TPCSSharedData.RemoveExistingInstances(AServerFilepath : string);
var
  i : integer;
begin
  // Fjerner alle eksisterende registreringer af den pågældende server
  // Hvis PCschematic ikke blev lukket korrekt (fx. nedbrud) skal alle
  // tidligere server-registreringer til denne server fjernes før den nye
  // tilføjes

  AllocSharedData;

  i:=0;
  with SharedData^ do
  begin
    while i < Count do
      if ProgramInstances[i].Filename = AServerFilepath then
      begin
        ArrayPullDown(i, ProgramInstances);
        Dec(Count);
      end
      else
        Inc(i);
  end;

  // Fjern fra ROT:
  RemoveThisServerFromROT;
end;

class procedure TPCSSharedData.RemoveInstance(AProcessID : cardinal);
var
  i : integer;
begin
  AllocSharedData;

  i:=0;
  with SharedData^ do
  begin
    while i < Count do
      if ProgramInstances[i].ProcessID = AProcessID then
      begin
        ArrayPullDown(i, ProgramInstances);
        Dec(Count);
      end
      else
        Inc(i);
  end;
end;

class function TPCSSharedData.FindProgram(AFilepath : string) : TProgramInstance;
var
  i : integer;
//  ModuleH : cardinal;
//  FilenameCh : array[0..MAX_PATH-1] of char;
//  Filename : string;
begin
  AllocSharedData;
  AFilepath := GetInstanceShortFilepath( Uppercase(AFilepath) );

  FillChar(RESULT, SizeOf(RESULT), 0);

  with SharedData^ do
  begin
    for i:=0 to Count-1 do
    begin
      if Uppercase(ExpandUNCFileName(ProgramInstances[i].Filename)) = AFilepath then
      begin
        RESULT := ProgramInstances[i];
        Exit;
      end;
    end;
  end;
end;

class function TPCSSharedData.GetInstanceCount : integer;
begin
  AllocSharedData;
  RESULT := SharedData^.Count;
end;

class function TPCSSharedData.GetInstance(AIndex : integer) : TProgramInstance;
begin
  AllocSharedData;
  if (AIndex < SharedData^.Count) and (AIndex >= 0) then
    RESULT := SharedData^.ProgramInstances[AIndex]
  else
    raise Exception.Create('Index out of range');
end;

function PCSSharedData : TPCSSharedData;
{ Funktion der giver adgang til metoderne i TPCSSharedData som indeholder }
{ sharede data}
begin
  if PCSSharedDataObj = NIL then
    PCSSharedDataObj := TPCSSharedData.Create;

  RESULT := PCSSharedDataObj;
end;

procedure FreePCSSharedData;
begin
  if PCSSharedDataObj<>NIL then
  begin
    PCSSharedDataObj.Free;
    PCSSharedDataObj := NIL;
  end;
end;


Initialization
  if FindCmdLineSwitch('killservers') then
  begin
    { Mulighed for klienter til at fjerne alle servere, så det ikke er nødvendigt }
    { at genstarte ved server-nedbrud.  }
    { Denne mulighed skal bortfalde når serveren bliver opdateret med denne unit }
    { fordi den da selv sørge for at fjerne "gamle" nedbrudte servere }
    KillAllServers;
    Halt;
  end;

Finalization
  try
    if not IsLibrary then
    begin
      // Hvis processen blev registreret i TPCSSharedData så fjern den nu:
      PCSSharedData.RemoveInstance(GetCurrentProcessId);

      // Frigiv evt. oprettede TPCSSharedData samt linket til den sharede hukommelse:
      FreePCSSharedData;
      FreeSharedData;
    end;
  except
    MessageBox(0, 'Error in ExIntf.finalization', NIL, MB_ICONERROR);
  end;
end.