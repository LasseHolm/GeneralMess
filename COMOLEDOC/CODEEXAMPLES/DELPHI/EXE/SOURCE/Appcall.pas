unit appcall;

interface

Uses
  Windows, ShellApi;

TYPE
  PHWnd = ^HWnd;
FUNCTION  EnumFuncApp(Wnd: HWnd; TargetWindow: PHWnd): BOOLEAN;STDCALL;
PROCEDURE GotoStartetApp(ATargetProcessId : DWORD);
FUNCTION  WinExecAndWait(CONST Path       : STRING;
                               Visibility : WORD;
                         VAR   AppID      : DWORD) : BOOLEAN;
FUNCTION CreateAProcess(CONST Path :STRING; WaitReady: BOOLEAN): DWORD;
PROCEDURE CreateAProcessAndWait(CONST APath: STRING);
FUNCTION GetAppWindow(ATargetProcessId : DWORD) : HWnd;

FUNCTION FindExecutableName(aprogname:String): STRING;

CONST
  hApp   : THANDLE = 0;
  OleApp : THandle = 0;

CONST
  AppCancel : BOOLEAN = TRUE;

implementation

uses
  SysUtils, Forms, DIALOGS;

VAR
  TargetProcessId : DWORD;

FUNCTION  WinExecAndWait(CONST Path       : STRING;
                               Visibility : WORD;
                         VAR   AppID      : DWORD) : BOOLEAN;
VAR
  pStrTemp   : PCHAR;
  AStartupInfo : TStartupInfo;
  AProcessInformation : TProcessInformation;
BEGIN
  pStrTemp := StrAlloc(Length(Path)+1);
  StrPCopy(pStrTemp, Path);
  GetStartupInfo(AStartupInfo);
  With AStartupInfo DO BEGIN
    cb          := SizeOf(AStartupInfo);
    lpReserved  := NIL;
    lpDesktop   := NIL;
    lpTitle     := NIL;
    dwFlags     := 0;{STARTF_USESHOWWINDOW;}
    wsHOWWINDOW := sw_shownormal;
  END;

  AProcessInformation.hProcess := 0;
  AppId := 0;
  CreateProcess(NIL,	    // pointer to name of executable module
                          pStrTemp,   // pointer to command line string
                          NIL,	// pointer to process security attributes
                          NIL,	// pointer to thread security attributes
                          FALSE,	// handle inheritance flag
                          0,	// creation flags
                          NIL,	// pointer to new environment block
                          NIL,	// pointer to current directory name
                          AStartUpInfo,	// pointer to STARTUPINFO
                          AProcessInformation 	// pointer to PROCESS_INFORMATION
                          );
  WaitForInputIdle(GetCurrentProcess, INFINITE);
  AppId := AProcessInformation.dwProcessId;
  IF AProcessInformation.hProcess <> 0 THEN BEGIN
    AppCancel := FALSE;
    REPEAT
      Application.ProcessMessages;
    UNTIL (WAIT_TIMEOUT <> WaitForSingleObject(AProcessInformation.hProcess, 1)) OR
          (Application.Terminated) OR AppCancel;
    CloseHandle(AProcessInformation.hProcess);
    CloseHandle(AProcessInformation.hThread);
    AProcessInformation.hProcess := 0;
    Result := TRUE;
    AppCancel := TRUE;
  END
  ELSE BEGIN
    Result := FALSE;
    ShowMessage('Error executing : ' + Path )
  END;
  StrDispose(pStrTemp);
END;


FUNCTION EnumFuncApp(Wnd: HWnd; TargetWindow: PHWnd): BOOLEAN;
{ Works only in DELPHI 3 }
VAR
  AProcessId : DWORD;
  AThreadId  : DWORD;
BEGIN                                                        { EnumFuncApp }
  RESULT := TRUE;
  IF (GetParent(Wnd) = 0) AND (IsWindowVisible(Wnd)) THEN BEGIN
    AThreadId := GetWindowThreadProcessId(Wnd, @AProcessId);
    IF TargetProcessId <> 0 THEN BEGIN
      { 32 - Bit Program }
      IF AProcessId = TargetProcessId THEN BEGIN
        TargetWindow^ := Wnd;
        RESULT := FALSE;
      END;
    END
    ELSE BEGIN
      { 16 - Bit Program }
    END;
  END;
END;

FUNCTION GetAppWindow(ATargetProcessId : DWORD) : HWnd;
VAR
  InstWnd : HWnd;
BEGIN                                                          { GetAppWindow }
  TargetProcessId := ATargetProcessId;
  IF NOT EnumWindows(@EnumFuncApp,LONGINT(@InstWnd)) THEN
    RESULT := InstWnd
  ELSE
    RESULT := 0;
END;                                                           { GetAppWindow }

PROCEDURE GotoStartetApp(ATargetProcessId : DWORD);
VAR
  InstWnd : HWnd;
BEGIN                                                        { GotoStartetApp }
  InstWnd := GetAppWindow(ATargetProcessId);
  IF InstWnd <> 0 THEN
  BEGIN
    IF IsIconic(InstWnd) THEN
      ShowWindow(InstWnd,SW_RESTORE)
    ELSE
      BringWindowToTop(InstWnd)
  END; { if }
END;                                                         { GotoStartetApp }

PROCEDURE CreateAProcessAndWait(CONST APath: STRING);
VAR
  AppID : DWORD;
BEGIN                                                 { CreateAProcessAndWait }
  WinExecAndWait(APath,0,AppID);
END;                                                  { CreateAProcessAndWait }

FUNCTION CreateAProcess(CONST Path :STRING; WaitReady: BOOLEAN): DWORD;
VAR
  pStrTemp   : PCHAR;
  AStartupInfo : TStartupInfo;
  AProcessInformation : TProcessInformation;

BEGIN
  RESULT := 0;
  pStrTemp := StrAlloc(Length(Path)+1);
  StrPCopy(pStrTemp, Path);
  GetStartupInfo(AStartupInfo);
  With AStartupInfo DO BEGIN
    cb          := SizeOf(AStartupInfo);
    lpReserved  := NIL;
    lpDesktop   := NIL;
    lpTitle     := NIL;
    dwFlags     := 0;{STARTF_USESHOWWINDOW;}
    wsHOWWINDOW := sw_shownormal;
//    wsHOWWINDOW := sw_minimize;
  END;

  AProcessInformation.hProcess := 0;

  IF CreateProcess(NIL,	    // pointer to name of executable module
                pStrTemp,   // pointer to command line string
                NIL,	// pointer to process security attributes
                NIL,	// pointer to thread security attributes
                FALSE,	// handle inheritance flag
                0,	// creation flags
                NIL,	// pointer to new environment block
                NIL,	// pointer to current directory name
                AStartUpInfo,	// pointer to STARTUPINFO
                AProcessInformation) THEN 	// pointer to PROCESS_INFORMATION
   BEGIN
     IF WaitReady THEN
       WaitForInputIdle(AProcessInformation.hProcess{GetCurrentProcess}, 10000);
     RESULT := AProcessInformation.dwProcessId;
   END;
   CloseHandle(AProcessInformation.hProcess);
   CloseHandle(AProcessInformation.hThread);
   AProcessInformation.hProcess := 0;
   StrDispose(pStrTemp);
END;

FUNCTION FindExecutableName(aprogname:String): STRING;
VAR
  ExeFile : Array[0..MAX_PATH] of CHAR;
  szProgName : Array[0..MAX_PATH] of CHAR;
  szProgPath : Array[0..MAX_PATH] of CHAR;

BEGIN
  Result := aProgName;
  IF AnsiCompareText(ExtractFileExt(aProgName),'.EXE') <> 0 THEN
    BEGIN  //ProgName not a Exe file find assosiatet exe file
       StrPCopy(SzProgName,ExtractFileName(aProgName));
       StrPCopy(SzProgPath,ExtractFilePath(aProgName));
       IF FindExecutable(szProgname,szProgPath,ExeFile) > 32 THEN
         IF Pos(' ',aProgname) > 0 THEN
           Result := StrPas(ExeFile)+ ' "' + aProgName + '"'
         ELSE
           Result := StrPas(ExeFile)+ ' ' + aProgName ;
    END;
END;

end.
