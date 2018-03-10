{
   code : https://pastebin.com/w4X4pHpC
    Author: ?
    License : MIT  ??
}

unit RunElevatedSupport;


{$WARN SYMBOL_PLATFORM OFF}
{$R+}

interface

uses
  Windows;

type
  TElevatedProc        = function(const AParameters: String): Cardinal;
  TProcessMessagesMeth = procedure of object;

var
  // Warning: this function will be executed in external process.
  // Do not use any global variables inside this routine!
  // Use only supplied AParameters.
  OnElevateProc: TElevatedProc;

// Call this routine after you have assigned OnElevateProc
procedure CheckForElevatedTask;

// Runs OnElevateProc under full administrator rights
function RunElevated(const AParameters: String; const AWnd: HWND = 0; const AProcessMessages: TProcessMessagesMeth = nil): Cardinal; overload;

function  IsAdministrator: Boolean;
function  IsAdministratorAccount: Boolean;
function  IsUACEnabled: Boolean;
function  IsElevated: Boolean;
procedure SetButtonElevated(const AButtonHandle: THandle);

function GetProcessIdentity(): String;

implementation

uses
  SysUtils, Registry, ShellAPI, ComObj;

const
  RunElevatedTaskSwitch = '0CC5C50CB7D643B68CB900BF000FFFD5'; // some unique value, just a GUID with removed '[', ']', and '-'

function CheckTokenMembership(TokenHandle: THANDLE; SidToCheck: Pointer; var IsMember: BOOL): BOOL; stdcall; external advapi32 name 'CheckTokenMembership';


//�˺���������,����,��RunAsAdmin����
function RunElevated(const AParameters: String; const AWnd: HWND = 0; const AProcessMessages: TProcessMessagesMeth = nil): Cardinal; overload;
var
  SEI: TShellExecuteInfo;
  Host: String;
  Args: String;
  av:boolean;
begin
  Assert(Assigned(OnElevateProc), 'OnElevateProc must be assigned before calling RunElevated');

  if IsElevated then
  begin
    if Assigned(OnElevateProc) then
      Result := OnElevateProc(AParameters)
    else
      Result := ERROR_PROC_NOT_FOUND;
    Exit;
  end;


  Host := ParamStr(0);
  Args := Format('/%s %s', [RunElevatedTaskSwitch, AParameters]);

  //FillChar(SEI, SizeOf(SEI), 0);

  //ZeroMemory(@sei, SizeOf(sei));
  //SEI.cbSize := SizeOf(SEI);
  //sei.cbSize := SizeOf(TShellExecuteInfo);
  //SEI.fMask := SEE_MASK_NOCLOSEPROCESS;
  //{$IFDEF UNICODE}
  //SEI.fMask := SEI.fMask or SEE_MASK_UNICODE;
  //{$ENDIF}
  //SEI.Wnd := AWnd;
  //SEI.lpVerb := 'runas';
  //SEI.lpFile := PChar(Host);
  //SEI.lpParameters := PChar(Args);
  //SEI.nShow := SW_NORMAL;
  //SEI.nShow := SW_SHOWNORMAL ;

    ZeroMemory(@sei, SizeOf(sei));
    sei.cbSize := SizeOf(TShellExecuteInfo);
    sei.Wnd := AWnd;
    sei.fMask := SEE_MASK_FLAG_DDEWAIT or SEE_MASK_FLAG_NO_UI;
    sei.lpVerb := PChar('runas');
    sei.lpFile := PChar(Host); // PAnsiChar;
    sei.lpParameters := PChar(Args); // PAnsiChar;
    sei.nShow := SW_SHOWNORMAL; //Integer;

    av:= ShellExecuteEx(@sei);
    if not av then   RaiseLastOSError;



  //if not ShellExecuteEx(@SEI) then RaiseLastOSError;


  try

    Result := ERROR_GEN_FAILURE;
    if Assigned(AProcessMessages) then
    begin
      repeat
        if not GetExitCodeProcess(SEI.hProcess, Result) then
          Result := ERROR_GEN_FAILURE;
        AProcessMessages;
      until Result <> STILL_ACTIVE;
    end
    else
    begin
      if WaitForSingleObject(SEI.hProcess, INFINITE) <> WAIT_OBJECT_0 then
        if not GetExitCodeProcess(SEI.hProcess, Result) then
          Result := ERROR_GEN_FAILURE;
    end;

  finally
    CloseHandle(SEI.hProcess);
  end;

end;


//code: http://www.delphitop.com/html/xitong/3062.html
function GetProcessIdentity(): String;
var
  hToken: THandle;
  UserName, UserDomain: String;
  cbName, cbDomainName: ULONG;
  ReturnLength: DWORD;
  Buff: array of Byte;
  tu: PTokenUser;
  peUse: SID_NAME_USE;
label
  Cleanup;
begin
  Result := '';
  // �򿪽�������
  if not OpenProcessToken(GetCurrentProcess, MAXIMUM_ALLOWED, hToken) then Exit;

  // ��ѯ�û��˻�����

  // ���Ȼ�ȡ��Ҫ�Ļ�������С
  if not GetTokenInformation(hToken, TokenUser, nil, 0, ReturnLength) then
    if GetLastError = ERROR_INSUFFICIENT_BUFFER then
    begin
      // ���û�������С
      SetLength(Buff, ReturnLength);
      // ��ȡ����
      GetTokenInformation(hToken, TokenUser, @Buff[0], ReturnLength, ReturnLength);
      tu := PTokenUser(@Buff[0]);
    end
    else
      goto Cleanup
  else
    goto Cleanup;

  // ͨ�� SID ��ѯ�û�������½��
  cbName := 0;
  cbDomainName := 0;
  // ��ȡ��Ҫ�Ļ�������С
  if not LookupAccountSid(nil, tu.User.Sid, nil, cbName, nil, cbDomainName, peUse) then
    if GetLastError = ERROR_INSUFFICIENT_BUFFER then
    begin
      // �����ַ�������(���� NULL �ַ�)
      SetLength(UserName, cbName);
      SetLength(UserDomain, cbDomainName);
      // ��ȡ����
      if LookupAccountSid(nil, tu.User.Sid, @UserName[1], cbName, @UserDomain[1], cbDomainName, peUse) then
      begin
        // ��ȥ���� NULL �ַ�
        SetLength(UserName, cbName);
        SetLength(UserDomain, cbDomainName);
      end
      else
        goto Cleanup;
    end
    else
      goto Cleanup
  else
    goto Cleanup;

  // �����Ϣ
  //Result := UserDomain + '\' + UserName;
  Result := UserName;
Cleanup:
  // �ر�����
  CloseHandle(hToken);
end;


function IsAdministrator: Boolean;
var
  psidAdmin: Pointer;
  B: BOOL;
const
  SECURITY_NT_AUTHORITY: TSidIdentifierAuthority = (Value: (0, 0, 0, 0, 0, 5));
  SECURITY_BUILTIN_DOMAIN_RID  = $00000020;
  DOMAIN_ALIAS_RID_ADMINS      = $00000220;
  SE_GROUP_USE_FOR_DENY_ONLY  = $00000010;
begin
  psidAdmin := nil;
  try
    // ����٧էѧק� SID �ԧ����� �ѧէާڧߧ�� �էݧ� ����ӧ֧�ܧ�
    //����һ�����Ա��SID������֤
    Win32Check(AllocateAndInitializeSid(SECURITY_NT_AUTHORITY, 2,
      SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0,
      psidAdmin));

    // �����ӧ֧��֧�, �ӧ��էڧ� �ݧ� �ާ� �� �ԧ����� �ѧէާڧߧ�� (�� ���ק��� �ӧ�֧� ����ӧ֧��� �ߧ� disabled SID)
    //����Ƿ�������Ա�飨���ǵ����н���SID�ļ�飩
    if CheckTokenMembership(0, psidAdmin, B) then
      Result := B
    else
      Result := False;
  finally
    if psidAdmin <> nil then
      FreeSid(psidAdmin);
  end;
end;

{$R-}

function IsAdministratorAccount: Boolean;
var
  psidAdmin: Pointer;
  Token: THandle;
  Count: DWORD;
  TokenInfo: PTokenGroups;
  HaveToken: Boolean;
  I: Integer;
const
  SECURITY_NT_AUTHORITY: TSidIdentifierAuthority = (Value: (0, 0, 0, 0, 0, 5));
  SECURITY_BUILTIN_DOMAIN_RID  = $00000020;
  DOMAIN_ALIAS_RID_ADMINS      = $00000220;
  SE_GROUP_USE_FOR_DENY_ONLY  = $00000010;
begin
  Result := Win32Platform <> VER_PLATFORM_WIN32_NT;
  if Result then
    Exit;

  psidAdmin := nil;
  TokenInfo := nil;
  HaveToken := False;
  try
    Token := 0;
    HaveToken := OpenThreadToken(GetCurrentThread, TOKEN_QUERY, True, Token);
    if (not HaveToken) and (GetLastError = ERROR_NO_TOKEN) then
      HaveToken := OpenProcessToken(GetCurrentProcess, TOKEN_QUERY, Token);
    if HaveToken then
    begin
      Win32Check(AllocateAndInitializeSid(SECURITY_NT_AUTHORITY, 2,
        SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0,
        psidAdmin));
      if GetTokenInformation(Token, TokenGroups, nil, 0, Count) or
         (GetLastError <> ERROR_INSUFFICIENT_BUFFER) then
        RaiseLastOSError;
      TokenInfo := PTokenGroups(AllocMem(Count));
      Win32Check(GetTokenInformation(Token, TokenGroups, TokenInfo, Count, Count));
      for I := 0 to TokenInfo^.GroupCount - 1 do
      begin
        Result := EqualSid(psidAdmin, TokenInfo^.Groups[I].Sid);
        if Result then
          Break;
      end;
    end;
  finally
    if TokenInfo <> nil then
      FreeMem(TokenInfo);
    if HaveToken then
      CloseHandle(Token);
    if psidAdmin <> nil then
      FreeSid(psidAdmin);
  end;
end;

{$R+}

function IsUACEnabled: Boolean;
var
  Reg: TRegistry;
begin
  Result := CheckWin32Version(6, 0);
  if Result then
  begin
    Reg := TRegistry.Create(KEY_READ);
    try
      Reg.RootKey := HKEY_LOCAL_MACHINE;
      if Reg.OpenKey('\Software\Microsoft\Windows\CurrentVersion\Policies\System', False) then
        if Reg.ValueExists('EnableLUA') then
          Result := (Reg.ReadInteger('EnableLUA') <> 0)
        else
          Result := False
      else
        Result := False;
    finally
      FreeAndNil(Reg);
    end;
  end;
end;

function IsElevated: Boolean;
const
  TokenElevation = TTokenInformationClass(20);
type
  TOKEN_ELEVATION = record
    TokenIsElevated: DWORD;
  end;
var
  TokenHandle: THandle;
  ResultLength: Cardinal;
  ATokenElevation: TOKEN_ELEVATION;
  HaveToken: Boolean;
begin
  if CheckWin32Version(6, 0) then
  begin
    TokenHandle := 0;
    HaveToken := OpenThreadToken(GetCurrentThread, TOKEN_QUERY, True, TokenHandle);
    if (not HaveToken) and (GetLastError = ERROR_NO_TOKEN) then
      HaveToken := OpenProcessToken(GetCurrentProcess, TOKEN_QUERY, TokenHandle);
    if HaveToken then
    begin
      try
        ResultLength := 0;
        if GetTokenInformation(TokenHandle, TokenElevation, @ATokenElevation, SizeOf(ATokenElevation), ResultLength) then
          Result := ATokenElevation.TokenIsElevated <> 0
        else
          Result := False;
      finally
        CloseHandle(TokenHandle);
      end;
    end
    else
      Result := False;
  end
  else
    Result := IsAdministrator;
end;

procedure SetButtonElevated(const AButtonHandle: THandle);
const
  BCM_SETSHIELD = $160C;
var
  Required: BOOL;
begin
  if not CheckWin32Version(6, 0) then
    Exit;
  if IsElevated then
    Exit;

  Required := True;
  SendMessage(AButtonHandle, BCM_SETSHIELD, 0, LPARAM(Required));
end;

procedure CheckForElevatedTask;

  function GetArgsForElevatedTask: String;

    function PrepareParam(const ParamNo: Integer): String;
    begin
      Result := ParamStr(ParamNo);
      if Pos(' ', Result) > 0 then
        Result := AnsiQuotedStr(Result, '"');
    end;

  var
    X: Integer;
  begin
    Result := '';
    for X := 1 to ParamCount do
    begin
      if (AnsiUpperCase(ParamStr(X)) = ('/' + RunElevatedTaskSwitch)) or
         (AnsiUpperCase(ParamStr(X)) = ('-' + RunElevatedTaskSwitch)) then
        Continue;

      Result := Result + PrepareParam(X) + ' ';
    end;

    Result := Trim(Result);
  end;

var
  ExitCode: Cardinal;
begin
  if not FindCmdLineSwitch(RunElevatedTaskSwitch) then
    Exit;

  ExitCode := ERROR_GEN_FAILURE;
  try
    if not IsElevated then
      ExitCode := ERROR_ACCESS_DENIED
    else
    if Assigned(OnElevateProc) then
      ExitCode := OnElevateProc(GetArgsForElevatedTask)
    else
      ExitCode := ERROR_PROC_NOT_FOUND;
  except
    on E: Exception do
    begin
      if E is EAbort then
        ExitCode := ERROR_CANCELLED
      else
      if E is EOleSysError then
        ExitCode := Cardinal(EOleSysError(E).ErrorCode)
      else
      if E is EOSError then
      else
        ExitCode := ERROR_GEN_FAILURE;
    end;
  end;

  if ExitCode = STILL_ACTIVE then
    ExitCode := ERROR_GEN_FAILURE;
  TerminateProcess(GetCurrentProcess, ExitCode);
end;

end.
