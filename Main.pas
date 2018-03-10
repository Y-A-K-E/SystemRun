{

    Author: Y.A.K.E
    License : MIT
    Date : 2018/02/01

}
unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Registry, ShellAPI, ComObj,
  Vcl.ExtCtrls,strUtils,tlhelp32,ActiveX,ShlObj, Vcl.ComCtrls,
  Vcl.Imaging.pngimage ;




type
  [ComponentPlatformsAttribute(pidWin32 or pidWin64)]
  TMainForm = class(TForm)
    Edit_File_Path: TEdit;
    Button1: TButton;
    StatusBar1: TStatusBar;
    Label1: TLabel;
    OpenDialog1: TOpenDialog;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Image1: TImage;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Edit_File_PathDblClick(Sender: TObject);
    procedure Edit_File_PathChange(Sender: TObject);
    procedure StatusBar1Click(Sender: TObject);

  private
    { Private declarations }
    procedure UpSystem; //ͳһ����Ȩ����
  public
    { Public declarations }
     procedure WmDropFiles(var Msg: TMessage); message WM_DROPFILES;  //��Ϣ���� : �Ϸ��ļ�
    //procedure WmDropFiles(var Msg: TWMDropFiles);message WM_DROPFILES ;
  end;

var
  MainForm: TMainForm;


  run_user :string;



  //��Ȩ��UAC����
  function RunAsAdmin(hWnd: HWND; filename: string; Parameters: string ; _isshow:integer=1): Boolean;
  //�����ļ��п�ݷ�ʽ
  procedure CreateDirLink(ProgramPath,LinkPath, Descr: String);
  //���UAC���Ϸ�֧��
  function UacDrag: Boolean;
  //�ж��ļ��Ƿ���ϸ�ʽ
  function TestRunBin:boolean;
  //ȡ��ݷ�ʽ��ԴĿ���ļ�
  function GetTargetOfShorCut(LinkFile:string):string;

const
  ArgUac    = '/uac';
  ArgRunExe = '/run_exe';
  PsExecBinPath = '\ypsexec.exe';
  PsExecBinPath64 = '\ypsexec64.exe';
  BatBinPath = '\yrunbat.bat';

  MSGFLT_ADD= 1 ;

  MSGFLT_REMOVE= 2 ;
implementation
uses
  RunElevatedSupport;
{$R *.dfm}









procedure TMainForm.Button1Click(Sender: TObject);
var
 _isrun: boolean;
 _s:string;
begin


  // ��uac����Ա����Ȩ��,��Ȩ
  if not IsElevated then
    begin
        //�ǹ���Ա�˻�
      if  (not IsAdministrator ) AND  (not IsAdministratorAccount) then
        begin
          ShowMessage('�ǹ���Ա���߹���Ա�鲻����');  exit;
        end else
          begin


          _isrun:= RunAsAdmin(Handle, ParamStr(0), ArgRunExe + ' /epath:'+ ExtractShortPathName(Edit_File_Path.Text) );

          if _isrun then
              Application.Terminate //��Ȩ�ɹ���ɱ,��UAC��Ȩ������������ʣ�๤��
            else
              showmessage('UAC��Ȩʧ��,���β�����ҪUAC��Ȩ����,���������ʾʧ��,����ϵͳ����');

          end;



    end else
      begin
         //�Ѿ���UACȨ�����е�
         //�����ж��Ƿ�system�û�����


          //��SYSTEM�û�����,��Ҫ�ٴ���Ȩ
          if run_user <> 'SYSTEM'  then
          begin
              Upsystem;
          end else
            begin

                if LowerCase(ExtractFileName(MainForm.Edit_File_Path.Text)) = 'explorer.exe' then
                  begin

                    if Application.MessageBox('�������ֿ���,������ֱ�� ʹ��system ���� [explorer] ��Դ������' + #10#13+
                    '����ǵó���,���Խ���������ͨ�û���explorer.exe������,'+ #10#13+
                    '��������д��������������explorer.'   + #10#13#10#13+
                    '�����и�α����,��� [ȷ��] �ڵ������ļ�������н��в���' + #10#13#10#13 +
                    'ע��1: ��ΪUAC����,��ͨ�û��޷������ļ���system����' + #10#13 +

                    'ע��2: C�̸�Ŀ¼�����û����ɷ���,������������ת��.' + #10#13 +
                    'ע��3: ������������ļ���������,���ƻ�ճ���ļ�����ʹ������Ҽ�����.'
                    , '����system�û�����', MB_OKCANCEL +
                    MB_ICONQUESTION) = IDOK then
                    begin



                      //WinExec(pansichar('taskkill /f /im explorer.exe'),SW_HIDE);
                      //WinExec(pansichar('taskkill /f /im explorer.exe'),SW_HIDE);
                      //WinExec(pansichar('taskkill /f /im explorer.exe'),SW_HIDE);

                      //ShellExecute(MainForm.Handle, 'open', pchar( MainForm.Edit_File_Path.Text), nil, nil, SW_SHOWNORMAL);

                      _s:=MainForm.opendialog1.Filter;
                      MainForm.opendialog1.Filter:='';

                      if MainForm.opendialog1.Execute then
                        begin
                          MainForm.Edit_File_Path.Text:=MainForm.opendialog1.FileName
                        end;
                      MainForm.opendialog1.Filter:=_s;

                    end;

                     //win10 ��  ��ʹ��system��������Դ������,���ļ���Ҳ���Զ������ͨ�û�.
                  end else
                    begin
                       ShellExecute(MainForm.Handle, 'open', pchar( MainForm.Edit_File_Path.Text), nil, nil, SW_SHOWNORMAL);
                    end;
            end;

      end;


end;


//�����ļ��п�ݷ�ʽ,�������ű���.��ʱ�ò���
procedure CreateDirLink(ProgramPath,LinkPath, Descr: String);
var
  AnObj: IUnknown;
  ShellLink: IShellLink;
  AFile: IPersistFile;
  FileName: WideString;
begin
  if UpperCase(ExtractFileExt(LinkPath)) <> '.LNK' then //�����չ���Ƿ���ȷ
  begin
    raise Exception.Create('��ݷ�ʽ����չ�������� ���LNK���!');
    //������������쳣
  end;
try
  OleInitialize(nil);//��ʼ��OLE�⣬��ʹ��OLE����ǰ������ó�ʼ��
  AnObj :=CreateComObject(CLSID_ShellLink); //���ݸ�����ClassID����
  //һ��COM���󣬴˴��ǿ�ݷ�ʽ
  ShellLink := AnObj as IShellLink;//ǿ��ת��Ϊ��ݷ�ʽ�ӿ�
  AFile := AnObj as IPersistFile;//ǿ��ת��Ϊ�ļ��ӿ�
  //���ÿ�ݷ�ʽ���ԣ��˴�ֻ�����˼������õ�����
  ShellLink.SetPath(PChar(ProgramPath)); // ��ݷ�ʽ��Ŀ���ļ���һ��Ϊ��ִ���ļ�
  //ShellLink.SetArguments(PChar(ProgramArg));// Ŀ���ļ�����
  ShellLink.SetWorkingDirectory(PChar(ExtractFilePath(ProgramPath)));//Ŀ���ļ��Ĺ���Ŀ¼
  ShellLink.SetDescription(PChar(Descr));// ��Ŀ���ļ�������
  FileName := LinkPath;//���ļ���ת��ΪWideString����
  AFile.Save(PWChar(FileName), False);//�����ݷ�ʽ
finally
 OleUninitialize;//�ر�OLE�⣬�˺���������OleInitialize�ɶԵ���
end;

end;




procedure TMainForm.UpSystem;
var
  MyReg:TRegistry;


  spath:string;
  stream : TResourceStream;
  arg_str:string;
  _Desktop_path:string;
  _Documents_path:string;

begin
  //��Ȩ��system׼��.

  //��һ��: д��psexec.exe�ļ�
 {$IFDEF WIN64}
  spath := GetEnvironmentVariable('TEMP')+ PsExecBinPath64;
 {$ELSE}
    spath := GetEnvironmentVariable('TEMP')+ PsExecBinPath;
 {$ENDIF}


     // psexec.exe�ļ�������
  if not FileExists(spath) then
    begin
      //������Դ

     {$IFDEF WIN64}
      stream := TResourceStream.Create(HInstance, 'PSEXEC64',  RT_RCDATA);
     {$ELSE}
        stream := TResourceStream.Create(HInstance, 'PSEXEC',  RT_RCDATA);
     {$ENDIF}


      //���浽tempĿ¼
      stream.SaveToFile(spath);
    end;

  //�ڶ�����,дע�����  , д��� PsExec�״����в��ᵯ����ʾ.һ���ھ�Ĭ״̬����
  //HKCU\Software\Sysinternals\PsExec\EulaAccepted = 1
  MyReg := TRegistry.Create;
  MyReg.RootKey := HKEY_CURRENT_USER;
  if not MyReg.KeyExists('Software\Sysinternals\PsExec\') then
  begin
    //key�������򴴽�
    if MyReg.CreateKey('Software\Sysinternals\PsExec\') then
     begin
       if MyReg.OpenKey('Software\Sysinternals\PsExec\',true ) then
        begin
          MyReg.WriteInteger('EulaAccepted',1);
        end;
     end;
  end
  else
    //key����
    if MyReg.OpenKey('Software\Sysinternals\PsExec\',true ) then
    begin
      MyReg.WriteInteger('EulaAccepted',1);
    end;

  MyReg.CloseKey;//�ر�����
  MyReg.Destroy;//�ͷ��ڴ�

  //������,����system������,�ĵ�,ͼƬ,��Ƶ ��ص�Ŀ¼ ,�����ʹ��system����Դ����������.
  //%systemroot%\system32\config\systemprofile
  //δ������xp����.
  _Desktop_path := GetEnvironmentVariable('SYSTEMROOT') + '\system32\config\systemprofile\Desktop\' ;
  _Documents_path :=  GetEnvironmentVariable('SYSTEMROOT') + '\system32\config\systemprofile\Documents\' ;

  if not DirectoryExists(_Desktop_path) then //�ж�Ŀ¼�Ƿ����
    try
      begin
        CreateDir(_Desktop_path);
      end;
    finally
      //raise Exception.Create('Cannot Create '+_Desktop_path);
    end;



  if not DirectoryExists(_Documents_path) then //�ж�Ŀ¼�Ƿ����
    try
      begin
        CreateDir(_Documents_path);
      end;
    finally
      //raise Exception.Create('Cannot Create '+_Documents_path);
    end;


  //���Ĳ�,��psexec ������Ȩ��system�û�

  // psexec ����
  if FileExists(spath) then
    begin
      arg_str:= '-i -d -s "' + ExtractShortPathName(paramstr(0)) + '" ' + ArgUac  + ' '  + '/epath:'+ExtractShortPathName(Edit_File_Path.Text) +'';   //��������
      ShellExecute(MainForm.Handle,nil,pchar(spath),pchar(arg_str),nil,SW_HIDE ) ;  //�����ش��ڵķ�ʽ����  ��psexec
      DeleteFile(spath); //����
      Application.Terminate; //��ɱ,ʣ�ಿ�ֽ����´��ڴ���.

    end;

end;

procedure TMainForm.Edit_File_PathChange(Sender: TObject);
//���ļ������Զ��ж�
begin
  if Edit_File_Path.Text <> '' then
  begin
    TestRunBin;
  end;
end;



function TestRunBin:boolean;
var
_ext:string;
_r:boolean;
strList: TStringList;

begin
  strList := TStringList.Create;
  strList.Add('.EXE');
  //strList.Add('.DLL');   //rundll32 ���õ�dll��ʽ.  ��ʱ����,����Ҫ����ͨ������һ��cmd�ļ�����.
  strList.Add('.COM');
  strList.Add('.CMD');
  strList.Add('.BAT');
  //vbs js ֮����ļ�Ҳ��ͨ��д��һ��bat�������ļ�,Ȼ����cmd�ķ�ʽ����

  //�����ݷ�ʽ�ļ�.
  if UpperCase(ExtractFileExt(MainForm.Edit_File_Path.Text)) = '.LNK' then
    begin
      MainForm.Edit_File_Path.Text:=GetTargetOfShorCut(MainForm.Edit_File_Path.Text);
    end;


  _ext:=UpperCase(ExtractFileExt(MainForm.Edit_File_Path.Text)) ; //��ȡ��׺,��д
  //��ʽ����
  case strList.IndexOf(_ext) of
    0: _r := true;
    1: _r := true;
    2: _r := true;
    3: _r := true;
    //4: _r := true;
  else
    _r := false;
  end;
  MainForm.button1.Enabled := _r;   //��ť�Ƿ���԰�

  if not _r then MainForm.Edit_File_Path.Text := '';   //��ʽ���Ϸ�ɾ��
  Result:= _r;

end;


//˫���༭��
procedure TMainForm.Edit_File_PathDblClick(Sender: TObject);
begin
    if MainForm.opendialog1.Execute then
      begin
        MainForm.Edit_File_Path.Text:=MainForm.opendialog1.FileName
      end;
end;


function RunAsAdmin(hWnd: HWND; filename: string; Parameters: string ; _isshow:integer=1): Boolean;
// See Step 3: Redesign for UAC Compatibility (UAC)
// http://msdn.microsoft.com/en-us/library/bb756922.aspx
// This code is released into the public domain. No attribution required.
//eg:
//RunAsAdmin(Handle, 'c:\Windows\system32\cmd.exe', '');
//code:https://www.cnblogs.com/findumars/p/5001753.html
var
  sei: TShellExecuteInfo;
begin
  ZeroMemory(@sei, SizeOf(sei));
  sei.cbSize := SizeOf(TShellExecuteInfo);
  sei.Wnd := hwnd;
  sei.fMask := SEE_MASK_FLAG_DDEWAIT or SEE_MASK_FLAG_NO_UI;
  sei.lpVerb := PChar('runas');
  sei.lpFile := PChar(Filename); // PAnsiChar;
  if parameters <> '' then
      sei.lpParameters := PChar(parameters); // PAnsiChar;

  if _isshow = 1 then
    sei.nShow := SW_SHOWNORMAL //Integer;
    else
      sei.nShow :=SW_HIDE;

  Result := ShellExecuteEx(@sei);
end;




procedure TMainForm.FormCreate(Sender: TObject);
begin


   //ShowMessage(Format('64bit: %s',             [BoolToStr(is64Bit, True)]));

  //�Ϸ�֧�� .UAC �ر���.
  DragAcceptFiles(MainForm.Handle, True);
  UacDrag;


  //��ȡ����,����Edit_File_Path��.  /epath:
  if ParamCount >=2 then
  begin
     if LeftStr(ParamStr(2), 7) = '/epath:'  then
      begin
        Edit_File_Path.Text :=  StringReplace (ParamStr(2), '/epath:', '', []);
      end;

  end;

  run_user := UpperCase(GetProcessIdentity()) ; //��ȡ��ǰ���������û�//��д


  //�����в���
  if (ParamCount >0) and SameText(ParamStr(1), ArgRunExe ) then
  begin
     //showmessage(ParamStr(1));
     //��ʱһ����Ҫ�ǹ���ԱȨ������.

    if not IsElevated then
    begin
       showmessage('��ȷ��UAC��Ȩ�ɹ�.������Ȩʧ��,�����Զ��Ƴ�.');
       Application.Terminate;
    end else
      begin
        //UAC����ԱȨ��

        //��SYSTEM�û�����,��Ҫ�ٴ���Ȩ
        if run_user <> 'SYSTEM'  then
        begin
          Upsystem;
        end;

      end;

  end
  else
  if (ParamCount >0) and SameText(ParamStr(1), ArgUac ) then
  begin
    //��SYSTEM�û�����,��Ҫ�ٴ���Ȩ
    if run_user <> 'SYSTEM'  then
    begin
       showmessage('������ȨSYSTEMʧ��,�����Զ��Ƴ�.');
       Application.Terminate;
    end else
      begin
        //����
         MainForm.Button1.Click;
      end;
  end;

  StatusBar1.Panels[0].Text := Format('�����û�: %s',             [run_user]);
  StatusBar1.Panels[1].Text := Format('UAC��Ȩ: %s',             [BoolToStr(IsElevated, True)]);

  //���ð�ť����
  SetButtonElevated(Button1.Handle);

end;




//����ҳ����
procedure TMainForm.StatusBar1Click(Sender: TObject);
const
  sCommand = 'https://www.yge.me/';
begin
  //��system�û�ֱ�ӵ���Ĭ�������
  //system����ie ,��ֹchrome����firefox֮�����������һ��system������
  if run_user <> 'SYSTEM'  then
  begin
     ShellExecute(0, 'OPEN', PChar(sCommand), '', '', SW_SHOWNORMAL);
  end else
    begin
      ShellExecute(0, 'OPEN', PChar('iexplore.exe'), pChar(sCommand) , '', SW_SHOWNORMAL);
    end;
end;

//�Ϸ���Ϣ����.
//code:http://www.delphitop.com/html/wenjian/3072.html
//code:http://owlsperspective.blogspot.jp/2008/07/dragacceptfiles.html
//code:https://www.cnblogs.com/gameking/archive/2012/10/25/2738789.html
 procedure TMainForm.WmDropFiles(var Msg: TMessage);
//procedure TMainForm.WmDropFiles(var Msg: TWMDropFiles);
var
   buffer: array[0..1024] of Char;
begin
    Inherited;
    Edit_File_Path.text:='';
    buffer[0] := #0;
    DragQueryFile(Msg.WParam, 0, buffer, sizeof(buffer)); //��һ���ļ�
    Edit_File_Path.text:=buffer;
end;




//ΪUAC�����Ϸ�֧��.
//code:http://www.fx114.net/qa-29-4738.aspx
//codd:https://github.com/cheat-engine/cheat-engine/blob/master/Cheat%20Engine/MainUnit.pas
function UacDrag: Boolean;
const
  MSGFLT_ADD = 1;
var
  ChangeWindowMessageFilter: function(msg: Cardinal; dwFlag: Word): BOOL; stdcall;
begin
  @ChangeWindowMessageFilter := GetProcAddress(LoadLibrary('user32.dll'), 'ChangeWindowMessageFilter');


  //vista ����ϵͳ
  if TOSVersion.Major >= 6 then
  begin
    try
     //WM_COPYGLOBALDATA = 73; MSGFLT_ADD = 1
     ChangeWindowMessageFilter(73, MSGFLT_ADD);
     ChangeWindowMessageFilter(WM_DROPFILES, MSGFLT_ADD);
     ChangeWindowMessageFilter(WM_COPYDATA, MSGFLT_ADD);  //�Ǳ�Ҫ


      {$IFDEF WIN64}
          //64λ������UAC��δ֪bug,��Ҫִ�����β����Ϸ�.
         ChangeWindowMessageFilter(73, 1);
         ChangeWindowMessageFilter(WM_DROPFILES, MSGFLT_ADD);
         ChangeWindowMessageFilter(WM_COPYDATA, MSGFLT_ADD);
      {$ENDIF}


     Result:= True;
    except;
      Result:= False;
    end;
  end else
   Result:= False;
end;



//code:http://blog.csdn.net/zhouzuoji/article/details/1865411
//code2:http://bbs.csdn.net/topics/390276355
//ȡ��ݷ�ʽ��Ŀ��Դ
function GetTargetOfShorCut(LinkFile:string):string;
const
  IID_IPersistFile:TGUID = '{0000010B-0000-0000-C000-000000000046}';
var
   intfLink:IShellLink;
   IntfPersist:IPersistFile;
   pfd:_WIN32_FIND_DATA;
   bSuccess:Boolean;
begin
   Result:='';
   IntfLink:=CreateComObject(CLSID_ShellLink) as IShellLink;
   SetString(Result,nil,MAX_PATH);
   bSuccess:=(IntfLink<>nil) and SUCCEEDED(IntfLink.QueryInterface(IID_IPersistFile,IntfPersist))
    and SUCCEEDED(IntfPersist.Load(PWideChar(WideString(LinkFile)),STGM_READ)) and
    SUCCEEDED(intfLink.GetPath(PChar(Result),MAX_PATH,pfd,SLGP_RAWPATH));
   if not bSuccess then Result:='';
end;


end.
