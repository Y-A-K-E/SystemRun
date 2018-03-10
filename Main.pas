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
    procedure UpSystem; //统一的提权处理
  public
    { Public declarations }
     procedure WmDropFiles(var Msg: TMessage); message WM_DROPFILES;  //消息处理 : 拖放文件
    //procedure WmDropFiles(var Msg: TWMDropFiles);message WM_DROPFILES ;
  end;

var
  MainForm: TMainForm;


  run_user :string;



  //提权到UAC运行
  function RunAsAdmin(hWnd: HWND; filename: string; Parameters: string ; _isshow:integer=1): Boolean;
  //创建文件夹快捷方式
  procedure CreateDirLink(ProgramPath,LinkPath, Descr: String);
  //添加UAC的拖放支持
  function UacDrag: Boolean;
  //判断文件是否符合格式
  function TestRunBin:boolean;
  //取快捷方式的源目标文件
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
begin


  // 非uac管理员运行权限,提权
  if not IsElevated then
    begin
        //非管理员账户
      if  (not IsAdministrator ) AND  (not IsAdministratorAccount) then
        begin
          ShowMessage('非管理员或者管理员组不可用');  exit;
        end else
          begin


          _isrun:= RunAsAdmin(Handle, ParamStr(0), ArgRunExe + ' /epath:'+ ExtractShortPathName(Edit_File_Path.Text) );

          if _isrun then
              Application.Terminate //提权成功自杀,由UAC提权后的自身程序处理剩余工作
            else
              showmessage('UAC提权失败,本次操作需要UAC提权运行,如果反复提示失败,请检查系统设置');

          end;



    end else
      begin
         //已经是UAC权限运行的
         //还需判断是否system用户运行


          //非SYSTEM用户运行,需要再次提权
          if run_user <> 'SYSTEM'  then
          begin
              Upsystem;
          end else
            begin

                if LowerCase(ExtractFileName(MainForm.Edit_File_Path.Text)) = 'explorer.exe' then
                  begin

                    if Application.MessageBox('处于种种考虑,不建议直接 使用system 运行 [explorer] 资源管理器' + #10#13+
                    '如果非得尝试,可以结束所有普通用户的explorer.exe进程先,'+ #10#13+
                    '再运行手写批处理命令运行explorer.'   + #10#13#10#13+
                    '这里有个伪方案,点击 [确定] 在弹出的文件浏览器中进行操作' + #10#13#10#13 +
                    '注意1: 因为UAC隔离,普通用户无法复制文件到system桌面' + #10#13 +

                    '注意2: C盘根目录所有用户都可访问,可用做复制中转点.' + #10#13 +
                    '注意3: 在这个弹出的文件管理器中,复制或粘贴文件尽量使用鼠标右键操作.'
                    , '创建system用户桌面', MB_OKCANCEL +
                    MB_ICONQUESTION) = IDOK then
                    begin


                      //WinExec(pansichar('taskkill /f /im explorer.exe'),SW_HIDE);
                      //WinExec(pansichar('taskkill /f /im explorer.exe'),SW_HIDE);
                      //WinExec(pansichar('taskkill /f /im explorer.exe'),SW_HIDE);

                      //ShellExecute(MainForm.Handle, 'open', pchar( MainForm.Edit_File_Path.Text), nil, nil, SW_SHOWNORMAL);

                      if MainForm.opendialog1.Execute then
                        begin
                          MainForm.Edit_File_Path.Text:=MainForm.opendialog1.FileName
                        end;

                    end;

                     //win10 下  即使用system运行了资源管理器,打开文件夹也会自动变成普通用户.
                  end else
                    begin
                       ShellExecute(MainForm.Handle, 'open', pchar( MainForm.Edit_File_Path.Text), nil, nil, SW_SHOWNORMAL);
                    end;
            end;

      end;


end;


//创建文件夹快捷方式,函数留着备用.暂时用不上
procedure CreateDirLink(ProgramPath,LinkPath, Descr: String);
var
  AnObj: IUnknown;
  ShellLink: IShellLink;
  AFile: IPersistFile;
  FileName: WideString;
begin
  if UpperCase(ExtractFileExt(LinkPath)) <> '.LNK' then //检查扩展名是否正确
  begin
    raise Exception.Create('快捷方式的扩展名必须是 ′′LNK′′!');
    //若不是则产生异常
  end;
try
  OleInitialize(nil);//初始化OLE库，在使用OLE函数前必须调用初始化
  AnObj :=CreateComObject(CLSID_ShellLink); //根据给定的ClassID生成
  //一个COM对象，此处是快捷方式
  ShellLink := AnObj as IShellLink;//强制转换为快捷方式接口
  AFile := AnObj as IPersistFile;//强制转换为文件接口
  //设置快捷方式属性，此处只设置了几个常用的属性
  ShellLink.SetPath(PChar(ProgramPath)); // 快捷方式的目标文件，一般为可执行文件
  //ShellLink.SetArguments(PChar(ProgramArg));// 目标文件参数
  ShellLink.SetWorkingDirectory(PChar(ExtractFilePath(ProgramPath)));//目标文件的工作目录
  ShellLink.SetDescription(PChar(Descr));// 对目标文件的描述
  FileName := LinkPath;//把文件名转换为WideString类型
  AFile.Save(PWChar(FileName), False);//保存快捷方式
finally
 OleUninitialize;//关闭OLE库，此函数必须与OleInitialize成对调用
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
  //提权到system准备.

  //第一步: 写出psexec.exe文件
 {$IFDEF WIN64}
  spath := GetEnvironmentVariable('TEMP')+ PsExecBinPath64;
 {$ELSE}
    spath := GetEnvironmentVariable('TEMP')+ PsExecBinPath;
 {$ENDIF}


     // psexec.exe文件不存在
  if not FileExists(spath) then
    begin
      //载入资源

     {$IFDEF WIN64}
      stream := TResourceStream.Create(HInstance, 'PSEXEC64',  RT_RCDATA);
     {$ELSE}
        stream := TResourceStream.Create(HInstance, 'PSEXEC',  RT_RCDATA);
     {$ENDIF}


      //保存到temp目录
      stream.SaveToFile(spath);
    end;

  //第二步骤,写注册表项  , 写入后 PsExec首次运行不会弹出提示.一切在静默状态运行
  //HKCU\Software\Sysinternals\PsExec\EulaAccepted = 1
  MyReg := TRegistry.Create;
  MyReg.RootKey := HKEY_CURRENT_USER;
  if not MyReg.KeyExists('Software\Sysinternals\PsExec\') then
  begin
    //key不存在则创建
    if MyReg.CreateKey('Software\Sysinternals\PsExec\') then
     begin
       if MyReg.OpenKey('Software\Sysinternals\PsExec\',true ) then
        begin
          MyReg.WriteInteger('EulaAccepted',1);
        end;
     end;
  end
  else
    //key存在
    if MyReg.OpenKey('Software\Sysinternals\PsExec\',true ) then
    begin
      MyReg.WriteInteger('EulaAccepted',1);
    end;

  MyReg.CloseKey;//关闭主键
  MyReg.Destroy;//释放内存

  //第三步,生成system的桌面,文档,图片,视频 相关的目录 ,避免打开使用system打开资源管理器报错.
  //%systemroot%\system32\config\systemprofile
  //未作兼容xp测试.
  _Desktop_path := GetEnvironmentVariable('SYSTEMROOT') + '\system32\config\systemprofile\Desktop\' ;
  _Documents_path :=  GetEnvironmentVariable('SYSTEMROOT') + '\system32\config\systemprofile\Documents\' ;

  if not DirectoryExists(_Desktop_path) then //判断目录是否存在
    try
      begin
        CreateDir(_Desktop_path);
      end;
    finally
      //raise Exception.Create('Cannot Create '+_Desktop_path);
    end;



  if not DirectoryExists(_Documents_path) then //判断目录是否存在
    try
      begin
        CreateDir(_Documents_path);
      end;
    finally
      //raise Exception.Create('Cannot Create '+_Documents_path);
    end;


  //第四步,用psexec 调用提权到system用户

  // psexec 存在
  if FileExists(spath) then
    begin
      arg_str:= '-i -d -s "' + ExtractShortPathName(paramstr(0)) + '" ' + ArgUac  + ' '  + '/epath:'+ExtractShortPathName(Edit_File_Path.Text) +'';   //保留参数
      ShellExecute(MainForm.Handle,nil,pchar(spath),pchar(arg_str),nil,SW_HIDE ) ;  //用隐藏窗口的方式调用  用psexec
      DeleteFile(spath); //清理
      Application.Terminate; //自杀,剩余部分交予新窗口处理.

    end;

end;

procedure TMainForm.Edit_File_PathChange(Sender: TObject);
//当文件载入自动判断
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
  //strList.Add('.DLL');   //rundll32 调用的dll格式.  临时放弃,有需要可以通过创建一个cmd文件运行.
  strList.Add('.COM');
  strList.Add('.CMD');
  strList.Add('.BAT');
  //vbs js 之类的文件也请通过写入一个bat批处理文件,然后用cmd的方式调用

  //处理快捷方式文件.
  if UpperCase(ExtractFileExt(MainForm.Edit_File_Path.Text)) = '.LNK' then
    begin
      MainForm.Edit_File_Path.Text:=GetTargetOfShorCut(MainForm.Edit_File_Path.Text);
    end;


  _ext:=UpperCase(ExtractFileExt(MainForm.Edit_File_Path.Text)) ; //提取后缀,大写
  //格式处理
  case strList.IndexOf(_ext) of
    0: _r := true;
    1: _r := true;
    2: _r := true;
    3: _r := true;
    //4: _r := true;
  else
    _r := false;
  end;
  MainForm.button1.Enabled := _r;   //按钮是否可以按

  if not _r then MainForm.Edit_File_Path.Text := '';   //格式不合法删空
  Result:= _r;

end;


//双击编辑框
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

  //拖放支持 .UAC 特别处理.
  DragAcceptFiles(MainForm.Handle, True);
  UacDrag;


  //读取参数,填入Edit_File_Path中.  /epath:
  if ParamCount >=2 then
  begin
     if LeftStr(ParamStr(2), 7) = '/epath:'  then
      begin
        Edit_File_Path.Text :=  StringReplace (ParamStr(2), '/epath:', '', []);
      end;

  end;

  run_user := UpperCase(GetProcessIdentity()) ; //获取当前程序运行用户//大写


  //命令行参数
  if (ParamCount >0) and SameText(ParamStr(1), ArgRunExe ) then
  begin
     //showmessage(ParamStr(1));
     //这时一定需要是管理员权限运行.

    if not IsElevated then
    begin
       showmessage('请确保UAC授权成功.本次提权失败,程序自动推出.');
       Application.Terminate;
    end else
      begin
        //UAC管理员权限

        //非SYSTEM用户运行,需要再次提权
        if run_user <> 'SYSTEM'  then
        begin
          Upsystem;
        end;

      end;

  end
  else
  if (ParamCount >0) and SameText(ParamStr(1), ArgUac ) then
  begin
    //非SYSTEM用户运行,需要再次提权
    if run_user <> 'SYSTEM'  then
    begin
       showmessage('本次提权SYSTEM失败,程序自动推出.');
       Application.Terminate;
    end else
      begin
        //运行
         MainForm.Button1.Click;
      end;
  end;

  StatusBar1.Panels[0].Text := Format('运行用户: %s',             [run_user]);
  StatusBar1.Panels[1].Text := Format('UAC提权: %s',             [BoolToStr(IsElevated, True)]);

  //设置按钮盾牌
  SetButtonElevated(Button1.Handle);

end;




//打开网页链接
procedure TMainForm.StatusBar1Click(Sender: TObject);
const
  sCommand = 'https://www.yge.me/';
begin
  //非system用户直接调用默认浏览器
  //system调用ie ,防止chrome或者firefox之类浏览器创建一个system的配置
  if run_user <> 'SYSTEM'  then
  begin
     ShellExecute(0, 'OPEN', PChar(sCommand), '', '', SW_SHOWNORMAL);
  end else
    begin
      ShellExecute(0, 'OPEN', PChar('iexplore.exe'), pChar(sCommand) , '', SW_SHOWNORMAL);
    end;
end;

//拖放消息处理.
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
    DragQueryFile(Msg.WParam, 0, buffer, sizeof(buffer)); //第一个文件
    Edit_File_Path.text:=buffer;
end;




//为UAC加入拖放支持.
//code:http://www.fx114.net/qa-29-4738.aspx
//codd:https://github.com/cheat-engine/cheat-engine/blob/master/Cheat%20Engine/MainUnit.pas
function UacDrag: Boolean;
const
  MSGFLT_ADD = 1;
var
  ChangeWindowMessageFilter: function(msg: Cardinal; dwFlag: Word): BOOL; stdcall;
begin
  @ChangeWindowMessageFilter := GetProcAddress(LoadLibrary('user32.dll'), 'ChangeWindowMessageFilter');


  //vista 以上系统
  if TOSVersion.Major >= 6 then
  begin
    try
     //WM_COPYGLOBALDATA = 73; MSGFLT_ADD = 1
     ChangeWindowMessageFilter(73, MSGFLT_ADD);
     ChangeWindowMessageFilter(WM_DROPFILES, MSGFLT_ADD);
     ChangeWindowMessageFilter(WM_COPYDATA, MSGFLT_ADD);  //非必要


      {$IFDEF WIN64}
          //64位程序在UAC下未知bug,需要执行两次才能拖放.
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
//取快捷方式的目标源
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
