unit UfrmMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms,
  Menus, StdCtrls, Buttons, ADODB,
  ComCtrls, ToolWin, ExtCtrls,
  inifiles,Dialogs,
  StrUtils, DB, ComObj,Variants,ShellAPI, CoolTrayIcon, Grids, DBGrids;

type
  TfrmMain = class(TForm)
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    ADOConnection1: TADOConnection;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    ToolButton8: TToolButton;
    ToolButton2: TToolButton;
    Memo1: TMemo;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    ToolButton5: TToolButton;
    ToolButton9: TToolButton;
    OpenDialog1: TOpenDialog;
    ToolButton7: TToolButton;
    SaveDialog1: TSaveDialog;
    ADOConn_BS: TADOConnection;
    Timer1: TTimer;
    LYTray1: TCoolTrayIcon;
    procedure N3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure N1Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure ToolButton7Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
    procedure UpdateConfig;{配置文件生效}
    function MakeDBConn:boolean;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses ucommfunction;

const
  CR=#$D+#$A;
  STX=#$2;ETX=#$3;ACK=#$6;NAK=#$15;EOT=#$4;ETB=#$17;
  sCryptSeed='lc';//加解密种子
  //SEPARATOR=#$1C;
  sCONNECTDEVELOP='错误!请与开发商联系!' ;
  IniSection='Setup';

var
  ConnectString:string;
  SpecStatus:string ;//
  LisFormCaption:string;//
  EquipChar:string;
  ifRecLog:boolean;//是否记录调试日志

  DaanConnStr:string;
  ifConnSucc:boolean;

  RFM:STRING;       //返回数据
  hnd:integer;
  bRegister:boolean;

{$R *.dfm}

function ExecSQLCmd(AConnectionString:string;ASQL:string):integer;
var
  Conn:TADOConnection;
  Qry:TAdoQuery;
begin
  Conn:=TADOConnection.Create(nil);
  Conn.LoginPrompt:=false;
  Conn.ConnectionString:=AConnectionString;
  Qry:=TAdoQuery.Create(nil);
  Qry.Connection:=Conn;
  Qry.Close;
  Qry.SQL.Clear;
  Qry.SQL.Text:=ASQL;
  Try
    Result:=Qry.ExecSQL;
  except
    on E:Exception do
    begin
      WriteLog(pchar('函数ExecSQLCmd失败:'+E.Message+'。错误的SQL:'+ASQL));
      Result:=-1;
    end;
  end;
  Qry.Free;
  Conn.Free;
end;

function ScalarSQLCmd(AConnectionString:string;ASQL:string):string;
var
  Conn:TADOConnection;
  Qry:TAdoQuery;
begin
  Result:='';
  Conn:=TADOConnection.Create(nil);
  Conn.LoginPrompt:=false;
  Conn.ConnectionString:=AConnectionString;
  Qry:=TAdoQuery.Create(nil);
  Qry.Connection:=Conn;
  Qry.Close;
  Qry.SQL.Clear;
  Qry.SQL.Text:=ASQL;
  Try
    Qry.Open;
  except
    on E:Exception do
    begin
      WriteLog(pchar('函数ScalarSQLCmd失败:'+E.Message+'。错误的SQL:'+ASQL));
      Qry.Free;
      Conn.Free;
      exit;
    end;
  end;
  Result:=Qry.Fields[0].AsString;
  Qry.Free;
  Conn.Free;
end;

function ifRegister:boolean;
var
  HDSn,RegisterNum,EnHDSn:string;
  configini:tinifile;
  pEnHDSn:Pchar;
begin
  result:=false;
  
  HDSn:=GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'');

  CONFIGINI:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  RegisterNum:=CONFIGINI.ReadString(IniSection,'RegisterNum','');
  CONFIGINI.Free;
  pEnHDSn:=EnCryptStr(Pchar(HDSn),sCryptSeed);
  EnHDSn:=StrPas(pEnHDSn);

  if Uppercase(EnHDSn)=Uppercase(RegisterNum) then result:=true;

  if not result then messagedlg('对不起,您没有注册或注册码错误,请注册!',mtinformation,[mbok],0);
end;

function GetConnectString:string;
var
  Ini:tinifile;
  userid, password, datasource, initialcatalog: string;
  ifIntegrated:boolean;//是否集成登录模式

  pInStr,pDeStr:Pchar;
  i:integer;
begin
  result:='';
  
  Ini := tinifile.Create(ChangeFileExt(Application.ExeName,'.INI'));
  datasource := Ini.ReadString('连接数据库', '服务器', '');
  initialcatalog := Ini.ReadString('连接数据库', '数据库', '');
  ifIntegrated:=ini.ReadBool('连接数据库','集成登录模式',false);
  userid := Ini.ReadString('连接数据库', '用户', '');
  password := Ini.ReadString('连接数据库', '口令', '107DFC967CDCFAAF');
  Ini.Free;
  //======解密password
  pInStr:=pchar(password);
  pDeStr:=DeCryptStr(pInStr,sCryptSeed);
  setlength(password,length(pDeStr));
  for i :=1  to length(pDeStr) do password[i]:=pDeStr[i-1];
  //==========

  result := result + 'user id=' + UserID + ';';
  result := result + 'password=' + Password + ';';
  result := result + 'data source=' + datasource + ';';
  result := result + 'Initial Catalog=' + initialcatalog + ';';
  result := result + 'provider=' + 'SQLOLEDB.1' + ';';
  //Persist Security Info,表示ADO在数据库连接成功后是否保存密码信息
  //ADO缺省为True,ADO.net缺省为False
  //程序中会传ADOConnection信息给TADOLYQuery,故设置为True
  result := result + 'Persist Security Info=True;';
  if ifIntegrated then
    result := result + 'Integrated Security=SSPI;';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  rfm:='';
  
  ConnectString:=GetConnectString;
  UpdateConfig;
  if ifRegister then bRegister:=true else bRegister:=false;  

  Caption:='数据接收服务'+ExtractFileName(Application.ExeName);
  lytray1.Hint:='数据接收服务'+ExtractFileName(Application.ExeName);
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action:=caNone;
  LYTray1.HideMainForm;
end;

procedure TfrmMain.N3Click(Sender: TObject);
begin
  if (MessageDlg('退出后将不再接收设备数据,确定退出吗？', mtWarning, [mbYes, mbNo], 0) <> mrYes) then exit;
  application.Terminate;
end;

procedure TfrmMain.N1Click(Sender: TObject);
begin
  LYTray1.ShowMainForm;
end;

procedure TfrmMain.UpdateConfig;
var
  INI:tinifile;
  autorun:boolean;
begin
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));

  autorun:=ini.readBool(IniSection,'开机自动运行',false);
  ifRecLog:=ini.readBool(IniSection,'调试日志',false);

  EquipChar:=trim(uppercase(ini.ReadString(IniSection,'仪器字母','')));//读出来是大写就万无一失了
  SpecStatus:=ini.ReadString(IniSection,'默认样本状态','');

  LisFormCaption:=ini.ReadString(IniSection,'检验系统窗体标题','');

  DaanConnStr:=ini.ReadString(IniSection,'连接华银数据库','');

  ini.Free;

  OperateLinkFile(application.ExeName,'\'+ChangeFileExt(ExtractFileName(Application.ExeName),'.lnk'),15,autorun);

  try
    ADOConn_BS.Connected := false;
    ADOConn_BS.ConnectionString := DaanConnStr;
    ADOConn_BS.Connected := true;
    ifConnSucc:=true;
  except
    on E:Exception do
    begin
      ifConnSucc:=false;
      MESSAGEDLG('连接华银数据库失败!'+E.Message,mtError,[mbOK],0);
    end;
  end;
end;

function TfrmMain.MakeDBConn:boolean;
var
  newconnstr,ss: string;
  Label labReadIni;
begin
  result:=false;

  labReadIni:
  newconnstr := GetConnectString;
  try
    ADOConnection1.Connected := false;
    ADOConnection1.ConnectionString := newconnstr;
    ADOConnection1.Connected := true;
    result:=true;
  except
  end;
  if not result then
  begin
    ss:='服务器'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '数据库'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '集成登录模式'+#2+'CheckListBox'+#2+#2+'0'+#2+#2+#3+
        '用户'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '口令'+#2+'Edit'+#2+#2+'0'+#2+#2+'1';
    if ShowOptionForm('连接数据库','连接数据库',Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
      goto labReadIni else application.Terminate;
  end;
end;

procedure TfrmMain.ToolButton2Click(Sender: TObject);
var
  ss:string;
begin
  ss:='连接华银数据库'+#2+'DBConn'+#2+#2+'1'+#2+#2+#3+
      '仪器字母'+#2+'Edit'+#2+#2+'1'+#2+'申请单视图v_cm_sample使用H,故联机字母必需为H'+#2+#3+
      '检验系统窗体标题'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '默认样本状态'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '开机自动运行'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '调试日志'+#2+'CheckListBox'+#2+#2+'0'+#2+'注:强烈建议在正常运行时关闭'+#2+#3;

  if ShowOptionForm('',Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
	  UpdateConfig;
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
begin
  Memo1.Lines.Clear;
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
  SaveDialog1.DefaultExt := '.txt';
  SaveDialog1.Filter := 'txt (*.txt)|*.txt';
  if not SaveDialog1.Execute then exit;
  memo1.Lines.SaveToFile(SaveDialog1.FileName);
  showmessage('保存成功!');
end;

procedure TfrmMain.ToolButton5Click(Sender: TObject);
var
  ss:string;
begin
  ss:='RegisterNum'+#2+'Edit'+#2+#2+'0'+#2+'将该窗体标题栏上的字符串发给开发者,以获取注册码'+#2;
  if bRegister then exit;
  if ShowOptionForm(Pchar('注册:'+GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'')),Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
    if ifRegister then bRegister:=true else bRegister:=false;
end;

procedure TfrmMain.ToolButton7Click(Sender: TObject);
begin
  if MakeDBConn then ConnectString:=GetConnectString;
end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
VAR
  adotemp22:tadoquery;
  ReceiveItemInfo:OleVariant;
  FInts:OleVariant;
  check_date:String;
  checkid:String;
begin
  if not ifConnSucc then exit;

  (Sender as TTimer).Enabled:=false;

  if length(memo1.Lines.Text)>=60000 then memo1.Lines.Clear;//memo只能接受64K个字符

  adotemp22:=tadoquery.Create(nil);
  adotemp22.Connection:=ADOConn_BS;
  adotemp22.Close;
  adotemp22.SQL.Clear;
  //                                             YkLis读取结果后设置为[已读取]    防呆,仅读取近7天结果     仅读取不为空的结果          
  adotemp22.SQL.Text:='select * from v_cm_result where isnull(staut,'''')='''' and repdate>GETDATE()-7 and isnull(result,'''')<>'''' ';
  adotemp22.Open;
  while not adotemp22.Eof do
  begin
    if '1'<>ScalarSQLCmd(ConnectString,'select 1 from chk_con where unid='+adotemp22.fieldbyname('TJH').AsString) then
    begin
      memo1.Lines.Add('chk_con.unid找不到:'+adotemp22.fieldbyname('TJH').AsString+'.barcode:'+adotemp22.fieldbyname('barcode').AsString+',itemcode:'+adotemp22.fieldbyname('itemcode').AsString+',hycode:'+adotemp22.fieldbyname('hycode').AsString);
      adotemp22.Next;
      continue;
    end;

    checkid:=ScalarSQLCmd(ConnectString,'select checkid from chk_con where unid='+adotemp22.fieldbyname('TJH').AsString);

    if(trim(checkid)<>'')and(uppercase(checkid)<>uppercase(EquipChar+adotemp22.fieldbyname('TJH').AsString))then
    begin
      memo1.Lines.Add('chk_con.unid联机号错误:'+adotemp22.fieldbyname('TJH').AsString);
      adotemp22.Next;
      continue;
    end;

    if trim(checkid)='' then
      //表示是LIS传给华银的申请单,并且联机号为空,自动生成联机号,以便Data2LisSvr.dll使用
      ExecSQLCmd(ConnectString,'update chk_con set checkid='''+EquipChar+'''+CONVERT(varchar(20),unid) where unid='+adotemp22.fieldbyname('TJH').AsString);

    check_date:=ScalarSQLCmd(ConnectString,'select CONVERT(CHAR(10),check_date,121) from chk_con where unid='+adotemp22.fieldbyname('TJH').AsString);

    ReceiveItemInfo:=VarArrayCreate([0,0],varVariant);
    ReceiveItemInfo[0]:=VarArrayof([adotemp22.FieldByName('itemcode').AsString,adotemp22.FieldByName('result').AsString,'','']);

    if bRegister then
    begin
      FInts :=CreateOleObject('Data2LisSvr.Data2Lis');
      FInts.fData2Lis(ReceiveItemInfo,adotemp22.fieldbyname('TJH').AsString,
        check_date,
        (''),adotemp22.fieldbyname('sampleTypeName').AsString,(SpecStatus),(EquipChar),
        (''),'',
        (LisFormCaption),(ConnectString),
        (''),(''),(''),'',
        ifRecLog,true,'常规',
        '',
        -1,
        '','','','',
        -1,-1,-1,-1,
        -1,-1,-1,-1,
        false,false,false,false);
      if not VarIsEmpty(FInts) then FInts:= unAssigned;
      
      //v_cm_result的联合主键字段:barcode+itemcode+hycode
      ExecSQLCmd(DaanConnStr,'update v_cm_result set staut=''已读取'' where barcode='''+adotemp22.fieldbyname('barcode').AsString+''' and itemcode='''+adotemp22.fieldbyname('itemcode').AsString+''' and hycode='''+adotemp22.fieldbyname('hycode').AsString+''' ');
      memo1.Lines.Add('获取结果.unid:'+adotemp22.fieldbyname('TJH').AsString+',name:'+adotemp22.fieldbyname('name').AsString+',barcode:'+adotemp22.fieldbyname('barcode').AsString+',itemcode:'+adotemp22.fieldbyname('itemcode').AsString+',hycode:'+adotemp22.fieldbyname('hycode').AsString);
    end;

    adotemp22.Next;
  end;
  adotemp22.Free;
  
  (Sender as TTimer).Enabled:=true;
end;

initialization
    hnd := CreateMutex(nil, True, Pchar(ExtractFileName(Application.ExeName)));
    if GetLastError = ERROR_ALREADY_EXISTS then
    begin
        MessageBox(application.Handle,pchar('该程序已在运行中！'),
                    '系统提示',MB_OK+MB_ICONinformation);   
        Halt;
    end;

finalization
    if hnd <> 0 then CloseHandle(hnd);

end.




        

