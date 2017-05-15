unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ComCtrls, dxdbtree, DBAccess, MSAccess,
  dxdbtrel, SdacVcl, StdCtrls, ExtCtrls, dxtree, MemDS, GridsEh, DBGridEh,
  DBCtrls, cxGraphics, cxControls, Mask, Buttons, ExtDlgs;

type
  TForm1 = class(TForm)
    conMS: TMSConnection;
    mscnctdlg: TMSConnectDialog;
    MSGroup: TMSQuery;
    grp1: TGroupBox;
    grp2: TGroupBox;
    grp3: TGroupBox;
    tvDBTreeView: TdxDBTreeView;
    pnl1: TPanel;
    MSDSGroup: TMSDataSource;
    MSUser: TMSQuery;
    MSDSUser: TMSDataSource;
    MSUserGroup: TMSQuery;
    MSDSUserGroup: TMSDataSource;
    pnl3: TPanel;
    gUsers: TDBGridEh;
    dbName: TDBEdit;
    DBFhone: TDBEdit;
    DBAddress: TDBEdit;
    BitBtn1: TBitBtn;
    DBImage1: TDBImage;
    dlgOpenPic: TOpenPictureDialog;
    AccountEnabled: TDBCheckBox;
    l3: TLabel;
    l4: TLabel;
    l5: TLabel;
    DBNote: TDBMemo;
    l6: TLabel;
    BitBtnAdd: TBitBtn;
    BitBtnAddChild: TBitBtn;
    BitBtnDel: TBitBtn;
    MSQuery1: TMSQuery;
    l1: TLabel;
    EditGrp: TEdit;
    BitBtnSearch: TBitBtn;
    l2: TLabel;
    EditUser: TEdit;
    BitBtnSearchUser: TBitBtn;
    pnl2: TPanel;
    l7: TLabel;
    dbGoup: TdxDBLookupTreeView;
    MSGroupTree: TMSQuery;
    MSDSGroupTree: TMSDataSource;
    MSUserTree: TMSQuery;
    MSDSUserTree: TMSDataSource;
    dxDBLookupTreeView1: TdxDBLookupTreeView;
    MSGroupList: TMSTable;
    MSDSGroupList: TMSDataSource;
    DBEditName: TDBEdit;
    l8: TLabel;
    l9: TLabel;
    MSGroupTree1: TMSQuery;
    MSDSGroupTree1: TMSDataSource;
    BitBtn2: TBitBtn;
    BitBtnNewUser: TBitBtn;
    BitBtnDelUser: TBitBtn;
    procedure FormShow(Sender: TObject);
    procedure gUsersGetCellParams(Sender: TObject; Column: TColumnEh;
      AFont: TFont; var Background: TColor; State: TGridDrawState);
    procedure BitBtn1Click(Sender: TObject);
    procedure MSUserAfterScroll(DataSet: TDataSet);
    procedure BitBtnDelClick(Sender: TObject);
    procedure BitBtnAddClick(Sender: TObject);
    procedure BitBtnAddChildClick(Sender: TObject);
    procedure EditGrpKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BitBtnSearchClick(Sender: TObject);
    procedure dbGoupCloseUp(Sender: TObject; Accept: Boolean);
    procedure BitBtnSearchUserClick(Sender: TObject);
    procedure EditUserKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtnNewUserClick(Sender: TObject);
    procedure BitBtnDelUserClick(Sender: TObject);
    procedure DBEditNameKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}
uses  IniFiles, jpeg;

function ReturnSql(sql:String):Variant;
begin
  //функция возвращает первое значение по запросу из базы данных
  Form1.MSQuery1.SQL.Text := sql;
  Form1.MSQuery1.Execute;
  Result :=  Form1.MSQuery1.Fields[0].AsVariant;
end;

procedure TForm1.FormShow(Sender: TObject);
var Conf : TIniFile;
    s    : string;
begin
     Conf := TIniFile.Create(ExtractFilePath(Application.ExeName)+'conf.ini');

    //Загружаем имя сервера
     S := Conf.ReadString('Connect', 'Server', '');
     if s <> '' then
      conMS.Server := s;

     //загружаем имя пользователя
     s := Conf.ReadString('Connect', 'User', '');
     if s <> '' then
      conMS.Username := s;

     //загружаем базу данных
     s := Conf.ReadString('Connect', 'DataBase', '');
     if s <> '' then
      conMS.Database := s;

     // пароль не будем сохранять :) хотя можно и его

     try
     conMS.Connect;
     except
     Application.Terminate;
     end;

     if not conMS.Connected then
      Form1.Close;
     //сохраняем имя сервера
     Conf.WriteString('Connect','Server',conMS.Server);
     //сохраняем имя пользователя
     Conf.WriteString('Connect','User',conMS.Username);
     //сохраняем базу данных
     Conf.WriteString('Connect','DataBase',conMS.Database);

     Conf.Free;

     MSGroup.Active := True;
     MSUserGroup.Active := True;
     MSUser.Active := True;
     MSUserTree.Active := True;
     MSGroupTree.Active := True;
     MSGroupTree1.Active := True;
     MSGroupList.Active := True;

end;

procedure TForm1.gUsersGetCellParams(Sender: TObject; Column: TColumnEh;
  AFont: TFont; var Background: TColor; State: TGridDrawState);
begin
    if not MSUserGroup.FieldByName('AccountEnabled').AsBoolean then
      Background :=  clCream;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
var Stream: TMemoryStream;
    Field: TBlobField;
    Jpg: TJpegImage;

begin
    if not dlgOpenPic.Execute then exit;

    Stream := TMemoryStream.Create;
    Jpg := TJpegImage.Create;
    Field := MSUser.FieldByName('photo') as TBlobField;
    MSUser.CheckBrowseMode;
    MSUser.Edit;
    Field.LoadFromFile(dlgOpenPic.FileName);
    MSUser.CheckBrowseMode;
    Field.SaveToStream(Stream);
    Stream.Seek(0, soFromBeginning);
    Jpg.LoadFromStream(Stream);
    DBImage1.Picture.Assign(jpg);
    Stream.Free;
    jpg.Free;
end;

procedure TForm1.MSUserAfterScroll(DataSet: TDataSet);
var Stream: TMemoryStream;
    Field: TBlobField;
    Jpg: TJpegImage;

begin
    DBImage1.Picture.Assign(nil);
    MsUserTree.Filter := 'UserId = 0';
    if MSUser.RecordCount = 0 then Exit;
    Stream := TMemoryStream.Create;
    Jpg := TJpegImage.Create;
    Field := MSUser.FieldByName('photo') as TBlobField;
    Field.SaveToStream(Stream);
    Stream.Seek(0, soFromBeginning);
    if Stream.Size > 0 then
    begin
     Jpg.LoadFromStream(Stream);
     DBImage1.Picture.Assign(jpg);
     Stream.Free;
     jpg.Free;
    end;
    MsUserTree.Filter := 'UserId = '+MsUser.FieldByName('UserId').AsString;
end;

procedure TForm1.BitBtnDelClick(Sender: TObject);
begin

  if (tvDBTreeView.Selected <> Nil) then
   begin
     if messagedlg('Удалить группу?',mtConfirmation,[mbYes,mbNo],0)=mrNo then Exit;
     if MSGroup.FieldByName('DelEnable').AsInteger = 0 then
      tvDBTreeView.Selected.Delete
     else
      ShowMessage('Невозможно удалить есть подгруппы или пользователи!') 

   end;
end;

procedure TForm1.BitBtnAddClick(Sender: TObject);
var idgroup:Integer;
begin
    idgroup := ReturnSql('select max(GroupId)+1 from dbo.GroupList');
    conMS.ExecSQL('insert into dbo.GroupList (DisplayName) values(:name)',['Группа '+IntToStr(idgroup)]);
    idgroup := ReturnSql('select max(GroupId) from dbo.GroupList');
    conMS.ExecSQL('insert into dbo.GroupTree (GroupId,ParentId) values(:id,:parent)',[idgroup,MSGroup.FieldByName('ParentId').AsInteger]);
    //обновляем деревья
    MSGroupTree.Active := False;
    MSGroupTree.Active := True;
    MSGroupTree1.Active := False;
    MSGroupTree1.Active := True;

    MSGroup.Active := False;
    MSGroup.Active := True;

    MSGroup.Locate('GroupId',VarArrayOf([idgroup]),[]);

    tvDBTreeView.SetFocus;


end;

procedure TForm1.BitBtnAddChildClick(Sender: TObject);
var idgroup:Integer;
begin
    idgroup := ReturnSql('select max(GroupId)+1 from dbo.GroupList');
    conMS.ExecSQL('insert into dbo.GroupList (DisplayName) values(:name)',['Группа '+IntToStr(idgroup)]);
    idgroup := ReturnSql('select max(GroupId) from dbo.GroupList');
    conMS.ExecSQL('insert into dbo.GroupTree (GroupId,ParentId) values(:id,:parent)',[idgroup,MSGroup.FieldByName('GroupId').AsInteger]);
    //обновляем деревья
    MSGroupTree.Active := False;
    MSGroupTree.Active := True;
    MSGroupTree1.Active := False;
    MSGroupTree1.Active := True;

    MSGroup.Active := False;
    MSGroup.Active := True;

    MSGroup.Locate('GroupId',VarArrayOf([idgroup]),[]);

    tvDBTreeView.SetFocus;

end;

procedure TForm1.EditGrpKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
      if (Key=13)and(EditGrp.Text<>'')  then
      begin
      //поиск группы
      BitBtnSearchClick(Sender);
      end;
end;

procedure TForm1.BitBtnSearchClick(Sender: TObject);
var idgroup:Integer;
begin
    idgroup := ReturnSql('select isnull(min(GroupId),0) from dbo.GroupList where DisplayName like ''%'+EditGrp.Text+'%''');
    MSGroup.Locate('GroupId',VarArrayOf([idgroup]),[]);
    tvDBTreeView.SetFocus;

end;

procedure TForm1.dbGoupCloseUp(Sender: TObject;
  Accept: Boolean);
var UserId:integer;
begin
    UserId := msUser.FieldByName('UserId').AsInteger;
    MSGroup.Locate('GroupId',VarArrayOf([msGroupTree.FieldByName('GroupId').AsInteger]),[]);
    tvDBTreeView.SetFocus;
    MSUserGroup.Active := false;
    MSUserGroup.Active := true;
    MSUserGroup.Locate('UserId',VarArrayOf([UserId]),[]);


end;

procedure TForm1.BitBtnSearchUserClick(Sender: TObject);
var idgroup : Integer;
    iduser  : Integer;
begin
    iduser := ReturnSql('select isnull(min(UserID),0)from dbo.UserList u where u.DisplayName like ''%'+EditUser.Text+'%''');
    if iduser > 0 then
     idgroup := ReturnSQL('select GroupId from dbo.UserTree where UserId = '+inttostr(iduser))
    else
     idgroup := 0;
    MSGroup.Locate('GroupId',VarArrayOf([idgroup]),[]);
    MSUserGroup.Locate('UserId',VarArrayOf([iduser]),[]);
    tvDBTreeView.SetFocus;

end;

procedure TForm1.EditUserKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
      if (Key=13)and(EditUser.Text<>'')  then
      begin
      //поиск пользователя
      //поиск группы
      BitBtnSearchUserClick(Sender);
      end;
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
begin
     MSGroup.CheckBrowseMode;
     MSGroup.Edit;
     MSGroup.FieldByName('ParentId').AsInteger := 0;
     MSGroup.CheckBrowseMode;
end;

procedure TForm1.BitBtnNewUserClick(Sender: TObject);
var   iduser  : Integer;
begin
    iduser := ReturnSql('select max(UserID)+1 from dbo.UserList');
    conMS.ExecSQL('insert into dbo.UserList (DisplayName,AccountEnabled) values(:name,:acc)',['Пользователь '+inttostr(iduser),1]);
    iduser := ReturnSql('select max(UserID) from dbo.UserList');
    conMS.ExecSQL('insert into dbo.UserTree (UserId,GroupId) values(:iduser,:idgroup)',[iduser,msgroup.FieldByName('GroupId').AsInteger]);

    MSUserGroup.Active := False;
    MSUserGroup.Active := True;
    MSUserGroup.Locate('UserId',VarArrayOf([iduser]),[]);
    dbName.SetFocus;
end;

procedure TForm1.BitBtnDelUserClick(Sender: TObject);
begin
     if MSUser.RecordCount = 0 then Exit;
     if messagedlg('Удалить пользователя?',mtConfirmation,[mbYes,mbNo],0)=mrNo then Exit;
     conMS.ExecSQL('delete from dbo.UserTree where UserId=:iduser',[MSUser.FieldByName('UserId').AsInteger]);
     conMS.ExecSQL('delete from dbo.UserList where UserId=:iduser',[MSUser.FieldByName('UserId').AsInteger]);
     MSUserGroup.Active := False;
     MSUserGroup.Active := True;
end;

procedure TForm1.DBEditNameKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var idgroup : Integer;
begin
     if (Key=13) then
     begin
       //изменение  имени группы
        idgroup := MSGroup.FieldByName('GroupId').AsInteger;
        MSGroupList.CheckBrowseMode;
        MSGroup.Active := False;
        MSGroup.Active := True;
        MSGroup.Locate('GroupId',VarArrayOf([idgroup]),[]);
     end;
end;

end.
