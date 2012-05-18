//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "UServer.h"
#include "inifiles.hpp";
#include "MasterPointer.h"
#include "CodeText.h"
#include "EditLogin.h"
#include "MDBConnector.h"
#include "Diary.h"

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
TNotifyIconData NID;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
Block=-1;

PageControl1->ActivePageIndex=0;

Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Server.ini");
int NumBase=Ini->ReadInteger("Main","NumDatabases",0);
int DiaryStoreDays=Ini->ReadInteger("Main","DiaryStoreDays",0);
int Days=Ini->ReadInteger("Main","StoreArchive",0);
VBases.clear();
Base->Clear();
for(int i=0;i<NumBase;i++)
{
 String NameSect="Base"+IntToStr(i+1);
 BaseItem BI;

 BI.Name=Ini->ReadString(NameSect,"Name","");

int MS=Ini->ReadInteger(NameSect,"MainSpec",0);
if(MS==0)
{
BI.MainSpec=false;
 Base->Items->Add(BI.Name);

 String Lic=Ini->ReadString(NameSect,"License","");

 FILE *F;
String File=ExtractFilePath(Application->ExeName)+Lic;
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),"Не определен", "Еще не определен", "Лицензия", "Открываю файл лицензии", File);

if ((F = fopen(File.c_str(), "rt")) == NULL)
{
 //ShowMessage("Файл лицензии не удается открыть");
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),"Не определен", "Еще не определен" "Ошибка лицензии", "Файл неудалось открыть", File);

 BI.LicCount=0;
}
else
{
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),"Не определен", "Еще не определен" "Лицензия", "Файл лицензии открыт", File); //ShowMessage("Файл лицензии открыт");
 char S[256];
 fgets(S, 257, F);
 String Text=S;
 int LC=AnalizLic(Text, BI.Name);
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),"Не определен", "Еще не определен" "Лицензия", "Файл лицензии расшифрован", "DB: "+BI.Name+" NumUsers="+IntToStr(LC)); //ShowMessage("Файл лицензии открыт");

 BI.LicCount=LC;

}

}
else
{
BI.MainSpec=true;
}


int AbsPath=Ini->ReadInteger(NameSect, "AbsPath",0);

if(AbsPath==0)
{
 BI.Patch=Path+Ini->ReadString(NameSect,"Base","");
}
else
{
 BI.Patch= Ini->ReadString(NameSect,"Base","");
}
String FN=ExtractFileName(BI.Patch);
String FN1=FN.SubString(0,FN.Length()-ExtractFileExt(BI.Patch).Length());
BI.FileName=FN1;
VBases.push_back(BI);

MDBConnector* ADOConn=new MDBConnector(ExtractFilePath(BI.Patch), FN1, this);
ADOConn->SetPatchBackUp("Archive");


ADOConn->ClearArchive(Days);
ADOConn->PackDB();
ADOConn->Backup("Archive");

if(BI.Name=="Diary")
{
MP<TADOCommand>Comm(this);
Comm->Connection=ADOConn;
Comm->CommandText="DELETE Now()-[Date_Time] AS Days FROM Events WHERE (((Now()-[Date_Time])>"+IntToStr(DiaryStoreDays)+"));";
Comm->Execute();
}
delete ADOConn;
}

Base->ItemIndex=0;

String File=GetFileDatabase(Base->Text);
int Num;
//int LicCount;
for(unsigned int i=0;i<VBases.size();i++)
{
 if(VBases[i].FileName==File)
 {
  Num=i;
  //LicCount=VBases[i].LicCount;
  break;
 }
}


Database->Connected=false;
Database->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+VBases[Num].Patch+";Persist Security Info=False";
Database->LoginPrompt=false;
Database->Connected=true;



MP<TADODataSet>Roles(this);
Roles->CommandText="Select * from Roles order by Num";
Roles->Connection=Database;
Roles->Active=true;
Role->Clear();
for(Roles->First();!Roles->Eof;Roles->Next())
{
Role->Items->Add(Roles->FieldByName("Name")->AsString);
}
Role->ItemIndex=0;

UpdateTempLogin();
Users->ItemIndex=0;

String File1=GetFileDatabase("Diary");
int Num1;
for(unsigned int i=0;i<VBases.size();i++)
{
 if(VBases[i].FileName==File1)
 {
  Num1=i;
  break;
 }
}



MyClients=new mClients(Form1, VBases[Num1].Patch);
MyClients->SetDatabasePatc(ExtractFilePath(VBases[Base->ItemIndex].Patch));
//MyClients->SetDiaryBase(VBases[Num1].Patch);
//MyClients->OnCountClientChanged=ChangeCountClient;
MyClients->CountClientChanged=ChangeCountClient;

VerifyLicense();

Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Служебное", "Сервер инициирован", "");

}
//---------------------------------------------------------------------------
void __fastcall TForm1::ChangeCountClient(TObject *Sender)
{
Label1->Caption=MyClients->ClientCount();

String S="Сервер. Клиентов: "+Label1->Caption;
strcpy(NID.szTip,S.c_str());

if(Net_Error)
{
ImageList1->GetIcon(2, Application->Icon);
}
else
{
if(MyClients->ClientCount()==0)
{
ImageList1->GetIcon(0, Application->Icon);
}
else
{
ImageList1->GetIcon(1, Application->Icon);
}
}


NID.hIcon=Application->Icon->Handle;
Shell_NotifyIcon(NIM_MODIFY, &NID);

ListBox1->Clear();
for(int i=0;i<MyClients->ClientCount();i++)
{
 ListBox1->Items->Add(IntToStr(MyClients->GetIDC(i))+". "+MyClients->GetName(i));
}
}
//-------------------------------------


void __fastcall TForm1::FormDestroy(TObject *Sender)
{
delete MyClients;
Shell_NotifyIcon(NIM_DELETE, &NID);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ListBox1Click(TObject *Sender)
{
ListBox2->Clear();
 int Number=ListBox1->ItemIndex;

Client=MyClients->GetClient(Number);
 //for(unsigned int i=0; i<C->
for(int i=0;i<Client->FormCount();i++)
{
mForm *F=Client->GetForm(i);
ListBox2->Items->Add(IntToStr(F->IDF())+". "+F->Name());
}
ListBox2->ItemIndex=-1;
}
//---------------------------------------------------------------------------



void __fastcall TForm1::RoleClick(TObject *Sender)
{
UpdateTempLogin();
Users->ItemIndex=0;
UpdateOtdel(0);
}
//---------------------------------------------------------------------------
void TForm1::UpdateTempLogin()
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" order by Login";
Tab->Active=true;

Users->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
 Users->Items->Add(Tab->FieldByName("Login")->AsString);
}
}
//-----------------------------------------------------------------
void TForm1::UpdateOtdel(int NumLogin)
{
if(NumLogin<0)
{
Otdels->Clear();
}
else
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(NumLogin);
int N1=Tab->FieldByName("Num")->AsInteger;

//int NumLogin=TempLogin->FieldByName("Num")->Value;
MP<TADODataSet>Verify(this);
Verify->Connection=Database;

Verify->Active=false;
Verify->CommandText="SELECT ObslOtdel.Login, Подразделения.[Название подразделения] FROM Подразделения INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel  where Login="+IntToStr(N1)+" Order by NumObslOtdel";
Verify->Active=true;
Otdels->Clear();
for(Verify->First();!Verify->Eof;Verify->Next())
{
Otdels->Items->Add(Verify->FieldByName("Название подразделения")->AsString);
}
}
}
//----------------------------------------------------------------

void __fastcall TForm1::UsersClick(TObject *Sender)
{
int N=Users->ItemIndex;

UpdateOtdel(N);

}
//---------------------------------------------------------------------------

void __fastcall TForm1::UsersMouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
int N=Users->ItemAtPos(Point(X,Y), true);
Users->ItemIndex=N;
UpdateOtdel(N);
if(Button==mbRight)
{
if(Role->ItemIndex>1)
{
if(N>=0)
{
N1->Enabled=true;
N2->Enabled=true;
N3->Enabled=true;
N4->Enabled=true;
}
else
{
N1->Enabled=true;
N2->Enabled=false;
N3->Enabled=false;
N4->Enabled=false;
}
}
else
{
if(N>=0)
{
N1->Enabled=false;
N2->Enabled=true;
N3->Enabled=false;
N4->Enabled=true;
}
else
{
N1->Enabled=false;
N2->Enabled=false;
N3->Enabled=false;
N4->Enabled=false;
}
}
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N1Click(TObject *Sender)
{
//Добавить логин
LicCount=GetLicenseCount(Base->Text);
if(Role->ItemIndex==2 & Users->Items->Count>=LicCount & LicCount!=-1)
{
 ShowMessage("Достигнуто максимальное количество пользователей!");
}
else
{
MP<TADODataSet>Tab(this);
Tab->Connection=Form1->Database;
Tab->CommandText="Select * From Logins";
Tab->Active=true;

Tab->Insert();
Tab->FieldByName("Login")->Value=" ";
Tab->FieldByName("Role")->Value=3;
Tab->Post();
Tab->Active=false;
Tab->Active=true;
Tab->Last();
EditLoginNumber=Tab->FieldByName("Num")->Value;

EditLogins->Mode=1;
EditLogins->Login="";
EditLogins->ShowModal();
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N2Click(TObject *Sender)
{
//Редактировать логин

MP<TADODataSet>Tab(this);
Tab->Connection=Form1->Database;
Tab->CommandText="Select * From Logins Where Role="+IntToStr(Role->ItemIndex+1)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(Users->ItemIndex);

EditLoginNumber=Tab->FieldByName("Num")->Value;

EditLogins->Mode=2;
EditLogins->Login=Tab->FieldByName("Login")->AsString;
EditLogins->ShowModal();
}
//---------------------------------------------------------------------------


void __fastcall TForm1::N4Click(TObject *Sender)
{
//Изменить пароль
MP<TADODataSet>Tab(this);
Tab->Connection=Form1->Database;
Tab->CommandText="Select * From Logins Where Role="+IntToStr(Role->ItemIndex+1)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(Users->ItemIndex);

EditLoginNumber=Tab->FieldByName("Num")->Value;

EditLogins->Mode=3;
EditLogins->Login=Tab->FieldByName("Login")->AsString;
EditLogins->ShowModal();
}
//---------------------------------------------------------------------------
void TForm1::SaveCode(String Login, String Password)
{
if(Role->ItemIndex==0)
{

Database->Connected=false;
String CS=Database->ConnectionString;
Database->Connected=true;

for(unsigned int i=0;i<VBases.size();i++)
{
if(VBases[i].MainSpec==0)
{
Database->Connected=false;
Database->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+VBases[i].Patch+";Persist Security Info=False";
Database->LoginPrompt=false;
Database->Connected=true;

MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From Logins Where Role<>1 AND Login='"+Login+"'";
Tab->Active=true;
if(Tab->RecordCount==0)
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;


Sleep(500);

CodeText *CT=new CodeText();
String T=Login+Login;
String CodeSTR=CT->Crypt(T, Password);
delete CT;

Tab->CommandText="Select * From Logins Where Role=1";
Tab->Active=true;

Tab->Edit();
Tab->FieldByName("Login")->Value=Login;
Tab->FieldByName("Code1")->Value=CodeSTR.SubString(1,128);
Tab->FieldByName("Code2")->Value=CodeSTR.SubString(129,128);
Tab->Post();
}
else
{
MP<TADOCommand>Comm(this);
Comm->Connection=Database;
Comm->CommandText="Delete * from Logins where Login=' '";
Comm->Execute();

ShowMessage("Такой логин уже зарегистрирован!");
}
}
}
Database->Connected=false;
Database->ConnectionString=CS;
Database->Connected=true;
UpdateTempLogin();

}
else
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From Logins Where Num<>"+IntToStr(EditLoginNumber)+" AND Login='"+Login+"'";
Tab->Active=true;
if(Tab->RecordCount==0)
{
Tab->Active=false;

CodeText *CT=new CodeText();
String T=Login+Login;
String CodeSTR=CT->Crypt(T, Password);
delete CT;


Tab->CommandText="Select * From Logins Where Num="+IntToStr(EditLoginNumber);
Tab->Active=true;
//Tab->MoveBy(Users->ItemIndex);

Tab->Edit();
Tab->FieldByName("Login")->Value=Login;
Tab->FieldByName("Code1")->Value=CodeSTR.SubString(1,128);
Tab->FieldByName("Code2")->Value=CodeSTR.SubString(129,128);
Tab->Post();

UpdateTempLogin();
}
else
{
MP<TADOCommand>Comm(this);
Comm->Connection=Database;
Comm->CommandText="Delete * from Logins where Login=' '";
Comm->Execute();

ShowMessage("Такой логин уже зарегистрирован!");
}
}
}
//------------------------------------------------------
void __fastcall TForm1::N3Click(TObject *Sender)
{
//Удалить логин
MP<TADODataSet>Tab(this);
Tab->Connection=Form1->Database;
Tab->CommandText="Select * From Logins Where Role=3 order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(Users->ItemIndex);

Tab->Delete();
UpdateTempLogin();
}
//---------------------------------------------------------------------------


void __fastcall TForm1::Button1Click(TObject *Sender)
{
MP<TADODataSet>Tab(this);
Tab->Connection=Form1->Database;
Tab->CommandText="Select * From Logins Where Role="+IntToStr(Role->ItemIndex+1)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(Users->ItemIndex);

String Code=Tab->FieldByName("Code1")->AsString+Tab->FieldByName("Code2")->AsString;
String TabLogin=Tab->FieldByName("Login")->Value;
CodeText* CT=new CodeText();
String Str=CT->EnCrypt(Code, Edit1->Text, 256);
delete CT;
String TLogin=Str.SubString(256-TabLogin.Length()*2+1,TabLogin.Length()*2);
String LG=TLogin.SubString(1,TLogin.Length()/2);
if(LG==TabLogin)
{
ShowMessage("Пароль верен");
}
else
{
ShowMessage("Пароль ошибочен");
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::OtdelsMouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
int N=Otdels->ItemAtPos(Point(X,Y),true);
Otdels->ItemIndex=N;

FreeOtdel->Items->Clear();
if(N>=0)
{
TMenuItem *M=new TMenuItem(FreeOtdel);
M->Caption="Удалить";
M->OnClick=SelectOtdel;
M->Tag=-1;
FreeOtdel->Items->Add(M);

M=new TMenuItem(FreeOtdel);
M->Caption="-";
M->Tag=0;
FreeOtdel->Items->Add(M);
}

MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="SELECT Подразделения.[Номер подразделения], Подразделения.[Название подразделения], ObslOtdel.NumObslOtdel FROM Подразделения LEFT JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel WHERE (((ObslOtdel.NumObslOtdel) Is Null));";
Tab->Active=true;

for(Tab->First();!Tab->Eof;Tab->Next())
{
TMenuItem *M=new TMenuItem(FreeOtdel);
M->Caption=Tab->FieldByName("Название подразделения")->AsString;
M->OnClick=SelectOtdel;
M->Tag=Tab->FieldByName("Номер подразделения")->AsInteger;
FreeOtdel->Items->Add(M);
}

if(Role->ItemIndex==2)
{
if(Tab->RecordCount==0)
{
TMenuItem *M=new TMenuItem(FreeOtdel);
M=new TMenuItem(FreeOtdel);
M->Caption="Нет свободных отделов";
M->Tag=0;
M->Enabled=false;
FreeOtdel->Items->Add(M);
}
}

}
//---------------------------------------------------------------------------
void __fastcall TForm1::SelectOtdel(TObject *Sender)
{

TMenuItem *Menu=(TMenuItem*)Sender;
int Num=Menu->Tag;

MP<TADODataSet>Tab1(this);
Tab1->Connection=Database;
Tab1->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" order by Login";
Tab1->Active=true;
Tab1->First();
Tab1->MoveBy(Users->ItemIndex);
int N1=Tab1->FieldByName("Num")->Value;

if(Num>0)
{


MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From ObslOtdel";
Tab->Active=true;

Tab->Insert();
Tab->FieldByName("Login")->Value=N1;
Tab->FieldByName("NumObslOtdel")->Value=Num;
Tab->Post();

UpdateOtdel(Users->ItemIndex);
}
else
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From ObslOtdel Where Login="+IntToStr(N1)+" Order by NumObslOtdel";
//Tab->CommandText="SELECT ObslOtdel.Login, Подразделения.[Название подразделения] FROM Подразделения INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel  where Login="+IntToStr(N1);
Tab->Active=true;

Tab->First();
Tab->MoveBy(Otdels->ItemIndex);
//ShowMessage(Tab->FieldByName("Login")->AsString+" "+Tab->FieldByName("NumObslOtdel")->AsString);
Tab->Delete();

UpdateOtdel(Users->ItemIndex);
}
}
//--------------------------------------------------------------------------
bool TForm1::MergeLogins()
{
MP<TADOCommand>Comm(this);
Comm->Connection=Database;
Comm->CommandText="Delete * from Logins where Del=true";


MP<TADODataSet>TempLogins(this);
TempLogins->Connection=Database;
TempLogins->CommandText="Select * From TempLogins";
TempLogins->Active=true;

MP<TADODataSet>Logins(this);
Logins->Connection=Database;
Logins->CommandText="Select * From Logins Order by Num";
Logins->Active=true;

MP<TADODataSet>TempObslOtdel(this);
TempObslOtdel->Connection=Database;

//Пройти по Logins, переписать данные совпадающих логинов,удаляя строки из TempLogins
//пометить к удалению те, у которых нет совпадений в TempLogins и удалить их в конце

for(Logins->First();!Logins->Eof;Logins->Next())
{
Logins->Edit();
Logins->FieldByName("Del")->Value=false;
Logins->Post();

String Log=Logins->FieldByName("Login")->Value;
bool B=TempLogins->Locate("Login",Log,SO);
if(B)
{
//есть совпадение
int NumLogin=Logins->FieldByName("Num")->Value;
int NumTempLogin= TempLogins->FieldByName("Num")->Value;

TempObslOtdel->Active=false;
TempObslOtdel->CommandText="Select * From TempObslOtdel where Login="+IntToStr(NumTempLogin);
TempObslOtdel->Active=true;
for(TempObslOtdel->First();!TempObslOtdel->Eof;TempObslOtdel->Next())
{
TempObslOtdel->Edit();
TempObslOtdel->FieldByName("Login")->Value=NumLogin;
TempObslOtdel->Post();
}

Logins->Edit();
Logins->FieldByName("Code1")->Value=TempLogins->FieldByName("Code1")->Value;
Logins->FieldByName("Code2")->Value=TempLogins->FieldByName("Code2")->Value;
Logins->FieldByName("Role")->Value=TempLogins->FieldByName("Role")->Value;
Logins->Post();
TempLogins->Delete();
}
else
{
Logins->Edit();
Logins->FieldByName("Del")->Value=true;
Logins->Post();
}
}
Comm->Execute();
//пройти по TempLogins и перенести в Logins оставшиеся
for(TempLogins->First();!TempLogins->Eof;TempLogins->Next())
{
Logins->Insert();
Logins->FieldByName("Login")->Value=TempLogins->FieldByName("Login")->Value;
Logins->FieldByName("Code1")->Value=TempLogins->FieldByName("Code1")->Value;
Logins->FieldByName("Code2")->Value=TempLogins->FieldByName("Code2")->Value;
Logins->FieldByName("Role")->Value=TempLogins->FieldByName("Role")->Value;
Logins->Post();

Logins->Active=false;
Logins->Active=true;
Logins->Last();

int NumLogin=Logins->FieldByName("Num")->Value;
int NumTempLogin= TempLogins->FieldByName("Num")->Value;

TempObslOtdel->Active=false;
TempObslOtdel->CommandText="Select * From TempObslOtdel where Login="+IntToStr(NumTempLogin);
TempObslOtdel->Active=true;
for(TempObslOtdel->First();!TempObslOtdel->Eof;TempObslOtdel->Next())
{
TempObslOtdel->Edit();
TempObslOtdel->FieldByName("Login")->Value=NumLogin;
TempObslOtdel->Post();
}
}

//очистить TempLogins
Comm->CommandText="Delete * from TempLogins";
Comm->Execute();



return true;
}
//--------------------------------------------------------------------------

void __fastcall TForm1::FormCreate(TObject *Sender)
{
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Служебное", "Форма сервера создана", "");

Application->ShowMainForm=false;
 NID.cbSize = sizeof(TNotifyIconData);
NID.hWnd=Handle;
NID.uID=1;
NID.uFlags=NIF_ICON | NIF_MESSAGE | NIF_TIP;
NID.uCallbackMessage=MyTrayIcon;


ImageList1->GetIcon(0, Application->Icon);

NID.hIcon=Application->Icon->Handle;

strcpy(NID.szTip,"Сервер");
Shell_NotifyIcon(NIM_ADD, &NID);


}
//---------------------------------------------------------------------------
void __fastcall TForm1::MTIcon(TMessage&a)
{
POINT P;
switch( a.LParam)
{
case 514:

 Show();
 SetForegroundWindow(Handle);
 break;

case 516:
 GetCursorPos(&P);
 PopupMenu2->Popup(P.x,P.y);

}
}
//-------------------------------------------------------------------------
void __fastcall TForm1::N5Click(TObject *Sender)
{
int N=MyClients->ClientCount();
if(N==0)
{
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Служебное", "Ручной останов сервера", "");

this->Close();
}
else
{
ShowMessage("К серверу подключено "+IntToStr(N)+" клиентов!\rПодождите пока они отключатся.");
}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::WMSysCommand(TMessage & Msg)
{
  switch (Msg.WParam)
  {

case SC_MINIMIZE:
{
Application->Minimize();
break;
}

case SC_CLOSE:
{

  Shell_NotifyIcon(NIM_ADD, &NID);
//  BCreate->Enabled = false;
// BDelete->Enabled = true;
// BHide->Enabled = true;
Form1->Visible=false;
//SendMessage(Handle,WM_SYSCOMMAND,SC_MINIMIZE,0);

//ShowMessage("SC_Close");
break;
}

default:
DefWindowProc(Handle,Msg.Msg,Msg.WParam,Msg.LParam);

 }

}
//------------------------------------------------------------------
void TForm1::Error()
{
ImageList1->GetIcon(2, Application->Icon);
NID.hIcon=Application->Icon->Handle;
Shell_NotifyIcon(NIM_MODIFY, &NID);
}
//-------------------------------------------------------------

void __fastcall TForm1::BaseClick(TObject *Sender)
{
int N=Users->ItemIndex;

Database->Connected=false;
Database->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+VBases[Base->ItemIndex].Patch+";Persist Security Info=False";
Database->LoginPrompt=false;
Database->Connected=true;

UpdateTempLogin();

Users->ItemIndex=N;

UpdateOtdel(N);
}
//---------------------------------------------------------------------------
String TForm1::GetFileDatabase(String NameDatabase)
{/*
long Num=0;
for(unsigned int i=0;i<VBases.size();i++)
{
 if(VBases[i].Name==NameDatabase)
 {
  Num=i;
 }
}
return Num;
*/
String Ret="";
for(unsigned int i=0;i<VBases.size();i++)
{
 if(VBases[i].Name==NameDatabase)
 {
  Ret=VBases[i].FileName;
 }
}

return Ret;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::N6Click(TObject *Sender)
{
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Служебное", "Ручной останов сервера", "");

this->Close();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::N7Click(TObject *Sender)
{
FDiary->ShowModal();
}
//---------------------------------------------------------------------------
int TForm1::AnalizLic(String Text, String Pass)
{
int Ret=0;
CodeText *CT=new CodeText();
String DText=CT->EnCrypt(Text,Pass, 256);
String DCopy=DText;
int Pos1=DText.AnsiPos("+");
String Key1=DText.SubString(1,Pos1-1);
DText=DText.SubString(Pos1+1,DText.Length());
int Pos2=DText.AnsiPos("+");
String Key2=DText.SubString(1,Pos2-1);

String Key1_2="+"+Key1+"+"+Key2;
int Pos3=DCopy.AnsiPos(Key1_2);

//ShowMessage("Pos3="+IntToStr(Pos3));
if(Pos3!=0)
{
//Лицензия верна
if(Key1=="Неограниченая" & Key2=="версия")
{
//Неограниченая лицензия
Ret=-1;
}
else
{
//Ограниченая лицензия
Ret=ReadRes(Key1, Key2);
}
}
else
{
Ret=0;
}
//ShowMessage("Ret="+IntToStr(Ret));
return Ret;
}
//---------------------------------------------------------------------------
int TForm1::ReadRes(String Key1, String Key2)
{
 int N1=Propis(Key1);
 int N2=Propis(Key2);
 int Res=N1*10+N2;
 return Res;
}
//----------------------------------------------------------------------------
int TForm1::Propis(String S)
{
int Res;
if(S=="Нуль")
{
  Res=0;
  return Res;
}
if(S=="Один")
{
  Res=1;
  return Res;
}
if(S=="Два")
{
  Res=2;
  return Res;
}
if(S=="Три")
{
  Res=3;
  return Res;
}
if(S=="Четыре")
{
  Res=4;
  return Res;
}
if(S=="Пять")
{
  Res=5;
  return Res;
}
if(S=="Шесть")
{
  Res=6;
  return Res;
}
if(S=="Семь")
{
  Res=7;
  return Res;
}
if(S=="Восемь")
{
  Res=8;
  return Res;
}
if(S=="Девять")
{
  Res=9;
  return Res;
}

return Res;
}
//------------------------------------------------------------------------
void TForm1::VerifyLicense()
{
//int Num;
//int LicCount;
MP<TADOConnection>DB(this);
DB->LoginPrompt=false;
for(unsigned int i=0;i<VBases.size();i++)
{
 if(VBases[i].MainSpec==0)
 {
  //Num=i;
  LicCount=VBases[i].LicCount;

  DB->Connected=false;
  DB->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+VBases[i].Patch+";Persist Security Info=False";
  DB->Connected=true;

if(LicCount>=0)
{
MP<TADODataSet>Logins(this);
Logins->Connection=DB;
Logins->CommandText="Select * From Logins Where Role=3 Order by Num";
Logins->Active=true;
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Лицензия", "Проверка числа пользователей", "DB: "+VBases[i].Name+" Разрешено="+IntToStr(LicCount)+" Имеется="+IntToStr(Logins->RecordCount)); //ShowMessage("Файл лицензии открыт");
if(LicCount<Logins->RecordCount)
{
int j=0;
for(Logins->First();!Logins->Eof;Logins->Next())
{
//ShowMessage("i="+IntToStr(i)+" LicCount="+IntToStr(LicCount));
if(j>=LicCount)
{
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Лицензия ошибка", "Удаление избыточных пользователей", "DB: "+VBases[i].Name+" Login: "+Logins->FieldByName("Login")->AsString); //ShowMessage("Файл лицензии открыт");

//Logins->Delete();
Logins->Edit();
Logins->FieldByName("Del")->Value=true;
Logins->Post();
}
j++;
}

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="Delete * From Logins Where Del=true";
Comm->Execute();
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Лицензия ошибка", "Удаление избыточных пользователей", "DB: "+VBases[i].Name+" Удалено пользователей - "+IntToStr(i-LicCount)); //ShowMessage("Файл лицензии открыт");

}
else
{
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Лицензия", "Проверка числа пользователей", "DB: "+VBases[i].Name+" Разрешенное число пользователей"); //ShowMessage("Файл лицензии открыт");

}
}
else
{
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Лицензия", "Неограниченная лицензия", "DB: "+VBases[i].Name); //ShowMessage("Файл лицензии открыт");

}
 }
}

/*
String File=GetFileDatabase(Base->Text);
int Num;
int LicCount;
for(unsigned int i=0;i<VBases.size();i++)
{
 if(VBases[i].FileName==File)
 {
  Num=i;
  LicCount=VBases[i].LicCount;
  break;
 }
}

Database->Connected=false;
Database->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+VBases[Num].Patch+";Persist Security Info=False";
Database->LoginPrompt=false;
Database->Connected=true;
*/

/*
if(LicCount>=0)
{
MP<TADODataSet>Logins(this);
Logins->Connection=Database;
Logins->CommandText="Select * From Logins Where Role=3 Order by Num";
Logins->Active=true;
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Лицензия", "Проверка числа пользователей", "DB: "+Name+" Разрешено="+IntToStr(LicCount)+" Имеется="+IntToStr(Logins->RecordCount)); //ShowMessage("Файл лицензии открыт");
if(LicCount<Logins->RecordCount)
{
int i=0;
for(Logins->First();!Logins->Eof;Logins->Next())
{
//ShowMessage("i="+IntToStr(i)+" LicCount="+IntToStr(LicCount));
if(i>=LicCount)
{
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Лицензия ошибка", "Удаление избыточных пользователей", "DB: "+Name+" Login: "+Logins->FieldByName("Login")->AsString); //ShowMessage("Файл лицензии открыт");

//Logins->Delete();
Logins->Edit();
Logins->FieldByName("Del")->Value=true;
Logins->Post();
}
i++;
}

MP<TADOCommand>Comm(this);
Comm->Connection=Database;
Comm->CommandText="Delete * From Logins Where Del=true";
Comm->Execute();
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Лицензия ошибка", "Удаление избыточных пользователей", "DB: "+Name+" Удалено пользователей - "+IntToStr(i-LicCount)); //ShowMessage("Файл лицензии открыт");

}
else
{
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Лицензия", "Проверка числа пользователей", "DB: "+Name+" Разрешенное число пользователей"); //ShowMessage("Файл лицензии открыт");

}
}
else
{
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Лицензия", "Неограниченная лицензия", "DB: "+Name); //ShowMessage("Файл лицензии открыт");

}
*/
}
//------------------------------------------------------------------------
long TForm1::GetLicenseCount(String DBName)
{
int Ret=0;
for(unsigned int i=0;i<VBases.size();i++)
{
 if(VBases[i].Name==DBName)
 {
  Ret=VBases[i].LicCount;
  break;
 }
}
return Ret;
}
//-------------------------------------------------------------------------
