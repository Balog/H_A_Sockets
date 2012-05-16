//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Main.h"
#include "inifiles.hpp";
#include "CodeText.h"
#include "MasterPointer.h"
#include "PassForm.h"
#include "Zastavka.h"
#include "EditLogin.h"
#include "Progress.h"
#include "About.h"
using namespace std;
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
Reg=false;
/*
Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Admin.ini");

int AbsPathAspect=Ini->ReadInteger("Main","AbsPathAspect",1);
String AspectBase=Ini->ReadString("Main","AdminBase","");
String DiaryBase=Ini->ReadString("Main","DiaryBase","");
if(AbsPathAspect==0)
{
AdminDatabase=Path+AspectBase;
DiaryDatabase=Path+DiaryBase;
}
else
{
AdminDatabase= AspectBase;
DiaryDatabase=DiaryBase;
}
*/
/*
Database=new MDBConnector(ExtractFilePath(AdminDatabase), ExtractFileName(AdminDatabase), this);
Database->SetPatchBackUp("Archive");

int Days=Ini->ReadInteger("Main","StoreArchive",0);
Database->ClearArchive(Days);
Database->PackDB();
Database->Backup("Archive");

Diary=new MDBConnector(ExtractFilePath(DiaryDatabase), ExtractFileName(DiaryDatabase), this);
Diary->SetPatchBackUp("Archive");


Diary->ClearArchive(Days);
Diary->PackDB();
Diary->Backup("Archive");


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

*/



}
//---------------------------------------------------------------------------












void __fastcall TForm1::FormCreate(TObject *Sender)
{
/*
Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Admin.ini");
String Server=Ini->ReadString("Main","Server","localhost");
int Port=Ini->ReadInteger("Main","Port",2000);
MClient=new Client(ClientSocket, ActionManager1, this);

MClient->VDB.clear();
CBDatabase->Clear();
ServerName=Ini->ReadString("Server","ServerName","");
int Num=Ini->ReadInteger("Server","NumServerBase",0);
for(int i=0;i<Num;i++)
{
String S="Base"+IntToStr(i+1);
String ServerBase=Ini->ReadString(S,"ServerDatabase","");
String Name=Ini->ReadString(S,"Name","");

//MClient->AddDatabase(ServerBase);

CDBItem DBI;
DBI.Name=Name;
DBI.ServerDB=ServerBase;
DBI.Num=i;
DBI.NumDatabase=Ini->ReadInteger(S,"NumDatabase",-1);

if(DBI.NumDatabase>=0)
{
CBDatabase->Items->Add(Name);
}

MClient->VDB.push_back(DBI);
}
CBDatabase->ItemIndex=0;


int AbsPathAspect=Ini->ReadInteger("Main","AbsPathAspect",1);
String AspectBase=Ini->ReadString("Main","AdminBase","");
String DiaryBase=Ini->ReadString("Main","DiaryBase","");
if(AbsPathAspect==0)
{
AdminDatabase=Path+AspectBase;
DiaryDatabase=Path+DiaryBase;
}
else
{
AdminDatabase= AspectBase;
DiaryDatabase=DiaryBase;
}

MClient->Database=new MDBConnector(ExtractFilePath(AdminDatabase), ExtractFileName(AdminDatabase), this);
MClient->Database->SetPatchBackUp("Archive");

int Days=Ini->ReadInteger("Main","StoreArchive",0);
MClient->Database->ClearArchive(Days);
MClient->Database->PackDB();
MClient->Database->Backup("Archive");

MClient->Diary=new MDBConnector(ExtractFilePath(DiaryDatabase), ExtractFileName(DiaryDatabase), this);
MClient->Diary->SetPatchBackUp("Archive");


MClient->Diary->ClearArchive(Days);
MClient->Diary->PackDB();
MClient->Diary->Backup("Archive");

MP<TADODataSet>Roles(this);
Roles->CommandText="Select * from Roles order by Num";
Roles->Connection=MClient->Database;
Roles->Active=true;
Role->Clear();
for(Roles->First();!Roles->Eof;Roles->Next())
{
Role->Items->Add(Roles->FieldByName("Name")->AsString);
}
Role->ItemIndex=0;


MClient->Connect(Server, Port);
 */
}
//---------------------------------------------------------------------------


void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{

switch (Application->MessageBoxA("Записать все изменения на сервер?","Выход из программы",MB_YESNOCANCEL	+MB_ICONQUESTION+MB_DEFBUTTON1))
{
 case IDYES:
 {
//MClient->Start();
//Main->MClient->WriteDiaryEvent("AdminARM","Завершение работы запись данных","База данных: "+CBDatabase->Text);
//Zast->Close();
Zast->Stop=true;
Zast->PrepareSaveLogins->Execute();
 Action=caNone;
//MClient->Stop();
 break;
 }
 case IDNO:
 {
//Main->MClient->WriteDiaryEvent("AdminARM","Завершение работы отказ от записи","База данных: "+CBDatabase->Text);
Zast->MClient->WriteDiaryEvent("AdminARM","Завершение работы отказ от записи","База данных: "+CBDatabase->Text);
Zast->Close();
 Action=caFree;
 break;
 }
 case IDCANCEL:
 {
 Action=caNone;
 break;
 }
}


}
//---------------------------------------------------------------------------

void __fastcall TForm1::RoleChange(TObject *Sender)
{
UpdateTempLogin();
Users->ItemIndex=0;
UpdateOtdel(0);
}
//---------------------------------------------------------------------------
void TForm1::UpdateTempLogin()
{
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
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
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(NumLogin);
int N1=Tab->FieldByName("Num")->AsInteger;

//int NumLogin=TempLogin->FieldByName("Num")->Value;
MP<TADODataSet>Verify(this);
Verify->Connection=Zast->MClient->Database;

Verify->Active=false;
Verify->CommandText="SELECT ObslOtdel.Login, Подразделения.[Название подразделения] FROM Подразделения INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel  where Login="+IntToStr(N1)+" Order by NumObslOtdel";
//Verify->CommandText="SELECT ObslOtdel.Login, Подразделения.[Название подразделения], Подразделения.NumDatabase, ObslOtdel.NumObslOtdel FROM Подразделения INNER JOIN ObslOtdel ON (Подразделения.NumDatabase = ObslOtdel.NumDatabase) AND (Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) WHERE Подразделения.NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" ORDER BY ObslOtdel.NumObslOtdel;";
Verify->Active=true;
Otdels->Clear();
for(Verify->First();!Verify->Eof;Verify->Next())
{
Otdels->Items->Add(Verify->FieldByName("Название подразделения")->AsString);
}
}
}
//----------------------------------------------------------------

void __fastcall TForm1::FormShow(TObject *Sender)
{
Zast->UpdateLogins->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CBDatabaseClick(TObject *Sender)
{
UpdateTempLogin();
Form1->Users->ItemIndex=0;
UpdateOtdel(0);
/*
//Main->MClient->WriteDiaryEvent("AdminARM","Переключение базы данных","База данных: "+CBDatabase->Text);
//LicCount=Zast->MClient->GetLicenseCount(CBDatabase->Text);
Reg=LicCount!=0;
//MClient->SetDatabaseConnect(MClient->IDC(), CBDatabase->ItemIndex, true);
try
{
//Zast->MClient->
//VerifyLicense(CBDatabase->Text);

//UpdateOtdels();
Zast->UpdateLogins->Execute();

Zast->MClient->VTrigger.clear();
Trig T;
T.Var=CBDatabase->ItemIndex;
T.Max=Zast->MClient->VDB.size();
T.TrueAction="";
T.FalseAction="";
Zast->MClient->VTrigger.push_back(T);
//SaveLogins->Execute();
Zast->MClient->ActTrigger(0);

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("UpdateOtdels");

 Zast->MClient->ReadTable(CBDatabase->Text,"Select [Номер подразделения], [Название подразделения] from Подразделения", "Select [Номер подразделения], [Название подразделения] from TempПодразделения");


//Main->MClient->WriteDiaryEvent("AdminARM","Переключение базы данных завершено","База данных: "+CBDatabase->Text);
}
catch(...)
{
//Main->MClient->WriteDiaryEvent("AdminARM ошибка","Ошибка переключения базы данных","База данных: "+CBDatabase->Text);
}
*/
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

void __fastcall TForm1::N3Click(TObject *Sender)
{
//Удалить логин
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * From Logins Where Role="+IntToStr(Role->ItemIndex+1)+"  AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(Users->ItemIndex);

Tab->Delete();
UpdateTempLogin();

int N=Users->ItemIndex;

UpdateOtdel(N);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N1Click(TObject *Sender)
{
//Добавить логин
EditLogins->Log->Text="";
EditLogins->Pass1->Text="";
EditLogins->Pass2->Text="";

EditLogins->Mode=false;
EditLogins->NumLogin=-1;

EditLogins->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N2Click(TObject *Sender)
{
//Редактировать логин
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * from Logins Where Role="+IntToStr(Form1->Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->RecNo=Users->ItemIndex+1;


EditLogins->Log->Text=Tab->FieldByName("Login")->AsString;
EditLogins->Pass1->Text="";
EditLogins->Pass2->Text="";

EditLogins->Mode=false;
EditLogins->NumLogin=Users->ItemIndex;

EditLogins->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N4Click(TObject *Sender)
{
//Редактировать пароль
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * from Logins Where Role="+IntToStr(Form1->Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->RecNo=Users->ItemIndex+1;


EditLogins->Log->Text=Tab->FieldByName("Login")->AsString;
EditLogins->Pass1->Text="";
EditLogins->Pass2->Text="";

EditLogins->Mode=true;
EditLogins->NumLogin=Users->ItemIndex;

EditLogins->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::OtdelsMouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
FreeOtdel->Items->Clear();
if(Users->ItemIndex>=0)
{
int N=Otdels->ItemAtPos(Point(X,Y),true);
Otdels->ItemIndex=N;


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
Tab->Connection=Zast->MClient->Database;
//Tab->CommandText="SELECT Подразделения.[Номер подразделения], Подразделения.[Название подразделения], ObslOtdel.NumObslOtdel FROM Подразделения LEFT JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel WHERE (((ObslOtdel.NumObslOtdel) Is Null));";
Tab->CommandText="SELECT Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Подразделения.Del, Подразделения.NumDatabase FROM Подразделения LEFT JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.[NumObslOtdel] WHERE (((Подразделения.NumDatabase)="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+") AND ((ObslOtdel.NumObslOtdel) Is Null));";
Tab->Active=true;

for(Tab->First();!Tab->Eof;Tab->Next())
{
TMenuItem *M=new TMenuItem(FreeOtdel);
M->Caption=Tab->FieldByName("Название подразделения")->AsString;
M->OnClick=SelectOtdel;
M->Tag=Tab->FieldByName("Номер подразделения")->AsInteger;
FreeOtdel->Items->Add(M);
}

if(Role->ItemIndex>=2)
{
if(Tab->RecordCount==0)
{
TMenuItem *M=new TMenuItem(FreeOtdel);
M=new TMenuItem(FreeOtdel);
M->Caption="Нет свободных подразделений";
M->Tag=0;
M->Enabled=false;
FreeOtdel->Items->Add(M);
}
}
}
else
{
TMenuItem *M=new TMenuItem(FreeOtdel);
M=new TMenuItem(FreeOtdel);
M->Caption="Выберите логин";
M->Tag=0;
M->Enabled=false;
FreeOtdel->Items->Add(M);
}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::SelectOtdel(TObject *Sender)
{

TMenuItem *Menu=(TMenuItem*)Sender;
int Num=Menu->Tag;

MP<TADODataSet>Tab1(this);
Tab1->Connection=Zast->MClient->Database;
Tab1->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab1->Active=true;
Tab1->First();
Tab1->MoveBy(Users->ItemIndex);
int N1=Tab1->FieldByName("Num")->Value;

if(Num>0)
{


MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * From ObslOtdel";
Tab->Active=true;

Tab->Insert();
Tab->FieldByName("Login")->Value=N1;
Tab->FieldByName("NumObslOtdel")->Value=Num;
Tab->FieldByName("NumDatabase")->Value=Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase;
Tab->Post();

UpdateOtdel(Users->ItemIndex);
}
else
{
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * From ObslOtdel Where Login="+IntToStr(N1)+" Order by NumObslOtdel";
Tab->Active=true;

Tab->First();
Tab->MoveBy(Otdels->ItemIndex);
Tab->Delete();

UpdateOtdel(Users->ItemIndex);
}
}
//--------------------------------------------------------------------------
void __fastcall TForm1::N5Click(TObject *Sender)
{
//Чтение подразделений, логинов и таблицы распределения

//Чтение подразделений
/*
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->MClient->Database;
Comm->CommandText="Delete * From TempПодразделения";
Comm->Execute();



 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("UpdateOtdelsMan");
 Zast->MClient->Act.NextCommand=5;

 Zast->MClient->ReadTable(CBDatabase->Text,"Select [Номер подразделения], [Название подразделения] from Подразделения", "Select [Номер подразделения], [Название подразделения] from TempПодразделения");
*/
Zast->PrepareRead->Execute();
}
//---------------------------------------------------------------------------



void __fastcall TForm1::N6Click(TObject *Sender)
{
Zast->PrepareSaveLogins->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::UsersClick(TObject *Sender)
{
int N=Users->ItemIndex;

UpdateOtdel(N);        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N7Click(TObject *Sender)
{
Prog->Show();
Prog->PB->Visible=true;
Prog->PB->Max=6;
Prog->PB->Position=0;


FDiary->LoadDiary();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N10Click(TObject *Sender)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"ADMINARM.HLP";
Application->HelpJump("IDH_РУКОВОДСТВО");        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N8Click(TObject *Sender)
{
FAbout->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N9Click(TObject *Sender)
{
this->Close();
}
//---------------------------------------------------------------------------

