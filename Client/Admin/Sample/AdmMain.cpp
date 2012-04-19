//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "AdmMain.h"
#include "inifiles.hpp";
#include "CodeText.h"
#include "MasterPointer.h"
#include "EditLogin.h"
#include "Password.h"
#include "Diary.h"
#include "About.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TMain *Main;
//---------------------------------------------------------------------------
__fastcall TMain::TMain(TComponent* Owner)
        : TForm(Owner)
{
Reg=false;
InputPass=false;
Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Admin.ini");
int AbsPathAspect=Ini->ReadInteger("Main","AbsPathAspect",1);
String AspectBase=Ini->ReadString("Main","AspectBase","");
if(AbsPathAspect==0)
{
AdminDatabase=Path+AspectBase;
}
else
{
AdminDatabase= AspectBase;
}

Database=new MDBConnector(ExtractFilePath(AdminDatabase), ExtractFileName(AdminDatabase), this);
Database->SetPatchBackUp("Archive");


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

int Days=Ini->ReadInteger("Main","StoreArchive",0);
Database->ClearArchive(Days);
Database->PackDB();
Database->Backup("Archive");
}
//---------------------------------------------------------------------------
void TMain::UpdateTempLogin()
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;

Users->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
 Users->Items->Add(Tab->FieldByName("Login")->AsString);
}
}
//-----------------------------------------------------------------


void __fastcall TMain::RoleClick(TObject *Sender)
{
UpdateTempLogin();
Users->ItemIndex=0;
UpdateOtdel(0);
}
//---------------------------------------------------------------------------
void TMain::UpdateOtdel(int NumLogin)
{
if(NumLogin<0)
{
Otdels->Clear();
}
else
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(NumLogin);
int N1=Tab->FieldByName("Num")->AsInteger;

//int NumLogin=TempLogin->FieldByName("Num")->Value;
MP<TADODataSet>Verify(this);
Verify->Connection=Database;

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

void __fastcall TMain::UsersClick(TObject *Sender)
{
int N=Users->ItemIndex;

UpdateOtdel(N);
}
//---------------------------------------------------------------------------

void __fastcall TMain::UsersMouseDown(TObject *Sender,
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

void __fastcall TMain::OtdelsMouseDown(TObject *Sender,
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
Tab->Connection=Database;
//Tab->CommandText="SELECT Подразделения.[Номер подразделения], Подразделения.[Название подразделения], ObslOtdel.NumObslOtdel FROM Подразделения LEFT JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel WHERE (((ObslOtdel.NumObslOtdel) Is Null));";
Tab->CommandText="SELECT Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Подразделения.Del, Подразделения.NumDatabase FROM Подразделения LEFT JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.[NumObslOtdel] WHERE (((Подразделения.NumDatabase)="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+") AND ((ObslOtdel.NumObslOtdel) Is Null));";
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
void __fastcall TMain::SelectOtdel(TObject *Sender)
{

TMenuItem *Menu=(TMenuItem*)Sender;
int Num=Menu->Tag;

MP<TADODataSet>Tab1(this);
Tab1->Connection=Database;
Tab1->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" order by Login";
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
Tab->FieldByName("NumDatabase")->Value=Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase;
Tab->Post();

UpdateOtdel(Users->ItemIndex);
}
else
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From ObslOtdel Where Login="+IntToStr(N1)+" Order by NumObslOtdel";
Tab->Active=true;

Tab->First();
Tab->MoveBy(Otdels->ItemIndex);
Tab->Delete();

UpdateOtdel(Users->ItemIndex);
}
}
//--------------------------------------------------------------------------
void TMain::LoadTable(Table* FromCopy, TADODataSet *ToCopy)
{

if(ToCopy->FieldCount==FromCopy->FieldsCount())
{
ToCopy->First();
//ShowMessage(ToCopy->FieldList->Fields[0]->AsString);
for(FromCopy->First();!FromCopy->eof();FromCopy->Next())
{
ToCopy->Insert();
for(int i=0;i<FromCopy->FieldsCount();i++)
{
Variant V=FromCopy->Fields(i);

ToCopy->FieldList->Fields[i]->Value=V;

}
ToCopy->Post();
}
}
else
{
ShowMessage("Число полей копируемых таблиц не совпадает");
}
}
//---------------------------------------------------------------
void TMain::LoadTable(TADODataSet* FromCopy, Table* ToCopy)
{

if(ToCopy->FieldsCount()==FromCopy->FieldCount)
{
ToCopy->First();
//ShowMessage(ToCopy->FieldList->Fields[0]->AsString);
for(FromCopy->First();!FromCopy->Eof;FromCopy->Next())
{
ToCopy->Insert();
for(int i=0;i<FromCopy->FieldCount;i++)
{
Variant V=FromCopy->FieldList->Fields[i]->Value;

ToCopy->Fields(V,i);

}
ToCopy->Post();
}
}
else
{
ShowMessage("Число полей копируемых таблиц не совпадает");
}
}
//---------------------------------------------------------------
int TMain::VerifyTable(TADODataSet* FromCopy, Table* ToCopy)
{
        int Ret=0;
        if(ToCopy->FieldsCount()==FromCopy->FieldCount)
        {
                if(ToCopy->RecordCount()==FromCopy->RecordCount)
                {
                        ToCopy->First();
                        for(FromCopy->First();!FromCopy->Eof;FromCopy->Next())
                        {
                                for(int i=0;i<FromCopy->FieldCount;i++)
                                {
                                        Variant V=FromCopy->FieldList->Fields[i]->Value;
                                        if(ToCopy->Fields(i)!=V)
                                        {
                                                Ret=1;
                                                break;
                                        }

                                }
                                ToCopy->Next();
                                if(Ret==1)
                                {
                                        break;
                                }
                        }
                }
                else
                {
                        Ret=2;
                }
        }
        else
        {
                Ret=3;
        }

return Ret;
}
//---------------------------------------------------------------
int TMain::VerifyTable(Table* FromCopy, TADODataSet* ToCopy)
{
return VerifyTable(ToCopy, FromCopy);
}
//---------------------------------------------------------------
void TMain::MergeLogins()
{
Main->MClient->WriteDiaryEvent("AdminARM","Начало объединения логинов","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Database;
Comm->CommandText="UPDATE Logins SET Logins.Del = False;";
Comm->Execute();


MP<TADODataSet>TempLogins(this);
TempLogins->Connection=Database;
TempLogins->CommandText="Select * From TempLogins";
TempLogins->Active=true;

MP<TADODataSet>Logins(this);
Logins->Connection=Database;
Logins->CommandText="Select * From Logins where Numdatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" Order by Num";
Logins->Active=true;

MP<TADODataSet>TempObslOtdel(this);
TempObslOtdel->Connection=Database;

//Пройти по Logins, переписать данные совпадающих логинов,удаляя строки из TempLogins
//пометить к удалению те, у которых нет совпадений в TempLogins и удалить их в конце

for(Logins->First();!Logins->Eof;Logins->Next())
{
int N= Logins->FieldByName("ServerNum")->Value;

bool B=TempLogins->Locate("Num",N,SO);
if(B)
{
//есть совпадение

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
Comm->CommandText="Delete * from Logins where Del=true AND Numdatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase);
Comm->Execute();

Comm->CommandText="INSERT INTO Logins ( ServerNum, Login, Code1, Code2, Role, NumDatabase ) SELECT TempLogins.Num, TempLogins.Login, TempLogins.Code1, TempLogins.Code2, TempLogins.Role, "+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" AS [Database] FROM TempLogins;";
Comm->Execute();
//пройти по TempLogins и перенести в Logins оставшиеся
/*
for(TempLogins->First();!TempLogins->Eof;TempLogins->Next())
{
Logins->Insert();
Logins->FieldByName("Login")->Value=TempLogins->FieldByName("Login")->Value;
Logins->FieldByName("Code1")->Value=TempLogins->FieldByName("Code1")->Value;
Logins->FieldByName("Code2")->Value=TempLogins->FieldByName("Code2")->Value;
Logins->FieldByName("Role")->Value=TempLogins->FieldByName("Role")->Value;
Logins->FieldByName("NumDatabase")->Value=Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase;
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
*/

//очистить TempLogins
Comm->CommandText="Delete * from TempLogins";
Comm->Execute();
Main->MClient->WriteDiaryEvent("AdminARM","Конец объединения логинов","");
}
catch(...)
{
Main->MClient->WriteDiaryEvent("Ошибка AdminARM","Ошибка объединения логинов"," Ошибка "+IntToStr(GetLastError()));

}
}
//---------------------------------------------------------------
void TMain::SaveCode(String Login, String Password)
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From Logins Where Login='"+Login+"'"+" AND NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase);
Tab->Active=true;
if(Tab->RecordCount==0 | EditLogins->Mode==3)
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
Tab->FieldByName("NumDatabase")->Value=Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase;
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
//------------------------------------------------------
String TMain::LoadLogins()
{
String Mess="Завершено";
Main->MClient->WriteDiaryEvent("AdminARM","Начало чтения данных","");
try
{

Main->MClient->WriteDiaryEvent("AdminARM","Начало чтения подразделений","");
MP<TADOCommand>Comm(this);
Comm->Connection=Database;

//Перенос данных о подразделениях
Comm->CommandText="Delete * From TempПодразделения";
Comm->Execute();
MP<TADODataSet>Local(this);
Local->Connection=Database;
Local->CommandText="Select TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения] From TempПодразделения Order by [Номер подразделения]";
Local->Active=true;


//CDBItem Item=VDB[CBDatabase->ItemIndex];
Table* Remote=MClient->CreateTable(this, ServerName, Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);

Remote->SetCommandText("Select Подразделения.[Номер подразделения], Подразделения.[Название подразделения] From Подразделения Order by [Номер подразделения]");
Remote->Active(true);

LoadTable(Remote, Local);
//Проверить качество переноса подразделений
Main->MClient->WriteDiaryEvent("AdminARM","Конец чтения подразделений","");
if(VerifyTable(Remote, Local)==0)
{

MergeOtdels();

Main->MClient->WriteDiaryEvent("AdminARM","Начало чтения логинов","");
Comm->CommandText="Delete * From TempLogins";
Comm->Execute();

MP<TADODataSet>LTable(this);
LTable->Connection=Database;
LTable->CommandText="Select TempLogins.Num, TempLogins.Login, TempLogins.Code1, TempLogins.Code2, TempLogins.Role From TempLogins Order by Num";
LTable->Active=true;

//CDBItem Item=VDB[CBDatabase->ItemIndex];
Table* RTable=MClient->CreateTable(this, ServerName, Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);
//
RTable->SetCommandText("Select  Logins.Num, Logins.Login, Logins.Code1, Logins.Code2, Logins.Role From Logins Order by Num");
RTable->Active(true);


//Перенести таблицу логинов
LoadTable(RTable, LTable);
//Проверить качество переноса логинов
if(VerifyTable(LTable, RTable)==0)
{
//Объединить таблицу логинов, одновременно корректируя таблицу TempObslOtdel
MergeLogins();
Main->MClient->WriteDiaryEvent("AdminARM","Начало чтения списка обслуживаемых подразделений","");
Comm->CommandText="Delete * From TempObslOtdel";
Comm->Execute();

MP<TADODataSet>ToTable(this);
ToTable->Connection=Database;
ToTable->CommandText="Select TempObslOtdel.Login, TempObslOtdel.NumObslOtdel From TempObslOtdel Order by Login, NumObslOtdel";
ToTable->Active=true;

//CDBItem Item=VDB[CBDatabase->ItemIndex];
Table* FromTable=MClient->CreateTable(this, ServerName, Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);
//
FromTable->SetCommandText("Select ObslOtdel.Login, ObslOtdel.NumObslOtdel From ObslOtdel Order by Login, NumObslOtdel");
FromTable->Active(true);
LoadTable(FromTable, ToTable);

if(VerifyTable(FromTable, ToTable)==0)
{
Main->MClient->WriteDiaryEvent("AdminARM","Начало объединения списка обслуживаемых подразделений","");
MP<TADODataSet>Otdel(this);
Otdel->Connection=Database;
Otdel->CommandText="Select * from Подразделения Where Numdatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase);
Otdel->Active=true;

MP<TADODataSet>TempObslOtd(this);
TempObslOtd->Connection=Database;
TempObslOtd->CommandText="Select * From TempObslOtdel";
TempObslOtd->Active=true;

MP<TADODataSet>Login(this);
Login->Connection=Database;
Login->CommandText="Select * From Logins  Where Numdatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase);
Login->Active=true;

for(TempObslOtd->First();!TempObslOtd->Eof;TempObslOtd->Next())
{
int Log2;
int Otd2;
int Log1=TempObslOtd->FieldByName("Login")->Value;
if(Login->Locate("ServerNum", Log1, SO))
{
Log2=Login->FieldByName("Num")->Value;
}
else
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Сбой объединения списка обслуживаемых подразделений","Номер логина: "+IntToStr(Log1));
Mess="Ошибка передачи";
}

int Otd1=TempObslOtd->FieldByName("NumObslOtdel")->Value;
if(Otdel->Locate("ServerNum", Otd1, SO))
{
Otd2=Otdel->FieldByName("Номер подразделения")->Value;
}
else
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Сбой объединения списка обслуживаемых подразделений","Номер подразделения: "+IntToStr(Otd1));

Mess="Ошибка передачи";
}

TempObslOtd->Edit();
TempObslOtd->FieldByName("Login")->Value=Log2;
TempObslOtd->FieldByName("NumObslOtdel")->Value=Otd2;
TempObslOtd->Post();
}

Comm->CommandText="Delete * From ObslOtdel Where Numdatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase);
Comm->Execute();

Comm->CommandText="INSERT INTO ObslOtdel ( Login, NumObslOtdel, Numdatabase ) SELECT TempObslOtdel.Login, TempObslOtdel.NumObslOtdel, "+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" AS [Database] FROM TempObslOtdel;";
Comm->Execute();

Comm->CommandText="Delete * From TempObslOtdel";
Comm->Execute();
}
else
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Сбой передачи списка обслуживаемых подразделений ","");

Mess="Ошибка передачи";
}
MClient->DeleteTable(this, FromTable);


}
else
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Сбой передачи логинов ","");
Mess="Ошибка передачи";
}
MClient->DeleteTable(this, RTable);

}
else
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Сбой передачи подразделений ","");
Mess="Ошибка передачи";
}

MClient->DeleteTable(this, Remote);
}
catch(...)
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Ошибка чтения данных"," Ошибка "+IntToStr(GetLastError()));

}
return Mess;
}
//-----------------------------------------------------
void __fastcall TMain::FormShow(TObject *Sender)
{
Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Admin.ini");
ServerName=Ini->ReadString("Server","ServerName","");
/*
ServerName=Ini->ReadString("Server","ServerName","");
String ServerBase=Ini->ReadString("Server","ServerDatabase","");

MClient=new Client();

MClient->Connect(ServerName);

MClient->RegForm(this);

IDDB=MClient->AddDatabase(ServerBase);
MClient->GetDatabaseConnect(MClient->IDC(), IDDB);
*/

MClient=new Client();

MClient->Connect(ServerName, "");

MClient->RegForm(this);

VDB.clear();
CBDatabase->Clear();
ServerName=Ini->ReadString("Server","ServerName","");
int Num=Ini->ReadInteger("Server","NumServerBase",0);
for(int i=0;i<Num;i++)
{
String S="Base"+IntToStr(i+1);
String ServerBase=Ini->ReadString(S,"ServerDatabase","");
String Name=Ini->ReadString(S,"Name","");

MClient->AddDatabase(ServerBase);

CDBItem DBI;
DBI.Name=Name;
DBI.ServerDB=ServerBase;
DBI.Num=i;
DBI.NumDatabase=Ini->ReadInteger(S,"NumDatabase",-1);
if(DBI.NumDatabase>=0)
{
CBDatabase->Items->Add(Name);
}

VDB.push_back(DBI);
}
CBDatabase->ItemIndex=0;
//IDDB=VDB[CBDatabase->ItemIndex].IDDB;
MClient->SetDatabaseConnect(MClient->IDC(), Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase, true);



Start=false;
Pass->ShowModal();

UpdateTempLogin();
Users->ItemIndex=0;

if(Start)
{




}
VerifyLicense();
//MClient->Start();
//Проверка наличия обновления
Main->MClient->WriteDiaryEvent("AdminARM","Начало загрузки данных (автоматическое)","База данных: "+CBDatabase->Text);
UpdateOtdels();

bool B=MClient->NetError();
if(B)
{
// this->Caption="Main  Требуется ручной перезапуск сервера";
}

//MClient->Disconnect();


Main->MClient->WriteDiaryEvent("AdminARM","Запуск модуля админа завершен","База данных: "+CBDatabase->Text);
MClient->Stop();
}
//---------------------------------------------------------------------------
void TMain::MergeOtdels()
{
Main->MClient->WriteDiaryEvent("AdminARM","Начало объединения подразделений","");
try
{
TLocateOptions SO;

MP<TADOCommand>Comm(this);
Comm->Connection=Database;
Comm->CommandText="UPDATE Подразделения SET Подразделения.Del = False;";
Comm->Execute();

MP<TADODataSet>Otdels(this);
Otdels->Connection=Database;
Otdels->CommandText="Select * From Подразделения where NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" Order by [Номер подразделения]";
Otdels->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=Database;
Temp->CommandText="Select * From TempПодразделения Order by [Номер подразделения]";
Temp->Active=true;

for(Otdels->First();!Otdels->Eof;Otdels->Next())
{
int SNum=Otdels->FieldByName("ServerNum")->Value;
if(Temp->Locate("Номер подразделения", SNum, SO))
{
//найден
//Обновить название подразделения
//удалить запись из TempПодразделения
Otdels->Edit();
Otdels->FieldByName("Название подразделения")->Value=Temp->FieldByName("Название подразделения")->Value;
Otdels->Post();

Temp->Delete();
}
else
{
//Ненайден
//Это подразделение на сервере удалено
//Пометить к удалению
Otdels->Edit();
Otdels->FieldByName("Del")->Value=true;
Otdels->Post();
}
}

//Удалить лишние подразделения
Comm->CommandText="Delete * from Подразделения Where Del=true";
Comm->Execute();

//если в TempПодразделения остались записи то это новые подразделения на сервере
//Перенести их в Подразделения

Comm->CommandText="INSERT INTO Подразделения ( ServerNum, [Название подразделения], NumDatabase ) SELECT TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения], "+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" AS [Database] FROM TempПодразделения;";
Comm->Execute();
Main->MClient->WriteDiaryEvent("AdminARM","Конец объединения подразделений","");
}
catch(...)
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Ошибка объединения подразделений"," Ошибка "+IntToStr(GetLastError()));

}

}
//--------------------------------------------------------

void __fastcall TMain::N1Click(TObject *Sender)
{
//Добавить логин

if(Role->ItemIndex==2 & Users->Items->Count>=LicCount & LicCount!=-1)
{
 ShowMessage("Достигнуто максимальное количество пользователей!");
}
else
{
MP<TADODataSet>Tabs(this);
Tabs->Connection=this->Database;
Tabs->CommandText="Select * From Logins where NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase);
Tabs->Active=true;

Tabs->Insert();
Tabs->FieldByName("Login")->Value=" ";
Tabs->FieldByName("Role")->Value=Role->ItemIndex+1;
Tabs->FieldByName("NumDatabase")->Value=Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase;
Tabs->Post();
Tabs->Active=false;
Tabs->Active=true;
Tabs->Last();
EditLoginNumber=Tabs->FieldByName("Num")->Value;

EditLogins->Mode=1;
EditLogins->Login="";
EditLogins->ShowModal();
}

}
//---------------------------------------------------------------------------

void __fastcall TMain::N2Click(TObject *Sender)
{
//Редактировать логин

MP<TADODataSet>Tab(this);
Tab->Connection=this->Database;
Tab->CommandText="Select * From Logins Where Role="+IntToStr(Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(Users->ItemIndex);

EditLoginNumber=Tab->FieldByName("Num")->Value;

EditLogins->Mode=2;
EditLogins->Login=Tab->FieldByName("Login")->AsString;
EditLogins->ShowModal();        
}
//---------------------------------------------------------------------------

void __fastcall TMain::N3Click(TObject *Sender)
{
//Удалить логин
MP<TADODataSet>Tab(this);
Tab->Connection=this->Database;
Tab->CommandText="Select * From Logins Where Role="+IntToStr(Role->ItemIndex+1)+"  AND NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(Users->ItemIndex);

Tab->Delete();
UpdateTempLogin();

int N=Users->ItemIndex;

UpdateOtdel(N);
}
//---------------------------------------------------------------------------

void __fastcall TMain::N4Click(TObject *Sender)
{
//Изменить пароль
MP<TADODataSet>Tab(this);
Tab->Connection=this->Database;
Tab->CommandText="Select * From Logins Where Role="+IntToStr(Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(Users->ItemIndex);

EditLoginNumber=Tab->FieldByName("Num")->Value;

EditLogins->Mode=3;
EditLogins->Login=Tab->FieldByName("Login")->AsString;
EditLogins->ShowModal();
}
//---------------------------------------------------------------------------

String TMain::UpLoadLogins()
{
String Message="";
//
Main->MClient->WriteDiaryEvent("AdminARM","Начало записи данных (автоматическое)","База данных: "+CBDatabase->Text);
try
{
//Label1->Caption="Начато";
MClient->SetCommandText(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB,"Delete * From TempLogins");
MClient->CommandExec(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);
MClient->SetCommandText(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB,"Delete * From TempObslOtdel");
MClient->CommandExec(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);
MClient->SetCommandText(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB,"Delete * From TempПодразделения");
MClient->CommandExec(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);

MP<TADODataSet>CLogins(this);
CLogins->Connection=Database;
CLogins->CommandText="Select  Logins.Num, Logins.Login, Logins.Code1, Logins.Code2, Logins.Role, Logins.ServerNum From Logins where NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" Order by Num";
CLogins->Active=true;

//CDBItem Item=VDB[CBDatabase->ItemIndex];
Table* RLogins=MClient->CreateTable(this, ServerName, Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);
//
RLogins->SetCommandText("Select TempLogins.Num, TempLogins.Login, TempLogins.Code1, TempLogins.Code2, TempLogins.Role, TempLogins.ServerNum From TempLogins Order by Num");
RLogins->Active(true);

LoadTable(CLogins, RLogins);
if(VerifyTable(CLogins, RLogins)==0)
{
MP<TADODataSet>CObslOtd(this);
CObslOtd->Connection=Database;
CObslOtd->CommandText="SELECT ObslOtdel.Login, ObslOtdel.NumObslOtdel FROM Подразделения INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel WHERE (((Подразделения.NumDatabase)="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+")) ORDER BY ObslOtdel.Login, ObslOtdel.NumObslOtdel;";
CObslOtd->Active=true;

//CDBItem Item=VDB[CBDatabase->ItemIndex];
Table* RObslOtd=MClient->CreateTable(this, ServerName, Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);
//
RObslOtd->SetCommandText("Select TempObslOtdel.Login, TempObslOtdel.NumObslOtdel From TempObslOtdel Order by Login, NumObslOtdel");
RObslOtd->Active(true);

LoadTable(CObslOtd, RObslOtd);

if(VerifyTable(CObslOtd, RObslOtd)==0)
{
MP<TADODataSet>Podrazd(this);
Podrazd->Connection=Database;
Podrazd->CommandText="Select [Подразделения].[Номер подразделения], [Подразделения].[Название подразделения], [Подразделения].ServerNum from Подразделения Where NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" Order by [Подразделения].[Номер подразделения]";
Podrazd->Active=true;

//CDBItem Item=VDB[CBDatabase->ItemIndex];
Table* TempPodrazd=MClient->CreateTable(this, ServerName, Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);
//
TempPodrazd->SetCommandText("Select [TempПодразделения].[Номер подразделения], [TempПодразделения].[Название подразделения], [TempПодразделения].ServerNum From [TempПодразделения] order by [TempПодразделения].[Номер подразделения]");
TempPodrazd->Active(true);

LoadTable(Podrazd, TempPodrazd);
if(VerifyTable(Podrazd, TempPodrazd)==0)
{


bool B=MClient->MergeLogins(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);
if(B)
{
MP<TADOCommand>Comm(this);
Comm->Connection=Database;
Comm->CommandText="Delete * from TempLogins";
Comm->Execute();

MP<TADODataSet>TempLogin(this);
TempLogin->Connection=Database;
TempLogin->CommandText="SELECT TempLogins.Num, TempLogins.ServerNum FROM TempLogins ORDER BY Num";
TempLogin->Active=true;

//CDBItem Item=VDB[CBDatabase->ItemIndex];
Table* STempLogin=MClient->CreateTable(this, ServerName, Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);
//
STempLogin->SetCommandText("SELECT TempLogins.Num, TempLogins.ServerNum FROM TempLogins ORDER BY Num");
STempLogin->Active(true);
Main->MClient->WriteDiaryEvent("AdminARM","Начало чтения с сервера номеров новых логинов","База данных: "+CBDatabase->Text);
LoadTable(STempLogin, TempLogin);

if(VerifyTable(STempLogin, TempLogin)==0)
{
CLogins->Active=false;
CLogins->CommandText="Select  * From Logins where NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" Order by Num";
CLogins->Active=true;

for(TempLogin->First();!TempLogin->Eof;TempLogin->Next())
{
 int N=TempLogin->FieldByName("Num")->Value;
 if(CLogins->Locate("Num",N,SO))
 {
  CLogins->Edit();
  CLogins->FieldByName("ServerNum")->Value=TempLogin->FieldByName("ServerNum")->Value;
  CLogins->Post();
 }
 else
 {
 Main->MClient->WriteDiaryEvent("AdminARM ошибка","Сбой чтения с сервера номеров новых логинов","База данных: "+CBDatabase->Text);
  ShowMessage("Ошибка получения ServerNum для новых логинов N="+IntToStr(N));
 }
}
Main->MClient->WriteDiaryEvent("AdminARM","Конец чтения с сервера номеров новых логинов","База данных: "+CBDatabase->Text);
Comm->CommandText="Delete * from TempLogins";
Comm->Execute();
}
else
{
Main->MClient->WriteDiaryEvent("AdminARM","Сбой объединения логинов на сервере","База данных: "+CBDatabase->Text);
}
MClient->DeleteTable(this, STempLogin);
Main->MClient->WriteDiaryEvent("AdminARM","Конец записи данных","База данных: "+CBDatabase->Text);
Message="Завершено";
}
else
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Сбой объединения логинов на сервере","База данных: "+CBDatabase->Text);
Message="Ошибка";
}
}
else
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Сбой записи подразделений","База данных: "+CBDatabase->Text);
Message="Ошибка передачи";
}
MClient->DeleteTable(this, TempPodrazd);
}
else
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Сбой записи обслуживаемых подразделений","База данных: "+CBDatabase->Text);
Message="Ошибка передачи";
}
MClient->DeleteTable(this, RObslOtd);

}
else
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Сбой записи логинов","База данных: "+CBDatabase->Text);
Message="Ошибка передачи";
}
MClient->DeleteTable(this, RLogins);
//ShowMessage(Message);


MClient->SetCommandText(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB,"Delete * From TempLogins");
MClient->CommandExec(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);
MClient->SetCommandText(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB,"Delete * From TempObslOtdel");
MClient->CommandExec(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);

Main->MClient->WriteDiaryEvent("AdminARM","Конец записи данных (автоматическое)","База данных: "+CBDatabase->Text);
}
catch(...)
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Ошибка записи данных (автоматическое)","База данных: "+CBDatabase->Text+" Ошибка "+IntToStr(GetLastError()));
}
return Message;
}
//------------------------------------------------------------------------
bool TMain::VerifyUpdates()
{
//Для админа проверяем только обновление списка подразделений
//которое редактируется главным специалистом
Main->MClient->WriteDiaryEvent("AdminARM","Проверка обновлений","");
try
{
MP<TADODataSet>Otdels(this);
Otdels->Connection=Database;
Otdels->CommandText="Select Подразделения.ServerNum, Подразделения.[Название подразделения] From Подразделения where NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" Order by [ServerNum]";
Otdels->Active=true;

//CDBItem Item=VDB[CBDatabase->ItemIndex];
Table* ROtdels=MClient->CreateTable(this, ServerName, Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);

ROtdels->SetCommandText("Select Подразделения.[Номер подразделения], Подразделения.[Название подразделения] From Подразделения Order by [Номер подразделения]");
ROtdels->Active(true);

int Res=VerifyTable(Otdels, ROtdels);
MClient->DeleteTable(this, ROtdels);

if(Res==0)
{
Main->MClient->WriteDiaryEvent("AdminARM","Обновление ненужно","");
return true;
}
else
{
Main->MClient->WriteDiaryEvent("AdminARM","Обновление необходимо","");

return false;
}
}
catch(...)
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Ошибка проверки обновлений"," Ошибка "+IntToStr(GetLastError()));
return false;
}
}
//-------------------------------------------------------
void TMain::UpdateOtdels()
{
//Main->MClient->WriteDiaryEvent("AdminARM","Начало загрузки данных (ручное)","База данных: "+CBDatabase->Text);
if(VerifyUpdates())
{
// ShowMessage("В таблице отделов нет изменений");
}
else
{
 //ShowMessage("В таблице отделов есть изменения");
 if(Application->MessageBox("В таблице отделов есть изменения, обновить таблицу отделов?","Синхронизация",MB_YESNO+MB_ICONEXCLAMATION+MB_DEFBUTTON1)==IDYES )
 {
Main->MClient->WriteDiaryEvent("AdminARM","Начало обновление подразделений","База данных: "+CBDatabase->Text);
MP<TADOCommand>Comm(this);
Comm->Connection=Database;
Comm->CommandText="Delete * From TempПодразделения";
Comm->Execute();

MP<TADODataSet>TOtdels(this);
TOtdels->Connection=Database;
TOtdels->CommandText="Select TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения] From TempПодразделения Order by [Номер подразделения]";
TOtdels->Active=true;

//CDBItem Item=VDB[CBDatabase->ItemIndex];
Table* ROtdels=MClient->CreateTable(this, ServerName, Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);
//
ROtdels->SetCommandText("Select Подразделения.[Номер подразделения], Подразделения.[Название подразделения] From Подразделения Order by [Номер подразделения]");
ROtdels->Active(true);

LoadTable(ROtdels, TOtdels);

TOtdels->Active=false;
TOtdels->Active=true;
int B=VerifyTable(ROtdels, TOtdels);
if(B==0)
{
Main->MClient->WriteDiaryEvent("AdminARM","Начало объединения подразделений","");
Comm->CommandText="UPDATE Подразделения SET Подразделения.Del = False;";
Comm->Execute();

MP<TADODataSet>Otdels(this);
Otdels->Connection=Database;
Otdels->CommandText="Select * From Подразделения Where NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" Order by [Номер подразделения]";
Otdels->Active=true;

for(Otdels->First();!Otdels->Eof;Otdels->Next())
{
int N=Otdels->FieldByName("ServerNum")->AsInteger;

if(TOtdels->Locate("Номер подразделения", N, SO))
{
//совпадение найдено
Otdels->Edit();
Otdels->FieldByName("Название подразделения")->Value=TOtdels->FieldByName("Название подразделения")->Value;
Otdels->Post();

TOtdels->Delete();
}
else
{
//совпадения ненайдено
Otdels->Edit();
Otdels->FieldByName("Del")->Value=true;
Otdels->Post();
}
}

Comm->CommandText="DELETE Подразделения.Del FROM Подразделения WHERE (((Подразделения.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO Подразделения ( ServerNum, [Название подразделения], NumDatabase ) SELECT TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения], "+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" AS Datbase FROM TempПодразделения;";
Comm->Execute();

Comm->CommandText="Delete * From TempПодразделения";
Comm->Execute();
/*
TOtdels->Active=false;
TOtdels->CommandText="Select * From TempПодразделения";
TOtdels->Active=true;

Otdels->Active=false;
Otdels->CommandText="Select * From Подразделения Where NumDatabase="+IntToStr(Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase)+" Order by [Номер подразделения]";
Otdels->Active=true;

for(TOtdels->First();!TOtdels->Eof;TOtdels->Next())
{
Otdels->Insert();
Otdels->FieldByName("Название подразделения")->Value=TOtdels->FieldByName("Название подразделения")->Value;
Otdels->FieldByName("NumDatabase")->Value=Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].NumDatabase;
Otdels->FieldByName("ServerNum")->Value=TOtdels->FieldByName("NumMSp")->Value;
Otdels->Post();
}

Comm->CommandText="Delete * From TempПодразделения";
Comm->Execute();
*/
Main->MClient->WriteDiaryEvent("AdminARM","Конец объединения подразделений","");
Users->ItemIndex=0;
UpdateOtdel(0);

ShowMessage("Обновление отделов завершено");
}
else
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Сбой объединения подразделений","");
ShowMessage("Обновление завершилось неудачно");
}


MClient->DeleteTable(this, ROtdels);

 }
else
{
Main->MClient->WriteDiaryEvent("AdminARM","Отказ от обновления подразделений","");
}
}
}
//-------------------------------------------------------






void __fastcall TMain::N5Click(TObject *Sender)
{
MClient->Start();
Main->MClient->WriteDiaryEvent("AdminARM","Начало загрузки данных (ручное)","База данных: "+CBDatabase->Text);

//UpdateOtdels();


String Mess=LoadLogins();
Main->MClient->WriteDiaryEvent("AdminARM","Конец загрузки данных (ручное)","База данных: "+CBDatabase->Text);
MClient->Stop();
UpdateTempLogin();
Users->ItemIndex=0;
UpdateOtdel(0);

ShowMessage(Mess);
int N=Users->ItemIndex;

UpdateOtdel(N);

Sleep(1000);
}
//---------------------------------------------------------------------------

void __fastcall TMain::N6Click(TObject *Sender)
{
MClient->Start();
Main->MClient->WriteDiaryEvent("AdminARM","Начало записи данных (ручное)","База данных: "+CBDatabase->Text);

String S=UpLoadLogins();
Main->MClient->WriteDiaryEvent("AdminARM","Конец записи данных (ручное)","База данных: "+CBDatabase->Text);
MClient->Stop();
ShowMessage(S);
Sleep(1000);
}
//---------------------------------------------------------------------------



void __fastcall TMain::CBDatabaseClick(TObject *Sender)
{

MClient->Start();

Main->MClient->WriteDiaryEvent("AdminARM","Переключение базы данных","База данных: "+CBDatabase->Text);
LicCount=MClient->GetLicenseCount(CBDatabase->Text);
Reg=LicCount!=0;
//MClient->SetDatabaseConnect(MClient->IDC(), CBDatabase->ItemIndex, true);
try
{
VerifyLicense();

UpdateOtdels();

int N=Users->ItemIndex;

UpdateTempLogin();
Users->ItemIndex=N;
UpdateOtdel(N);
Main->MClient->WriteDiaryEvent("AdminARM","Переключение базы данных завершено","База данных: "+CBDatabase->Text);
}
catch(...)
{
Main->MClient->WriteDiaryEvent("AdminARM ошибка","Ошибка переключения базы данных","База данных: "+CBDatabase->Text);
}
MClient->Stop();
Sleep(1000);
}
//---------------------------------------------------------------------------
int TMain::GetIDDBName(String Name)
{
int Ret=-1;
for(unsigned int i=0;i<VDB.size();i++)
{
 if(VDB[i].Name==Name)
 {
  Ret=i;
  break;
 }
}
return Ret;
}
//---------------------------------------------------------------------------
void __fastcall TMain::FormClose(TObject *Sender, TCloseAction &Action)
{
if(InputPass)
{
switch (Application->MessageBoxA("Записать все изменения на сервер?","Выход из программы",MB_YESNOCANCEL	+MB_ICONQUESTION+MB_DEFBUTTON1))
{
 case IDYES:
 {
MClient->Start();
Main->MClient->WriteDiaryEvent("AdminARM","Завершение работы запись данных","База данных: "+CBDatabase->Text);
String S=UpLoadLogins();
MClient->Stop();
 break;
 }
 case IDNO:
 {
Main->MClient->WriteDiaryEvent("AdminARM","Завершение работы отказ от записи","База данных: "+CBDatabase->Text);
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
}
//---------------------------------------------------------------------------




void __fastcall TMain::FormDestroy(TObject *Sender)
{
delete MClient;
}
//---------------------------------------------------------------------------

void __fastcall TMain::N7Click(TObject *Sender)
{
 FDiary->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TMain::N8Click(TObject *Sender)
{
FAbout->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TMain::N9Click(TObject *Sender)
{
this->Close();
}
//---------------------------------------------------------------------------
void TMain::VerifyLicense()
{
LicCount=Main->MClient->GetLicenseCount(CBDatabase->Text);
Reg=LicCount!=0;
MP<TADODataSet>Log3(this);
Log3->Connection=Database;
Log3->CommandText="Select * From Logins Where Role=3 AND NumDatabase="+IntToStr(CBDatabase->ItemIndex);
Log3->Active=true;
if(LicCount==-1)
{
Main->MClient->WriteDiaryEvent("AdminARM лицензия","Неограниченая лицензия","База данных: "+CBDatabase->Text);

}
else
{
if(Log3->RecordCount>LicCount)
{
Main->MClient->WriteDiaryEvent("AdminARM лицензия","Удалено часть пользователей по причине лицензии","База данных: "+CBDatabase->Text+" Было: "+IntToStr(Log3->RecordCount)+" Стало: "+IntToStr(LicCount));
Main->MClient->WriteDiaryEvent("AdminARM лицензия","Чтение списка логинов","База данных: "+CBDatabase->Text+" Было: "+IntToStr(Log3->RecordCount)+" Стало: "+IntToStr(LicCount));

LoadLogins();
 //загрузка логинов
 Main->MClient->WriteDiaryEvent("AdminARM лицензия","Загрузка логинов завершена","База данных: "+CBDatabase->Text);

}
}
}
//-------------------------------------------------------------------------
void __fastcall TMain::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"ADMINARM.HLP";
Application->HelpJump("IDH_ГЛАВНАЯ_ФОРМА");
}
}
//---------------------------------------------------------------------------

void __fastcall TMain::N10Click(TObject *Sender)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"ADMINARM.HLP";
Application->HelpJump("IDH_РУКОВОДСТВО");        
}
//---------------------------------------------------------------------------

