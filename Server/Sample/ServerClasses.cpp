//---------------------------------------------------------------------------


#pragma hdrstop

#include "ServerClasses.h"
#include "MasterPointer.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

mClients::mClients()
{
LastIDC=1;


}
//---------------------------------------------
mClients::mClients(TComponent* Owner, String DiaryBase)
{
NameComp="Не определен";
Login="Еще не определен";
//mClients();
LastIDC=1;

pOwner=Owner;
Form1->Net_Error=false;

this->DiaryBase=DiaryBase;

DiaryEvent=new Diary(Owner, DiaryBase);

DiaryEvent->WriteEvent(Now(), NameComp, Login, "Служебное", "Создание класса сервера");
//WriteEvent(TDateTime DT, String Comp, String Login, String Type, String Name)
}
//--------------------------------------------
mClients::~mClients()
{
for(unsigned int i=0;i<VClients.size();i++)
{
delete VClients[i];
}
VClients.clear();

DiaryEvent->WriteEvent(Now(), NameComp, Login, "Служебное", "Удаление класса сервера");
delete DiaryEvent;
}
//--------------------------------------------
long mClients::Connect(String NameComp, String Login)
{

long IDC=LastIDC;
LastIDC++;

Client *C=new Client(IDC, NameComp, Login, Form1, DiaryBase);
//C->ClTimer->Interval=10000;
C->ClTimer->Interval=600000;
C->ClTimer->OnTimer=ClientTimer;
C->ClTimer->Enabled=true;
VClients.push_back(C);


if (FCountClientChanged) CountClientChanged (Sender) ;

this->NameComp=NameComp;
if(Login=="")
{
 Login="Еще не определен";
}
this->Login=Login;
DiaryEvent->WriteEvent(Now(), NameComp, Login, "Служебное", "Подключение клиента", "Comp: "+NameComp+" Login: "+Login+" IDC="+IntToStr(IDC));
return IDC;
}
//--------------------------------------------
bool mClients::Disconnect(long IDC)
{
DiaryEvent->WriteEvent(Now(), NameComp, Login, "Служебное", "Отключение клиента", "IDC="+IntToStr(IDC));
bool Ret=false;
IVC=VClients.begin();
for(unsigned int i=0;i<VClients.size();i++)
{
 if(VClients[i]->IDC()==IDC)
 {
 Ret=true;
  VClients[i]->ClTimer->Enabled=false;
  //delete VClients[i]->ClTimer;
  delete VClients[i];
  VClients.erase(IVC);
  break;
 }
 IVC++;
}

if (FCountClientChanged) CountClientChanged (Sender) ;
return Ret;
}
//-------------------------------------------
bool mClients::Disconnect(String Name)
{

bool Ret=false;
IVC=VClients.begin();
for(unsigned int i=0;i<VClients.size();i++)
{
 if(VClients[i]->NameComp()==Name)
 {
 Ret=true;
 DiaryEvent->WriteEvent(Now(), NameComp, Login, "Служебное", "Отключение клиента", "IDC="+IntToStr(VClients[i]->IDC()));
  delete VClients[i];
  VClients.erase(IVC);
  break;
 }
 IVC++;
}
if (FCountClientChanged) CountClientChanged (Sender) ;

return Ret;
}
//-------------------------------------------
bool mClients::DisconnectAll()
{
bool Ret=false;
IVC=VClients.begin();
for(unsigned int i=0;i<VClients.size();i++)
{
 Ret=true;
  delete VClients[i];
  VClients.erase(IVC);
 IVC++;
}
if (FCountClientChanged) CountClientChanged (Sender) ;
DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Служебное", "Отключение всех клиентов");
return Ret;
}
//--------------------------------------------
long mClients::ClientCount()
{
return (long)VClients.size();
}
//---------------------------------------------
String mClients::GetName(int Number)
{
return VClients[Number]->NameComp();
}
//-------------------------------------------
int mClients::GetIDC(int Number)
{
return VClients[Number]->IDC();
}
//-------------------------------------------
Client* mClients::GetClient(long IDC)
{

for(unsigned int i=0;i<VClients.size();i++)
{
 if (VClients[i]->IDC()==IDC)
 {
  return VClients[i];

 }
}
return 0;
}
//-------------------------------------------
Client* mClients::GetClient(int Number)
{
 return VClients[Number];
}
//-------------------------------------------
long mClients::AddDatabase(long IDC, String DatabaseName)
{

long IDDB=0;
GetClient(IDC)->AddDB(DatabasePath, DatabaseName, &IDDB, Form1);

//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Служебное", "Добавление базы данных",DatabasePath+DatabaseName+" IDDB="+IntToStr(IDDB) );
return IDDB;

}
//------------------------------------------
void mClients::SetDatabasePatc(String Path)
{
DatabasePath=Path;
}
//-------------------------------------------

bool mClients::GetDatabaseConnect(long IDC, long IDDB)
{

  return GetClient(IDC)->GetDatabaseConnect(IDDB);

}
//----------------------------------------------
bool mClients::SetDatabaseConnect(long IDC, long IDDB, bool Connect)
{

  GetClient(IDC)->SetDatabaseConnect(IDDB, Connect);
  return true;

}
//-----------------------------------------------
long mClients::CreateTable(long IDC, long IDDB, long IDF)
{

//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Служебное", "Создание таблицы","IDC="+IntToStr(IDC)+" IDDB="+IntToStr(IDDB)+" IDF="+IntToStr(IDF));
  return GetClient(IDC)->CreateTable(IDDB, IDF);

}
//----------------------------------------------

bool mClients::DeleteTable(long IDC, long IDF, long IDT)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Служебное", "Удаление таблицы","IDC="+IntToStr(IDC)+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT));
  return GetClient(IDC)->DeleteTable(IDF, IDT);

}
//-------------------------------------------------
void mClients::SetCommandText(String Text, long IDC, long IDF, long IDT)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Служебное", "TableCommandText","IDC="+IntToStr(IDC)+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT)+" Text="+Text);

GetClient(IDC)->SetCommandText(Text, IDF, IDT);

}
//-------------------------------------------------
String mClients::GetCommandText(long IDC, long IDF, long IDT)
{
return GetClient(IDC)->GetCommandText(IDF, IDT);
}
//-------------------------------------------------
void mClients::Active(bool Value, long IDC, long IDF, long IDT)
{

//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Служебное", "Активация таблицы","IDC="+IntToStr(IDC)+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT)+" Value="+V);

GetClient(IDC)->Active(Value, IDF, IDT);
}
//-------------------------------------------------
void  mClients::Moving(long Comand, long IDC, long IDF, long IDT)
{
GetClient(IDC)->Moving(Comand, IDF, IDT);
}
//-------------------------------------------------
bool mClients::BOF_EOF(long Comand, long IDC, long IDF, long IDT)
{
return GetClient(IDC)->BOF_EOF(Comand, IDF, IDT);
}
//--------------------------------------------------
Variant mClients::FieldByName(String FieldName, long IDC, long IDF, long IDT)
{
return GetClient(IDC)->FieldByName(FieldName, IDF, IDT);
}
//--------------------------------------------------
void mClients::FieldByName(Variant Value, String FieldName, long IDC, long IDF, long IDT)
{
GetClient(IDC)->FieldByName(Value, FieldName, IDF, IDT);
}
//---------------------------------------------------
void mClients::EditingRecord(long Comand, long IDC, long IDF, long IDT)
{
GetClient(IDC)->EditingRecord(Comand, IDF, IDT);
}


//----------------------------------------------------
long mClients::RecordCount(long IDC, long IDF, long IDT)
{
return GetClient(IDC)->RecordCount(IDF, IDT);
}
//-------------------------------------------------------
long mClients::RecNo(long IDC, long IDF, long IDT)
{
return GetClient(IDC)->RecNo(IDF, IDT);
}
//-------------------------------------------------------
void mClients::Move(long Step, long IDC, long IDF, long IDT)
{
GetClient(IDC)->Move(Step, IDF, IDT);
}
//-------------------------------------------
bool mClients::Locate(long IDC, long IDF, long IDT, String FieldName, String Key, long SearchParam)
{
TLocateOptions SearchOptions;
switch(SearchParam)
{
 case 0:
 {

 break;
 }
 case 1:
 {
 SearchOptions<<loCaseInsensitive;
 break;
 }
 case 2:
 {
 SearchOptions<<loPartialKey;
 break;
 }
 case 3:
 {
 SearchOptions<<loCaseInsensitive<<loPartialKey;
 break;
 }


}
return GetClient(IDC)->Locate(IDF, IDT, FieldName, Key, SearchOptions);
}
//--------------------------------------------
void mClients::Delete(long IDC, long IDF, long IDT)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Служебное", "Удаление текущей записи","IDC="+IntToStr(IDC)+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT));

GetClient(IDC)->Delete(IDF, IDT);
}
//-------------------------------------------
Variant mClients::Fields(long NumField, long IDC, long IDF, long IDT)
{
return GetClient(IDC)->Fields(NumField, IDF, IDT);
}
//-------------------------------------------
void mClients::Fields(Variant Value, long NumField, long IDC, long IDF, long IDT)
{
GetClient(IDC)->Fields(Value, NumField, IDF, IDT);
}
//-----------------------------------------------------
long mClients::FieldsCount(long IDC, long IDF, long IDT)
{
return GetClient(IDC)->FieldsCount(IDF, IDT);
}
//------------------------------------------------------
TADOCommand* mClients::Command(long IDC, long IDDB)
{
return GetClient(IDC)->Command(IDDB);
}
//--------------------------------------------------
bool mClients::MergeLogins(long IDC, long IDDB)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Администратор", "Объединение логинов","IDC="+IntToStr(IDC)+" IDDB="+IntToStr(IDDB));

return GetClient(IDC)->MergeLogins(IDDB);
}
//------------------------------------------------
void __fastcall mClients::ClientTimer(TObject *Sender)
{

TTimer* T=(TTimer*)Sender;
long IDC=T->Tag;
Form1->Label3->Caption=IDC;
T->Enabled=false;
DiaryEvent->WriteEvent(Now(), NameComp, Login, "Служебная ошибка", "Удаление упавшего клиента", "IDC="+IntToStr(IDC)+" Компьютер: "+GetClient(IDC)->GetComputer()+" Login: "+GetClient(IDC)->GetLogin());


Disconnect(IDC);

Form1->Net_Error=true;

Form1->Error();
}
//-------------------------------------------
void mClients::VerifyConnect(long IDC)
{

GetClient(IDC)->ClTimer->Tag=IDC;
GetClient(IDC)->ClTimer->Enabled=false;
GetClient(IDC)->ClTimer->Interval=600000;
GetClient(IDC)->ClTimer->Enabled=true;

}
//-----------------------------------------------
bool mClients::PrepareReadMetod(long IDC, String NameDatabase)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Подготовка к чтению методики","IDC="+IntToStr(IDC));

return GetClient(IDC)->PrepareReadMetod(NameDatabase);
}
//------------------------------------------------
bool mClients::MergeNodeBranch(long IDC, String NameDatabase1, String NameNode, String NameBranch, String NameDatabase2, String NameTable, String NameField, String NameKey, String Name)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Объединение узлов и ветвей (Full)","IDC="+IntToStr(IDC)+" Узел: "+NameNode+" Ветвь: "+NameBranch);
return GetClient(IDC)->MergeNodeBranch(NameDatabase1, NameNode, NameBranch, NameDatabase2, NameTable, NameField, NameKey, Name);
}
//------------------------------------------------
bool mClients::MergeNodeBranch(long IDC, String NameDatabase1, String NameNode, String NameBranch)
{
//ShowMessage("Задержка");
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Объединение узлов и ветвей (Low)","IDC="+IntToStr(IDC)+" Узел: "+NameNode+" Ветвь: "+NameBranch);

return GetClient(IDC)->MergeNodeBranch(NameDatabase1, NameNode, NameBranch);
}
//------------------------------------------------
void mClients::MergeZn(long IDC, String Database1, String Database2)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Объединение критериев","IDC="+IntToStr(IDC)+" DB1: "+Database1+" DB2: "+Database2);

GetClient(IDC)->MergeZn(Database1, Database2);
}
//-------------------------------------------------
void mClients::SaveMetod(long IDC, String Database)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Запись методики","IDC="+IntToStr(IDC)+" DB: "+Database);

GetClient(IDC)->SaveMetod(Database);
}
//-------------------------------------------------
void mClients::SavePodr(long IDC, String DB1)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Запись подразделений","IDC="+IntToStr(IDC)+" DB1: "+DB1);

GetClient(IDC)->SavePodr(DB1);
}
//--------------------------------------------------
void mClients::SaveSit(long IDC, String DB1, String DB2)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Запись ситуаций","IDC="+IntToStr(IDC)+" DB1: "+DB1+" DB2: "+DB2);

GetClient(IDC)->SaveSit(DB1, DB2);
}
//--------------------------------------------------
void mClients::MergeAspectsMainSpec(long IDC)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Объединение аспектов (главспец)","IDC="+IntToStr(IDC));

GetClient(IDC)->MergeAspectsMainSpec();
}
//--------------------------------------------------
void mClients::MergeAspectsMainSpecH(long IDC)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Объединение аспектов (главспец)","IDC="+IntToStr(IDC));

GetClient(IDC)->MergeAspectsMainSpecH();
}
//--------------------------------------------------
void mClients::PrepareLoadAspects(long IDC, String NameDatabase, long Podr, long Width1, long Width2, long Width3, long Width4)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Пользователи", "Подготовка к чтению аспектов","IDC="+IntToStr(IDC)+" DB: "+NameDatabase+" Podr="+IntToStr(Podr));

GetClient(IDC)->PrepareLoadAspects(NameDatabase, Podr, Width1, Width2,  Width3, Width4);
}
//----------------------------------------------------
void mClients::PrepareLoadAspectsH(long IDC, String NameDatabase, long Podr, long Width1, long Width2, long Width3, long Width4)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Пользователи", "Подготовка к чтению аспектов","IDC="+IntToStr(IDC)+" DB: "+NameDatabase+" Podr="+IntToStr(Podr));

GetClient(IDC)->PrepareLoadAspectsH(NameDatabase, Podr, Width1, Width2,  Width3, Width4);
}
//----------------------------------------------------
void mClients::PrepareMergeAspects(String NameDatabase, long IDC, long NumLogin, long W1, long W2, long W3, long W4)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Пользователи", "Подготовка к объединению аспектов","IDC="+IntToStr(IDC)+" NumLogin="+IntToStr(NumLogin));

GetClient(IDC)->PrepareMergeAspects(NameDatabase, NumLogin, W1, W2, W3, W4);
}
//----------------------------------------------------
void mClients::PrepareMergeAspectsH(String NameDatabase, long IDC, long NumLogin, long W1, long W2, long W3, long W4)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Пользователи", "Подготовка к объединению аспектов","IDC="+IntToStr(IDC)+" NumLogin="+IntToStr(NumLogin));

GetClient(IDC)->PrepareMergeAspectsH(NameDatabase, NumLogin, W1, W2, W3, W4);
}
//----------------------------------------------------
void mClients::WriteDiaryEvent(String Comp, String Login, String Type, String Name, String Prim)
{
DiaryEvent->WriteEvent(Now(), Comp, Login, Type, Name,Prim);

}
//----------------------------------------------------
//********************************************
Client::Client()
{
LastIDF=1;
LastIDDB=1;
ClTimer=new TTimer(Form1);
ClTimer->Tag=pIDC;

}
//********************************************
Client::Client(TComponent *S)
{
Login="Еще не определен";
//Client();
LastIDF=1;
LastIDDB=1;
Sender=S;

ClTimer=new TTimer(Form1);
ClTimer->Tag=pIDC;
//
//Dict* D=new Dict(
}
///******************************************
Client::~Client()
{
delete ClTimer;
for(unsigned int i=0;i<VForm.size();i++)
{
 delete VForm[i];
}
VForm.clear();

for(unsigned int i=0;i<VDatabase.size();i++)
{
VDatabase[i]->Database->Connected=false;
delete VDatabase[i]->Database;
delete VDatabase[i]->Command;
delete VDatabase[i];
}
VDatabase.clear();

DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Клиент", "Удаление клиента","IDC="+IntToStr(this->IDC()));
}
//*********************************************
Client::Client(long idc, String NC, String Login, TControl *Owner, String DiaryBase)
{
LastIDF=1;
LastIDDB=1;
pNameComp=NC;
pIDC=idc;
pOwner=Owner;
ClTimer=new TTimer(Form1);
ClTimer->Tag=pIDC;

if(Login=="")
{
 Login="Еще не определен";
}
this->Login=Login;
DiaryEvent=new Diary(Owner, DiaryBase);
//ShowMessage(Login);
//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Клиент", "Подключение клиента");
}
//**********************************************
String Client::NameComp()
{
 return pNameComp;
}
//*********************************************
long Client::IDC()
{
 return pIDC;
}
//***********************************************
long Client::RegForm(String Name)
{
long IDF=LastIDF;
LastIDF++;
mForm *F=new mForm(IDF, Name);
VForm.push_back(F);
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебное", "Регистрация формы","IDF="+IntToStr(IDF)+" Форма: "+Name);
return IDF;
}
//*************************************************
bool Client::UnRegForm(long IDF)
{
IFC=VForm.begin();
bool B=false;
for(unsigned int i=0;i<VForm.size();i++)
{
 if(VForm[i]->IDF()==IDF)
 {
  B=true;
  delete VForm[i];
  VForm.erase(IFC);
 }
 IFC++;
}
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебное", "Удаление регистрации формы", "IDF="+IntToStr(IDF));

return B;
}
//*************************************************
bool Client::UnRegForm(String Name)
{
IFC=VForm.begin();
bool B=false;
for(unsigned int i=0;i<VForm.size();i++)
{
 if(VForm[i]->Name()==Name)
 {
  B=true;
  delete VForm[i];
  VForm.erase(IFC);
  break;
 }
 IFC++;
}
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебное", "Удаление регистрации формы"," Форма: "+Name);

return B;
}
//*************************************************
bool Client::UnRegAll()
{
for(unsigned int i=0;i<VForm.size();i++)
{
  delete VForm[i];
}
VForm.clear();
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебное", "Удаление регистрации всех форм");

return true;
}
//*************************************************
long Client::FormCount()
{
return VForm.size();
}
//*************************************************
mForm* Client::GetForm(long IDF)
{
 for(unsigned int i=0;i<VForm.size();i++)
 {
 if(VForm[i]->IDF()==IDF)
 {
  return VForm[i];
 }
 }
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка получения формы", IntToStr(IDF));

 return NULL;
}
//*************************************************
mForm* Client::GetForm(int Number)
{
 return VForm[Number];
}
//*************************************************
bool Client::AddDB(String DBPath, String Name, long *IDDB, TComponent* Owner)
{
try
{
DatabaseClient* DBC=new DatabaseClient();

DBC->Database=new TADOConnection (Owner);
DBC->Command=new TADOCommand(Owner);
DBC->FileName=Name;
String P;

P=DBPath+Name+".dtb";

P="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+P+";Persist Security Info=False";

DBC->Database->ConnectionString=P;
DBC->Database->LoginPrompt=false;
DBC->Database->Connected=true;

DBC->Command->Connection=DBC->Database;
*IDDB=LastIDDB;

DBC->IDDB=LastIDDB;
LastIDDB++;
VDatabase.push_back(DBC);
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Клиент", "Подключение к базе данных","IDDB="+IntToStr(DBC->IDDB)+" DB: "+DBPath+Name+".dtb");

return true;
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка добавления базы", "DB: "+DBPath+Name);

return false;
}
}
//*************************************************
TADOCommand* Client::Command(long IDDB)
{
try
{
for(unsigned int i=0;i<VDatabase.size();i++)
{
 if(VDatabase[i]->IDDB==IDDB)
 {
  return VDatabase[i]->Command;
 }
}
return NULL;
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка получения TADOCommand", "IDDB="+IntToStr(IDDB));

return NULL;
}
}
//*************************************************
bool Client::MergeLogins(long IDDB)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Администратор", "Объединение логинов","IDC="+IntToStr(IDC)+" IDDB="+IntToStr(IDDB));
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Администратор", "Объединение логинов", "IDC="+IntToStr(IDC())+"IDDB="+IntToStr(IDDB));

try
{
TADOConnection *Database=NULL;
for(unsigned int i=0;i<VDatabase.size();i++)
{
 if(VDatabase[i]->IDDB==IDDB)
 {
  Database=VDatabase[i]->Database;
  break;
 }
}
if(Database!=NULL)
{

MP<TADOCommand>Comm(Form1);
Comm->Connection=Database;
Comm->CommandText="UPDATE Logins SET Logins.Del = False;";
Comm->Execute();

MP<TADODataSet>TempLogins(Form1);
TempLogins->Connection=Database;
TempLogins->CommandText="Select * From TempLogins";
TempLogins->Active=true;

MP<TADODataSet>Logins(Form1);
Logins->Connection=Database;
Logins->CommandText="Select * From Logins Order by Num";
Logins->Active=true;

MP<TADODataSet>TempObslOtdel(Form1);
TempObslOtdel->Connection=Database;
TempObslOtdel->CommandText="Select * from TempObslOtdel";
TempObslOtdel->Active=true;

MP<TADODataSet>ObslOtdel(Form1);
ObslOtdel->Connection=Database;
ObslOtdel->CommandText="Select * from ObslOtdel";
ObslOtdel->Active=true;

MP<TADODataSet>TempPodr(Form1);
TempPodr->Connection=Database;
TempPodr->CommandText="Select * from TempПодразделения";
TempPodr->Active=true;

for(TempLogins->First();!TempLogins->Eof;TempLogins->Next())
{
 String Log=TempLogins->FieldByName("Login")->Value;
 if(Logins->Locate("Login",Log,SO))
 {
  int Role=Logins->FieldByName("Role")->Value;
  int TempRole=TempLogins->FieldByName("Role")->Value;
  if(Role==4 & TempRole!=Role)
  {
   Logins->Delete();
  }
 }
}

for( Logins->First();!Logins->Eof;Logins->Next())
{
 int N=Logins->FieldByName("Num")->AsInteger;
 if(TempLogins->Locate("ServerNum",N,SO))
 {
Logins->Edit();
Logins->FieldByName("Login")->Value=TempLogins->FieldByName("Login")->Value;
Logins->FieldByName("Code1")->Value=TempLogins->FieldByName("Code1")->Value;
Logins->FieldByName("Code2")->Value=TempLogins->FieldByName("Code2")->Value;
Logins->FieldByName("Role")->Value=TempLogins->FieldByName("Role")->Value;
Logins->Post();

TempLogins->Edit();
TempLogins->FieldByName("Del")->Value=true;
TempLogins->Post();
 }
 else
 {
  Logins->Edit();
  Logins->FieldByName("Del")->Value=true;
  Logins->Post();
 }
}

Comm->CommandText="DELETE Logins.Del FROM Logins WHERE (((Logins.Del)=True));";
Comm->Execute();

TempLogins->Active=false;
TempLogins->CommandText="Select * From TempLogins Where Del=false";
TempLogins->Active=true;


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

TempLogins->Edit();
TempLogins->FieldByName("ServerNum")->Value=Logins->FieldByName("Num")->Value;
TempLogins->Post();
}


Comm->CommandText="Delete * from ObslOtdel";
Comm->Execute();

TempLogins->Active=false;
TempLogins->CommandText="Select * From TempLogins";
TempLogins->Active=true;
int Log2;
int Otd2;
for(TempObslOtdel->First();!TempObslOtdel->Eof;TempObslOtdel->Next())
{
 int Log1=TempObslOtdel->FieldByName("Login")->Value;
 if(TempLogins->Locate("Num", Log1, SO))
 {
  Log2=TempLogins->FieldByName("ServerNum")->Value;
 }
 else
 {
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка объединения таблицы ObslOtdel по логинам", "IDC="+IntToStr(IDC())+" Log1="+IntToStr(Log1)+" Log2="+IntToStr(Log2));
 }

 int Otd1=TempObslOtdel->FieldByName("NumObslOtdel")->Value;
 if(TempPodr->Locate("Номер подразделения", Otd1, SO))
 {
  Otd2=TempPodr->FieldByName("ServerNum")->Value;

  TempObslOtdel->Edit();
  TempObslOtdel->FieldByName("Login")->Value=Log2;
  TempObslOtdel->FieldByName("NumObslOtdel")->Value=Otd2;
  TempObslOtdel->Post();
 }
 else
 {
  TempObslOtdel->Edit();
  TempObslOtdel->FieldByName("Del")->Value=True;
  TempObslOtdel->Post();

DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка объединения таблицы ObslOtdel по подразделениям", "IDC="+IntToStr(IDC())+" Log1="+IntToStr(Otd1)+" Log2="+IntToStr(Otd2));

 }
}

Comm->CommandText="DELETE * From TempObslOtdel where Del=true";
Comm->Execute();

Comm->CommandText="DELETE * From ObslOtdel";
Comm->Execute();

Comm->CommandText="DELETE * From TempПодразделения";
Comm->Execute();

Comm->CommandText="INSERT INTO ObslOtdel ( Login, NumObslOtdel ) SELECT TempObslOtdel.Login, TempObslOtdel.NumObslOtdel FROM TempObslOtdel;";
Comm->Execute();

return true;
}
else
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка объединения логинов (нет базы данных)", "IDC="+IntToStr(IDC())+" IDDB="+IntToStr(IDDB));

return false;
}
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка объединения логинов", "IDC="+IntToStr(IDC())+" IDDB="+IntToStr(IDDB));

return false;
}
}
//*************************************************
bool Client::GetDatabaseConnect(long IDDB)
{
try
{
for(unsigned int i=0;i<VDatabase.size();i++)
{
 if(VDatabase[i]->IDDB==IDDB)
 {
  return VDatabase[i]->Database->Connected;
 }
}
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка чтения состояния базы (нет базы данных)", "IDC="+IntToStr(IDC())+" IDDB="+IntToStr(IDDB));

return false;
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка чтения состояния базы", "IDC="+IntToStr(IDC())+" IDDB="+IntToStr(IDDB));

return false;
}
}
//**********************************************
bool Client::SetDatabaseConnect(long IDDB, bool Connect)
{
try
{
for(unsigned int i=0;i<VDatabase.size();i++)
{
 if(VDatabase[i]->IDDB==IDDB)
 {
  VDatabase[i]->Database->Connected=Connect;
  return true;
 }
}
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка подключения базы (нет базы данных)", "IDC="+IntToStr(IDC())+" IDDB="+IntToStr(IDDB));

return false;
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка подключения базы", "IDC="+IntToStr(IDC())+" IDDB="+IntToStr(IDDB));

return false;
}
}
//**********************************************
long Client::CreateTable(long IDDB, long IDF)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Служебное", "Создание таблицы","IDC="+IntToStr(IDC)+" IDDB="+IntToStr(IDDB)+" IDF="+IntToStr(IDF));

try
{
TADOConnection * Conn=0;
for(unsigned int i=0;i<VDatabase.size();i++)
{
 if(VDatabase[i]->IDDB==IDDB)
 {
  Conn=VDatabase[i]->Database;
 }
}
if(Conn!=0)
{
long Ret=GetForm(IDF)->CreateTable( Conn);
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебное", "Создание таблицы", "IDC="+IntToStr(IDC())+" IDDB="+IntToStr(IDDB)+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(Ret));

  return Ret;

}
else
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка создания таблицы (нет базы данных)", "IDC="+IntToStr(IDC())+" IDDB="+IntToStr(IDDB)+" IDF="+IntToStr(IDF));

 return NULL;
}
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка создания таблицы", "IDC="+IntToStr(IDC())+" IDDB="+IntToStr(IDDB)+" IDF="+IntToStr(IDF));

 return NULL;
}
}
//**********************************************
bool Client::DeleteTable(long IDF, long IDT)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебное", "Удаление таблицы", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT));

try
{
  return GetForm(IDF)->DeleteTable(IDT);
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка удаления таблицы", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT));

 return false;
}
}
//**********************************************

void Client::SetCommandText(String Text,long IDF, long IDT)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебное", "SetCommandText", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT)+" Text: "+Text);

try
{
GetForm(IDF)->SetCommandText(Text, IDT);
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка установки CommandText", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT)+"Text: "+Text);

}
}
//**********************************************
String Client::GetCommandText(long IDF, long IDT)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебное", "GetCommandText", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT));

try
{
return GetForm(IDF)->GetCommandText(IDT);
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка чтения CommandText", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT));

 return " ";
}
}
//**********************************************
void Client::Active(bool Value, long IDF, long IDT)
{
if(Value)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебное", "Активация таблицы", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT)+" Value: True");
}
else
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебное", "Активация таблицы", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT)+" Value: False");
}

try
{
GetForm(IDF)->Active(Value, IDT);
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка активации таблицы", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT));

Form1->Block=-1;
}
}
//***********************************************
void Client::Moving(long Comand, long IDF, long IDT)
{
try
{
GetForm(IDF)->Moving(Comand, IDT);
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка перемещения по таблице", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT));

Form1->Block=-1;
}
}
//************************************************
bool Client::BOF_EOF (long Comand, long IDF, long IDT)
{
return GetForm(IDF)->BOF_EOF(Comand, IDT);
}
//************************************************
Variant Client::FieldByName(String FieldName, long IDF, long IDT)
{
try
{
return GetForm(IDF)->FieldByName(FieldName, IDT);
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка чтения поля по имени", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT)+" FieldName: "+FieldName);

 return NULL;
}
}
//*************************************************
void Client::FieldByName(Variant Value, String FieldName, long IDF, long IDT)
{
try
{
GetForm(IDF)->FieldByName(Value, FieldName, IDT);
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка записи поля по имени", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT)+" FieldName: "+FieldName+" Value="+(String)Value);

Form1->Block=-1;
}
}
//*************************************************

void Client::EditingRecord(long Comand, long IDF, long IDT)
{
try
{
GetForm(IDF)->EditingRecord(Comand, IDT);
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка редактирования записи", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT));

Form1->Block=-1;
}
}
//*************************************************
long Client::RecordCount(long IDF, long IDT)
{
return GetForm(IDF)->RecordCount(IDT);
}
//*************************************************
long Client::RecNo(long IDF, long IDT)
{
return GetForm(IDF)->RecNo(IDT);
}
//*************************************************
void Client::Move(long Step, long IDF, long IDT)
{
GetForm(IDF)->Move(Step, IDT);
}
//*************************************************
bool Client::Locate(long IDF, long IDT, String FieldName, String Key, TLocateOptions SearchOptions)
{
return GetForm(IDF)->Locate(IDT, FieldName, Key, SearchOptions);
}
//*************************************************
void Client::Delete(long IDF, long IDT)
{
GetForm(IDF)->Delete(IDT);
}
//************************************************
Variant Client::Fields(long NumField, long IDF, long IDT)
{
try
{
return GetForm(IDF)->Fields(NumField, IDT);
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка чтения поля по номеру", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT)+" Num: "+IntToStr(NumField));

 return NULL;
}
}
//************************************************
void Client::Fields(Variant Value, long NumField, long IDF, long IDT)
{
try
{
GetForm(IDF)->Fields(Value, NumField, IDT);
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка записи поля по номеру", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT)+" Num: "+IntToStr(NumField)+" Value="+(String)Value);

Form1->Block=-1;
}
}
//*************************************************
long Client::FieldsCount(long IDF, long IDT)
{
try
{
return GetForm(IDF)->FieldsCount(IDT);
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка чтения числа записей", "IDC="+IntToStr(IDC())+" IDF="+IntToStr(IDF)+" IDT="+IntToStr(IDT));

 return -1;
}
}
//*************************************************
bool Client::PrepareReadMetod(String NameDatabase)
{
bool Ret=false;
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Подготовка к чтению методики","IDC="+IntToStr(IDC));
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Главспец", "Подготовка к чтению методики", "IDC="+IntToStr(IDC())+" Database="+NameDatabase);

try
{
String FN=Form1->GetFileDatabase(NameDatabase);

TADOConnection *Database=NULL;

for(unsigned int i=0;i<VDatabase.size();i++)
{
 if(VDatabase[i]->FileName==FN)
 {
  Database=VDatabase[i]->Database;
  break;
 }
}

if(Database!=NULL)
{
MP<TADOCommand>Comm(Form1);
Comm->Connection=Database;
Comm->CommandText="Delete * from TempMet";
Comm->Execute();

MP<TADODataSet>Tab(Form1);
Tab->Connection=Database;
Tab->CommandText="Select * from TempMet Order by Номер";
Tab->Active=true;

MP<TADODataSet>Tab1(Form1);
Tab1->Connection=Database;
Tab1->CommandText="Select * from Методика Order by Номер";
Tab1->Active=true;
Tab1->First();

MP<TDataSource>DS(Form1);
DS->DataSet=Tab1;
DS->Enabled=true;

MP<TDBMemo>TDBM(Form1);
TDBM->Parent=Form1;
TDBM->Visible=false;
TDBM->Width=1009;
TDBM->DataSource=DS;
TDBM->DataField="Методика";


TStrings* TS=TDBM->Lines;
int N=TS->Count;
for(int i=0;i<N;i++)
{
String S=TS->Strings[i];
Tab->Insert();
Tab->FieldByName("Номер")->Value=i;
Tab->FieldByName("Методика")->Value=S;
Tab->Post();
}

Ret=true;
}



return Ret;
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка подготовки чтения методики", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase);

 return false;
}
}
//***********************************************
bool Client::MergeNodeBranch(String NameDatabase1, String NameNode, String NameBranch, String NameDatabase2, String NameTable, String NameField, String KeyName, String Name)
{//                            база справочников   |     узел     |       ветвь      |      база аспектов   |     таблица      | поле аспектов | ключ целевой таблицы | Наименование в целевой таблице
//int NumDatabase1=Form1->GetNumDatabase(NameDatabase1);
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Объединение узлов и ветвей (Full)","IDC="+IntToStr(IDC)+" Узел: "+NameNode+" Ветвь: "+NameBranch);
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Главспец", "Объединение узлов и ветвей (Full)", "IDC="+IntToStr(IDC())+" Узел: "+NameNode+" Ветвь: "+NameBranch);

try
{
TADOConnection *Database1=GetDBConnection(NameDatabase1);

if(Database1!=NULL)
{
MP<TADODataSet>TempNode(Form1);
TempNode->Connection=Database1;
TempNode->CommandText="Select * From TempNode";
TempNode->Active=true;

MP<TADODataSet>Node(Form1);
Node->Connection=Database1;
Node->CommandText="Select * From "+NameNode;
Node->Active=true;

MP<TADOCommand>Comm(Form1);
Comm->Connection=Database1;
Comm->CommandText="UPDATE "+NameNode+" SET "+NameNode+".Del = False, "+NameNode+".CopyNum = -1;";
Comm->Execute();

for(Node->First();!Node->Eof;Node->Next())
{
int NumNode=Node->FieldByName("Номер узла")->Value;

if(TempNode->Locate("Номер узла", NumNode, SO))
{
//Найдено
Node->Edit();
Node->FieldByName("Родитель")->Value=TempNode->FieldByName("Родитель")->Value;
Node->FieldByName("Название")->Value=TempNode->FieldByName("Название")->Value;
Node->Post();

TempNode->Delete();
}
else
{
//Ненайдено
Node->Edit();
Node->FieldByName("Del")->Value=true;
Node->Post();
}
}

Comm->CommandText="DELETE "+NameNode+".Del FROM "+NameNode+" WHERE ((("+NameNode+".Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO "+NameNode+" ( CopyNum, Родитель, Название ) SELECT TempNode.[Номер узла], TempNode.Родитель, TempNode.Название FROM TempNode;";
Comm->Execute();


MP<TADODataSet>TempBranch(Form1);
TempBranch->Connection=Database1;
TempBranch->CommandText="Select * From TempBranch";
TempBranch->Active=true;

MP<TADODataSet>Branch(Form1);
Branch->Connection=Database1;
Branch->CommandText="Select * From "+NameBranch;
Branch->Active=true;

Comm->CommandText="UPDATE "+NameBranch+" SET "+NameBranch+".Del = False, "+NameBranch+".NumCopy = -1; ";
Comm->Execute();
Node->Active=false;
Node->Active=true;

for(Branch->First();!Branch->Eof;Branch->Next())
{
 int NumBranch=Branch->FieldByName("Номер ветви")->Value;

 if(TempBranch->Locate("Номер ветви",NumBranch,SO))
 {
  Branch->Edit();
  int NumPar=TempBranch->FieldByName("Номер родителя")->Value;
  if(Node->Locate("CopyNum",NumPar,SO))
  {
  NumPar=Node->FieldByName("Номер узла")->Value;
  }

  Branch->FieldByName("Номер родителя")->Value=NumPar;
  Branch->FieldByName("Название")->Value=TempBranch->FieldByName("Название")->Value;
  Branch->FieldByName("Показ")->Value=TempBranch->FieldByName("Показ")->Value;
  Branch->Post();

  TempBranch->Delete();
 }
 else
 {
  Branch->Edit();
  Branch->FieldByName("Del")->Value=true;
  Branch->Post();
 }
}

Comm->CommandText="DELETE "+NameBranch+".Del FROM "+NameBranch+" WHERE ((("+NameBranch+".Del)=True));";
Comm->Execute();

for(TempBranch->First();!TempBranch->Eof;TempBranch->Next())
{
 int NumTemp=TempBranch->FieldByName("Номер родителя")->Value;

 if(Node->Locate("CopyNum", NumTemp, SO))
 {
  int NumPar=Node->FieldByName("Номер узла")->Value;

  Branch->Insert();
  Branch->FieldByName("Номер родителя")->Value=NumPar;
  Branch->FieldByName("Название")->Value=TempBranch->FieldByName("Название")->Value;
  Branch->FieldByName("Показ")->Value=TempBranch->FieldByName("Показ")->Value;
  Branch->Post();
 }
 else
 {
 if(Node->Locate("Номер узла", NumTemp, SO))
 {
  int NumPar=Node->FieldByName("Номер узла")->Value;

  Branch->Insert();
  Branch->FieldByName("Номер родителя")->Value=NumPar;
  Branch->FieldByName("Название")->Value=TempBranch->FieldByName("Название")->Value;
  Branch->FieldByName("Показ")->Value=TempBranch->FieldByName("Показ")->Value;
  Branch->Post();
 }
 else
 {
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка записи структуры", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase1+" Номер узла: "+IntToStr(NumTemp)+"NameNode "+NameNode+" NameBravch "+NameBranch);

 }
 }
}
}
else
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка записи узлов и ветвей 1 (Full) (нет базы данных)", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase1);

}




TADOConnection *Database2=GetDBConnection(NameDatabase2);



if(Database2!=NULL)
{
//План таков:
//На всякий случай снимаем все пометки в целевой таблице
//проходим по таблице ветви в Reference и ищем совпадение номера в целевой таблице базы аспектов
//если находим - обновляем название и помечаем запись в целевой таблице
//если не находим - добавляем запись в целевую таблицу и помечаем
//после окончания прохода все непомеченные записи целевой таблицы уже отсутствуют в новой версии справочника.
//запросом на исправление в таблице аспектов заменяем индекс ссылки на непомеченые записи нулями.
//Потом удаляем непомеченые записи целевой и снимаем пометки целевой таблицы.

//Database1 (Comm) - Reference
//Database2 (Comm2) - аспекты

MP<TADOCommand>Comm2(Form1);
Comm2->Connection=Database2;
Comm2->CommandText="UPDATE "+NameTable+" SET "+NameTable+".Del = False;";
Comm2->Execute();

MP<TADODataSet>Branch(Form1);
Branch->Connection=Database1;
Branch->CommandText="Select * From "+NameBranch+" Where показ=true Order by [Номер ветви]";
Branch->Active=true;

MP<TADODataSet>CurrTable(Form1);
CurrTable->Connection=Database2;
CurrTable->CommandText="Select * from "+NameTable+" Where показ=true";
CurrTable->Active=true;


for(Branch->First();!Branch->Eof;Branch->Next())
{
 int NumBranch=Branch->FieldByName("Номер ветви")->Value;
 if(CurrTable->Locate(KeyName, NumBranch, SO))
 {
  //Найдено
  CurrTable->Edit();
  CurrTable->FieldByName(Name)->Value=Branch->FieldByName("Название")->Value;
  CurrTable->FieldByName("Показ")->Value=Branch->FieldByName("Показ")->Value;
  CurrTable->FieldByName("Del")->Value=true;
  CurrTable->Post();
 }
 else
 {
  //Ненайдено
  CurrTable->Insert();
  CurrTable->FieldByName(KeyName)->Value=Branch->FieldByName("Номер ветви")->Value;
  CurrTable->FieldByName(Name)->Value=Branch->FieldByName("Название")->Value;
  CurrTable->FieldByName("Показ")->Value=Branch->FieldByName("Показ")->Value;
  CurrTable->FieldByName("Del")->Value=true;
  CurrTable->Post();
 }
}

Comm2->CommandText="UPDATE ["+NameTable+"] INNER JOIN Аспекты ON ["+NameTable+"].["+KeyName+"] = Аспекты.["+NameField+"] SET Аспекты.["+NameField+"] = 0 WHERE (((["+NameTable+"].Показ)=True) AND ((["+NameTable+"].Del)=False)); ";
//Comm2->CommandText="UPDATE "+NameTable+" INNER JOIN Аспекты ON "+NameTable+".[Номер воздействия] = Аспекты."+NameField+" SET Аспекты."+NameField+" = 0, "+NameTable+".Показ = True WHERE ((("+NameTable+".Del)=False)); ";
//                  UPDATE Территория INNER JOIN Аспекты ON Территория.[Номер территории] = Аспекты.[Вид территории] SET Аспекты.[Вид территории] = 0 WHERE (((Территория.Показ)=True) AND ((Территория.Del)=False));
Comm2->Execute();

Comm2->CommandText="Delete * From "+NameTable+" Where "+NameTable+".Del=false AND "+NameTable+".Показ=true";
Comm2->Execute();

Comm2->CommandText="UPDATE "+NameTable+" SET "+NameTable+".Del = False";
Comm2->Execute();

MP<TADOCommand>Comm(Form1);
Comm->Connection=Database1;
Comm->CommandText="Delete * From TempBranch";
Comm->Execute();
return true;
}
else
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Сбой", "Сбой записи узлов и ветвей (Full) (нет базы данных)", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase2);

 return false;
}

}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка записи узлов и ветвей (Full)", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase1+" DB2: "+NameDatabase2);

 return false;
}
}
//***********************************************
bool Client::MergeNodeBranch(String NameDatabase1, String NameNode, String NameBranch)
{//                            база справочников   |     узел     |       ветвь      |      база аспектов   |     таблица      | поле аспектов | ключ целевой таблицы | Наименование в целевой таблице
//int NumDatabase1=Form1->GetNumDatabase(NameDatabase1);
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Объединение узлов и ветвей (Low)","IDC="+IntToStr(IDC)+" Узел: "+NameNode+" Ветвь: "+NameBranch);
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Главспец", "Объединение узлов и ветвей (Low)", "IDC="+IntToStr(IDC())+" Узел: "+NameNode+" Ветвь: "+NameBranch);

try
{
TADOConnection *Database1=GetDBConnection(NameDatabase1);



if(Database1!=NULL)
{
MP<TADOCommand>Comm(Form1);
Comm->Connection=Database1;
Comm->CommandText="Delete * From "+NameNode;
Comm->Execute();

Comm->CommandText="INSERT INTO "+NameNode+" ( [Номер узла], Родитель, Название ) SELECT TempNode.[Номер узла], TempNode.Родитель, TempNode.Название FROM TempNode;";
Comm->Execute();

Comm->CommandText="INSERT INTO "+NameBranch+" ( [Номер ветви], [Номер родителя], Название, Показ ) SELECT TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.Название, TempBranch.Показ FROM TempBranch;";
Comm->Execute();

Comm->CommandText="Delete * From TempNode";
Comm->Execute();

return true;
}
else
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Сбой", "Сбой записи узлов и ветвей (Low) (нет базы данных)", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase1);

return false;
}

}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка записи узлов и ветвей (Low)", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase1);

 return false;
}
}

//***********************************************
void Client::MergeZn(String DB1, String DB2)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Объединение критериев","IDC="+IntToStr(IDC)+" DB1: "+Database1+" DB2: "+Database2);
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Главспец", "Объединение критериев", "IDC="+IntToStr(IDC())+" DB1: "+DB1+" DB2: "+DB2);


try
{
//Reference
TADOConnection *Database1=GetDBConnection(DB1);


MP<TADOCommand>Comm1(Form1);
Comm1->Connection=Database1;
Comm1->CommandText="Delete * From Значимость";
Comm1->Execute();

Comm1->CommandText="INSERT INTO Значимость ( [Номер значимости], [Наименование значимости], Критерий1, Критерий, [Мин граница], [Макс граница], [Необходимая мера] ) SELECT TempZn.[Номер значимости], TempZn.[Наименование значимости], TempZn.Критерий1, TempZn.Критерий, TempZn.[Мин граница], TempZn.[Макс граница], TempZn.[Необходимая мера] FROM TempZn;";
Comm1->Execute();

//Aspects
TADOConnection *Database2=GetDBConnection(DB2);

MP<TADOCommand>Comm2(Form1);
Comm2->Connection=Database2;
Comm2->CommandText="Delete * From Значимость";
Comm2->Execute();

MP<TADODataSet>Tab1(Form1);
Tab1->Connection=Database1;
Tab1->CommandText="Select * From Значимость order by [Номер значимости]";
Tab1->Active=true;

MP<TADODataSet>Tab2(Form1);
Tab2->Connection=Database2;
Tab2->CommandText="Select * From Значимость";
Tab2->Active=true;

for(Tab1->First();!Tab1->Eof;Tab1->Next())
{
 Tab2->Insert();
 Tab2->FieldByName("Номер значимости")->Value=Tab1->FieldByName("Номер значимости")->Value;
 Tab2->FieldByName("Наименование значимости")->Value=Tab1->FieldByName("Наименование значимости")->Value;
 Tab2->FieldByName("Критерий1")->Value=Tab1->FieldByName("Критерий1")->Value;
 Tab2->FieldByName("Критерий")->Value=Tab1->FieldByName("Критерий")->Value;
 Tab2->FieldByName("Мин граница")->Value=Tab1->FieldByName("Мин граница")->Value;
 Tab2->FieldByName("Макс граница")->Value=Tab1->FieldByName("Макс граница")->Value;
 Tab2->FieldByName("Необходимая мера")->Value=Tab1->FieldByName("Необходимая мера")->Value;
 Tab2->Post();
}

Comm1->CommandText="Delete * From TempZn";
Comm1->Execute();
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка объединения значимостей", "IDC="+IntToStr(IDC())+" DB1: "+DB1+" DB2: "+DB2);

Form1->Block=-1;
}
}
//***********************************************
void Client::SaveMetod(String Db)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Запись методики","IDC="+IntToStr(IDC)+" DB: "+Database);
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Главспец", "Запись методики", "IDC="+IntToStr(IDC())+" DB: "+Db);

try
{
TADOConnection *Database=GetDBConnection(Db);


MP<TADODataSet>TempMet(Form1);
TempMet->Connection=Database;
TempMet->CommandText="Select TempMet.[Номер], TempMet.[Методика] From TempMet order by Номер";
TempMet->Active=true;

MP<TADOCommand>Comm(Form1);
Comm->Connection=Database;
Comm->CommandText="Delete * from  Методика";
Comm->Execute();

MP<TADODataSet>Metod(Form1);
Metod->Connection=Database;
Metod->CommandText="Select * from Методика ";
Metod->Active=true;

MP<TMemo>M(Form1);
M->Visible=false;
M->Parent=Form1;

TStrings* TT=M->Lines;
TT->Clear();
for(TempMet->First();!TempMet->Eof;TempMet->Next())
{
String S=TempMet->FieldByName("методика")->AsString;
TT->Append(S);
}
Metod->Insert();
Metod->FieldByName("Номер")->Value=1;
Metod->FieldByName("Методика")->Assign(TT);
Metod->Post();
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка записи методики", "IDC="+IntToStr(IDC())+" DB: "+Db);

Form1->Block=-1;
}
}
//***********************************************
void Client::SavePodr(String DB1)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Запись подразделений","IDC="+IntToStr(IDC)+" DB1: "+DB1);
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Главспец", "Запись подразделений", "IDC="+IntToStr(IDC())+" DB1: "+DB1);


try
{
/*
TADOConnection *Database1=GetDBConnection(DB1);

MP<TADOCommand>Comm1(Form1);
Comm1->Connection=Database1;
Comm1->CommandText="Delete * From подразделения";
Comm1->Execute();

Comm1->CommandText="INSERT INTO Подразделения ( [Номер подразделения], [Название подразделения] ) SELECT TempPodr.[Номер подразделения], TempPodr.[Название подразделения] FROM TempPodr;";
Comm1->Execute();
*/
TADOConnection *Database2=GetDBConnection(DB1);

MP<TADOCommand>Comm2(Form1);
Comm2->Connection=Database2;
Comm2->CommandText="UPDATE Подразделения SET Подразделения.Del = False;";
Comm2->Execute();
Comm2->CommandText="UPDATE TempПодразделения SET TempПодразделения.Del = False;";
Comm2->Execute();

MP<TADODataSet>Temp(Form1);
Temp->Connection=Database2;
Temp->CommandText="Select * From TempПодразделения Order by [Номер подразделения]";
Temp->Active=true;

MP<TADODataSet>Podr(Form1);
Podr->Connection=Database2;
Podr->CommandText="Select * From Подразделения Order by [Номер подразделения]";
Podr->Active=true;

for(Podr->First();!Podr->Eof;Podr->Next())
{
int Num=Podr->FieldByName("Номер подразделения")->AsInteger;

if(Temp->Locate("ServerNum", Num, SO))
{
 //Найдено
 Podr->Edit();
 Podr->FieldByName("Название подразделения")->Value=Temp->FieldByName("Название подразделения")->Value;
 Podr->Post();

 Temp->Edit();
 Temp->FieldByName("Del")->Value=true;
 Temp->Post();

}
else
{
 //Ненайдено
 Podr->Edit();
 Podr->FieldByName("Del")->Value=true;
 Podr->Post();
}
}

Comm2->CommandText="Delete * From подразделения Where Del=true";
Comm2->Execute();

Temp->Active=false;
Temp->CommandText="Select * From TempПодразделения where Del=false Order by [Номер подразделения]";
Temp->Active=true;

for(Temp->First();!Temp->Eof;Temp->Next())
{
Podr->Insert();
Podr->FieldByName("Название подразделения")->Value=Temp->FieldByName("Название подразделения")->Value;
Podr->Post();
Podr->Active=false;
Podr->Active=true;
Podr->Last();
Temp->Edit();
Temp->FieldByName("ServerNum")->Value=Podr->FieldByName("Номер подразделения")->Value;
Temp->Post();
}



/*
Comm2->CommandText="INSERT INTO Подразделения ( NumMSp, [Название подразделения] ) SELECT TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения] FROM TempПодразделения;";
Comm2->Execute();

Temp->CommandText="Select * From TempПодразделения Order by [Номер подразделения]";
Temp->Active=true;
*/
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка записи подразделений", "IDC="+IntToStr(IDC())+" DB1: "+DB1);

Form1->Block=-1;
}
}
//***********************************************
void Client::SaveSit(String DB1, String DB2)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Запись ситуаций","IDC="+IntToStr(IDC)+" DB1: "+DB1+" DB2: "+DB2);
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Главспец", "Запись ситуаций", "IDC="+IntToStr(IDC())+" DB1: "+DB1+" DB2: "+DB2);

try
{

TADOConnection *Database1=GetDBConnection(DB1);

MP<TADOCommand>Comm1(Form1);
Comm1->Connection=Database1;
Comm1->CommandText="Delete * From Ситуации";
Comm1->Execute();

Comm1->CommandText="INSERT INTO Ситуации ( [Номер ситуации], [Название ситуации] ) SELECT TempSit.[Номер ситуации], TempSit.[Название ситуации] FROM TempSit;";
Comm1->Execute();

TADOConnection *Database2=GetDBConnection(DB2);

MP<TADOCommand>Comm2(Form1);
Comm2->Connection=Database2;
Comm2->CommandText="UPDATE Ситуации SET Ситуации.Del = False;";
Comm2->Execute();

MP<TADODataSet>Temp(Form1);
Temp->Connection=Database2;
Temp->CommandText="Select * From TempSit Order by [Номер ситуации]";
Temp->Active=true;

MP<TADODataSet>Sit(Form1);
Sit->Connection=Database2;
Sit->CommandText="Select * From Ситуации Where [Номер ситуации]<>0 Order by [Номер ситуации]";
Sit->Active=true;

for(Sit->First();!Sit->Eof;Sit->Next())
{
int Num=Sit->FieldByName("номер ситуации")->Value;

if(Temp->Locate("Номер ситуации", Num, SO))
{
 //Найдено
 Sit->Edit();
 Sit->FieldByName("Номер ситуации")->Value=Temp->FieldByName("Номер ситуации")->Value;
 Sit->FieldByName("Название ситуации")->Value=Temp->FieldByName("Название ситуации")->Value;
 Sit->FieldByName("Показ")->Value=true;
 Sit->Post();

 Temp->Delete();
}
else
{
 //Ненайдено
 Sit->Edit();
 Sit->FieldByName("Del")->Value=true;
 Sit->Post();
}
}
MP<TADODataSet>Asp(Form1);
Asp->Connection=Database2;
Asp->CommandText="SELECT Аспекты.* FROM Ситуации INNER JOIN Аспекты ON Ситуации.[Номер ситуации] = Аспекты.Ситуация WHERE (((Ситуации.Del)=True));";
Asp->Active=true;

for(Asp->First();!Asp->Eof;Asp->Next())
{
 Asp->Edit();
 Asp->FieldByName("Ситуация")->Value=0;
 Asp->Post();
}

Comm2->CommandText="Delete * From Ситуации Where Del=true AND Показ=True";
Comm2->Execute();

Comm2->CommandText="INSERT INTO Ситуации ( [Номер ситуации], [Название ситуации], [показ] ) SELECT TempSit.[Номер ситуации], TempSit.[Название ситуации], true FROM TempSit;";
Comm2->Execute();

//Temp->CommandText="Select * From TempSit Order by [Номер ситуации]";
//Temp->Active=true;

Comm2->CommandText="Delete * From TempSit";
Comm2->Execute();
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка записи ситуаций", "IDC="+IntToStr(IDC())+" DB1: "+DB1+" DB2: "+DB2);

Form1->Block=-1;
}
}
//***********************************************
TADOConnection* Client::GetDBConnection(String NameDatabase)
{
try
{
String FN=Form1->GetFileDatabase(NameDatabase);

TADOConnection *Database=NULL;

for(unsigned int i=0;i<VDatabase.size();i++)
{
 if(VDatabase[i]->FileName==FN)
 {
  Database=VDatabase[i]->Database;
  break;
 }
}
return Database;
}
catch(...)
{
Form1->Block=-1;
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка получения подключения к базе данных", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase);

 return NULL;
}
}
//************************************************
void Client::MergeAspectsMainSpec()
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Объединение аспектов (главспец)","IDC="+IntToStr(IDC));
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Главспец", "Объединение аспектов (главспец)", "IDC="+IntToStr(IDC()));

try
{
MP<TADOCommand>Comm(Form1);
Comm->Connection=GetDBConnection("Аспекты");
Comm->CommandText="UPDATE Аспекты SET Аспекты.Del = False";
Comm->Execute();

MP<TADODataSet>Aspects(Form1);
Aspects->Connection=GetDBConnection("Аспекты");
Aspects->CommandText="Select * From Аспекты";
Aspects->Active=true;

MP<TADODataSet>Temp(Form1);
Temp->Connection=GetDBConnection("Аспекты");
Temp->CommandText="Select * From TempAspects";
Temp->Active=true;

for(Aspects->First();!Aspects->Eof;Aspects->Next())
{
 int N=Aspects->FieldByName("номер аспекта")->Value;

 if(Temp->Locate("Номер аспекта", N, SO))
 {
  Aspects->Edit();
  Aspects->FieldByName("Подразделение")->Value=Temp->FieldByName("Подразделение")->Value;
  Aspects->Post();
 }
 else
 {
//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка объединения аспектов главным специалистом (не найден аспект)", "IDC="+IntToStr(IDC())+" Номер="+IntToStr(N));

 }
}
/*
MP<TADOCommand>Comm(Form1);
Comm->Connection=GetDBConnection("Аспекты");
Comm->CommandText="UPDATE Аспекты SET Аспекты.Del = False";
Comm->Execute();

MP<TADODataSet>Podr(Form1);
Podr->Connection=GetDBConnection("Аспекты");
Podr->CommandText="Select * From Подразделения";
Podr->Active=true;

MP<TADODataSet>Aspects(Form1);
Aspects->Connection=GetDBConnection("Аспекты");
Aspects->CommandText="Select * From Аспекты";
Aspects->Active=true;

MP<TADODataSet>Temp(Form1);
Temp->Connection=GetDBConnection("Аспекты");
Temp->CommandText="Select * From TempAspects";
Temp->Active=true;

MP<TADODataSet>Zn(Form1);
Zn->Connection=GetDBConnection("Аспекты");
Zn->CommandText="select * From Значимость Order by [Мин граница]";
Zn->Active=true;

for(Temp->First();!Temp->Eof;Temp->Next())
{
 int N=Temp->FieldByName("Подразделение")->Value;
 if(Podr->Locate("NumMSp",N,SO))
 {
  int Num=Podr->FieldByName("Номер подразделения")->Value;

  Temp->Edit();
  Temp->FieldByName("подразделение")->Value=Num;
  Temp->Post();
 }
 else
 {
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Ошибка обработки данных", "Ошибка объединения аспектов главным специалистом (не найдено подразделение)", "IDC="+IntToStr(IDC()));

 }
}

for(Aspects->First();!Aspects->Eof;Aspects->Next())
{
 int SNum=Aspects->FieldByName("Номер аспекта")->Value;

 if(Temp->Locate("Номер аспекта",SNum, SO))
 {
  //Найдено
  Aspects->Edit();

  Aspects->FieldByName("Подразделение")->Value=Temp->FieldByName("Подразделение")->Value;

  Aspects->Post();

  Temp->Delete();
 }
 else
 {
  //Ненайдено
  Aspects->Insert();
  Aspects->FieldByName("Del")->Value=true;
  Aspects->Post();
 }
}


Comm->CommandText="UPDATE Аспекты SET Аспекты.Del = False";
Comm->Execute();
*/
MP<TADODataSet>Zn(Form1);
Zn->Connection=GetDBConnection("Аспекты");
Zn->CommandText="select * From Значимость Order by [Мин граница]";
Zn->Active=true;

Aspects->Active=false;
Aspects->Active=true;

for(Aspects->First();!Aspects->Eof;Aspects->Next())
{
 double Z=Aspects->FieldByName("Z")->Value;
 bool Znak=false;

 for(Zn->First();!Zn->Eof;Zn->Next())
 {
  double Min=Zn->FieldByName("Мин граница")->Value;
  double Max=Zn->FieldByName("Макс граница")->Value;

  if(Z>=Min & Z<=Max)
  {
   Znak=Zn->FieldByName("Критерий")->Value;
   break;
  }
 }

 Aspects->Edit();
 Aspects->FieldByName("Значимость")->Value=Znak;
 Aspects->Post();
}
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка объединения аспектов главным специалистом", "IDC="+IntToStr(IDC()));

Form1->Block=-1;
}
}
//************************************************
void Client::MergeAspectsMainSpecH()
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Главспец", "Объединение аспектов (главспец)","IDC="+IntToStr(IDC));
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Главспец", "Объединение рисков (главспец)", "IDC="+IntToStr(IDC()));

try
{
MP<TADOCommand>Comm(Form1);
Comm->Connection=GetDBConnection("Опасности");
Comm->CommandText="UPDATE Аспекты SET Аспекты.Del = False";
Comm->Execute();

MP<TADODataSet>Aspects(Form1);
Aspects->Connection=GetDBConnection("Опасности");
Aspects->CommandText="Select * From Аспекты";
Aspects->Active=true;

MP<TADODataSet>Temp(Form1);
Temp->Connection=GetDBConnection("Опасности");
Temp->CommandText="Select * From TempAspects";
Temp->Active=true;

for(Aspects->First();!Aspects->Eof;Aspects->Next())
{
 int N=Aspects->FieldByName("номер аспекта")->Value;

 if(Temp->Locate("Номер аспекта", N, SO))
 {
  Aspects->Edit();
  Aspects->FieldByName("Подразделение")->Value=Temp->FieldByName("Подразделение")->Value;
  Aspects->Post();
 }
 else
 {
//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка объединения аспектов главным специалистом (не найден аспект)", "IDC="+IntToStr(IDC())+" Номер="+IntToStr(N));

 }
}
/*
MP<TADOCommand>Comm(Form1);
Comm->Connection=GetDBConnection("Аспекты");
Comm->CommandText="UPDATE Аспекты SET Аспекты.Del = False";
Comm->Execute();

MP<TADODataSet>Podr(Form1);
Podr->Connection=GetDBConnection("Аспекты");
Podr->CommandText="Select * From Подразделения";
Podr->Active=true;

MP<TADODataSet>Aspects(Form1);
Aspects->Connection=GetDBConnection("Аспекты");
Aspects->CommandText="Select * From Аспекты";
Aspects->Active=true;

MP<TADODataSet>Temp(Form1);
Temp->Connection=GetDBConnection("Аспекты");
Temp->CommandText="Select * From TempAspects";
Temp->Active=true;

MP<TADODataSet>Zn(Form1);
Zn->Connection=GetDBConnection("Аспекты");
Zn->CommandText="select * From Значимость Order by [Мин граница]";
Zn->Active=true;

for(Temp->First();!Temp->Eof;Temp->Next())
{
 int N=Temp->FieldByName("Подразделение")->Value;
 if(Podr->Locate("NumMSp",N,SO))
 {
  int Num=Podr->FieldByName("Номер подразделения")->Value;

  Temp->Edit();
  Temp->FieldByName("подразделение")->Value=Num;
  Temp->Post();
 }
 else
 {
Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Ошибка обработки данных", "Ошибка объединения аспектов главным специалистом (не найдено подразделение)", "IDC="+IntToStr(IDC()));

 }
}

for(Aspects->First();!Aspects->Eof;Aspects->Next())
{
 int SNum=Aspects->FieldByName("Номер аспекта")->Value;

 if(Temp->Locate("Номер аспекта",SNum, SO))
 {
  //Найдено
  Aspects->Edit();

  Aspects->FieldByName("Подразделение")->Value=Temp->FieldByName("Подразделение")->Value;

  Aspects->Post();

  Temp->Delete();
 }
 else
 {
  //Ненайдено
  Aspects->Insert();
  Aspects->FieldByName("Del")->Value=true;
  Aspects->Post();
 }
}


Comm->CommandText="UPDATE Аспекты SET Аспекты.Del = False";
Comm->Execute();
*/
MP<TADODataSet>Zn(Form1);
Zn->Connection=GetDBConnection("Опасности");
Zn->CommandText="select * From Значимость Order by [Мин граница]";
Zn->Active=true;

Aspects->Active=false;
Aspects->Active=true;

for(Aspects->First();!Aspects->Eof;Aspects->Next())
{
 double Z=Aspects->FieldByName("Z")->Value;
 bool Znak=false;

 for(Zn->First();!Zn->Eof;Zn->Next())
 {
  double Min=Zn->FieldByName("Мин граница")->Value;
  double Max=Zn->FieldByName("Макс граница")->Value;

  if(Z>=Min & Z<=Max)
  {
   Znak=Zn->FieldByName("Критерий")->Value;
   break;
  }
 }

 Aspects->Edit();
 Aspects->FieldByName("Значимость")->Value=Znak;
 Aspects->Post();
}
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка объединения рисков главным специалистом", "IDC="+IntToStr(IDC()));

Form1->Block=-1;
}
}
//************************************************
void Client::PrepareLoadAspects(String NameDatabase, long Podr, long Width1, long Width2, long Width3, long Width4)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Пользователи", "Подготовка к чтению аспектов","IDC="+IntToStr(IDC)+" DB: "+NameDatabase+" Podr="+IntToStr(Podr));
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Пользователи", "Подготовка к чтению аспектов", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase+" Podr="+IntToStr(Podr));

try
{
TADOConnection *Database=GetDBConnection(NameDatabase);
MP<TADOCommand>Comm(Form1);
Comm->Connection=Database;
Comm->CommandText="Delete * From TempAspects";
Comm->Execute();

String CT="";
if(Podr==0)
{
 //Все подразделения
 CT="INSERT INTO TempAspects ( [Номер аспекта], Подразделение, Ситуация, [Вид территории], Деятельность, Специальность, Аспект, Воздействие, G, O, R, S, T, L, N, Z, Значимость, [Проявление воздействия], [Тяжесть последствий], Приоритетность, [Выполняющиеся мероприятия], [Предлагаемые мероприятия], [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Исполнитель, [Дата создания], [Начало действия], [Конец действия] ) ";
CT=CT+" SELECT Аспекты.[Номер аспекта], Аспекты.Подразделение, Аспекты.Ситуация, Аспекты.[Вид территории], Аспекты.Деятельность, Аспекты.Специальность, Аспекты.Аспект, Аспекты.Воздействие, Аспекты.G, Аспекты.O, Аспекты.R, Аспекты.S, Аспекты.T, Аспекты.L, Аспекты.N, Аспекты.Z, Аспекты.Значимость, Аспекты.[Проявление воздействия], Аспекты.[Тяжесть последствий], Аспекты.Приоритетность, Аспекты.[Выполняющиеся мероприятия], Аспекты.[Предлагаемые мероприятия], Аспекты.[Мониторинг и контроль], Аспекты.[Предлагаемый мониторинг и контроль], Аспекты.Исполнитель, Аспекты.[Дата создания], Аспекты.[Начало действия], Аспекты.[Конец действия] ";
CT=CT+" FROM Аспекты;";
}
else
{
 //Одно подразделение

 CT="INSERT INTO TempAspects ( [Номер аспекта], Подразделение, Ситуация, [Вид территории], Деятельность, Специальность, Аспект, Воздействие, G, O, R, S, T, L, N, Z, Значимость, [Проявление воздействия], [Тяжесть последствий], Приоритетность, [Выполняющиеся мероприятия], [Предлагаемые мероприятия], [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Исполнитель, [Дата создания], [Начало действия], [Конец действия] ) ";
CT=CT+" SELECT Аспекты.[Номер аспекта], Аспекты.Подразделение, Аспекты.Ситуация, Аспекты.[Вид территории], Аспекты.Деятельность, Аспекты.Специальность, Аспекты.Аспект, Аспекты.Воздействие, Аспекты.G, Аспекты.O, Аспекты.R, Аспекты.S, Аспекты.T, Аспекты.L, Аспекты.N, Аспекты.Z, Аспекты.Значимость, Аспекты.[Проявление воздействия], Аспекты.[Тяжесть последствий], Аспекты.Приоритетность, Аспекты.[Выполняющиеся мероприятия], Аспекты.[Предлагаемые мероприятия], Аспекты.[Мониторинг и контроль], Аспекты.[Предлагаемый мониторинг и контроль], Аспекты.Исполнитель, Аспекты.[Дата создания], Аспекты.[Начало действия], Аспекты.[Конец действия] ";
CT=CT+" FROM Аспекты WHERE (((Аспекты.Подразделение)="+IntToStr(Podr)+"));";
}

Comm->CommandText=CT;
Comm->Execute();

MP<TADODataSet>Temp(Form1);
Temp->Connection=Database;
Temp->CommandText="Select * From TempAspects";
Temp->Active=true;

for(Temp->First();!Temp->Eof;Temp->Next())
{
 int Num=Temp->FieldByName("Номер аспекта")->Value;

MP<TDataSource>DS(Form1);
DS->DataSet=Temp;
DS->Enabled=true;

MP<TDBMemo>TDBM(Form1);
TDBM->Parent=Form1;
TDBM->Visible=false;
TDBM->Width=Width1;
TDBM->DataSource=DS;
TDBM->DataField="Выполняющиеся мероприятия";

TStrings* TS=TDBM->Lines;
int N=TS->Count;

MP<TADODataSet>Memo1(Form1);
Memo1->Connection=Database;
Memo1->CommandText="Select * From Memo1";
Memo1->Active=true;

for(int i=0;i<N;i++)
{
String S=TS->Strings[i];
Memo1->Insert();
Memo1->FieldByName("Num")->Value=Num;
Memo1->FieldByName("NumStr")->Value=i;
Memo1->FieldByName("Text")->Value=S;
Memo1->Post();
}

MP<TDBMemo>TDBM2(Form1);
TDBM2->Parent=Form1;
TDBM2->Visible=false;
TDBM2->Width=Width2;
TDBM2->DataSource=DS;
TDBM2->DataField="Предлагаемые мероприятия";

TStrings* TS2=TDBM2->Lines;
int N2=TS2->Count;

MP<TADODataSet>Memo2(Form1);
Memo2->Connection=Database;
Memo2->CommandText="Select * From Memo2";
Memo2->Active=true;

for(int i=0;i<N2;i++)
{
String S=TS2->Strings[i];
Memo2->Insert();
Memo2->FieldByName("Num")->Value=Num;
Memo2->FieldByName("NumStr")->Value=i;
Memo2->FieldByName("Text")->Value=S;
Memo2->Post();
}

MP<TDBMemo>TDBM3(Form1);
TDBM3->Parent=Form1;
TDBM3->Visible=false;
TDBM3->Width=Width3;
TDBM3->DataSource=DS;
TDBM3->DataField="Мониторинг и контроль";

TStrings* TS3=TDBM3->Lines;
int N3=TS3->Count;

MP<TADODataSet>Memo3(Form1);
Memo3->Connection=Database;
Memo3->CommandText="Select * From Memo3";
Memo3->Active=true;

for(int i=0;i<N3;i++)
{
String S=TS3->Strings[i];
Memo3->Insert();
Memo3->FieldByName("Num")->Value=Num;
Memo3->FieldByName("NumStr")->Value=i;
Memo3->FieldByName("Text")->Value=S;
Memo3->Post();
}

MP<TDBMemo>TDBM4(Form1);
TDBM4->Parent=Form1;
TDBM4->Visible=false;
TDBM4->Width=Width4;
TDBM4->DataSource=DS;
TDBM4->DataField="Предлагаемый мониторинг и контроль";

TStrings* TS4=TDBM4->Lines;
int N4=TS4->Count;

MP<TADODataSet>Memo4(Form1);
Memo4->Connection=Database;
Memo4->CommandText="Select * From Memo4";
Memo4->Active=true;

for(int i=0;i<N4;i++)
{
String S=TS4->Strings[i];
Memo4->Insert();
Memo4->FieldByName("Num")->Value=Num;
Memo4->FieldByName("NumStr")->Value=i;
Memo4->FieldByName("Text")->Value=S;
Memo4->Post();
}
}
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка подготовки к чтению аспектов", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase+" Подразделение: "+IntToStr(Podr));

Form1->Block=-1;
}
}
//*************************************************
void Client::PrepareLoadAspectsH(String NameDatabase, long Podr, long Width1, long Width2, long Width3, long Width4)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Пользователи", "Подготовка к чтению аспектов","IDC="+IntToStr(IDC)+" DB: "+NameDatabase+" Podr="+IntToStr(Podr));
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Пользователи", "Подготовка к чтению рисков", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase+" Podr="+IntToStr(Podr));

try
{
TADOConnection *Database=GetDBConnection(NameDatabase);
MP<TADOCommand>Comm(Form1);
Comm->Connection=Database;
Comm->CommandText="Delete * From TempAspects";
Comm->Execute();

String CT="";
if(Podr==0)
{
 //Все подразделения
 CT="INSERT INTO TempAspects ( [Номер аспекта], Подразделение, Ситуация, [Вид территории], Деятельность, Специальность, Аспект, Воздействие, G, O, R, S, T, L, N, Z, Значимость, [Наименование значимости], [Проявление воздействия], [Тяжесть последствий], Приоритетность, Приор, [Выполняющиеся мероприятия], [Предлагаемые мероприятия], [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Исполнитель, [Дата создания], [Начало действия], [Конец действия] ) ";
CT=CT+" SELECT Аспекты.[Номер аспекта], Аспекты.Подразделение, Аспекты.Ситуация, Аспекты.[Вид территории], Аспекты.Деятельность, Аспекты.Специальность, Аспекты.Аспект, Аспекты.Воздействие, Аспекты.G, Аспекты.O, Аспекты.R, Аспекты.S, Аспекты.T, Аспекты.L, Аспекты.N, Аспекты.Z, Аспекты.Значимость, [Наименование значимости], Аспекты.[Проявление воздействия], Аспекты.[Тяжесть последствий], Аспекты.Приоритетность, Приор, Аспекты.[Выполняющиеся мероприятия], Аспекты.[Предлагаемые мероприятия], Аспекты.[Мониторинг и контроль], Аспекты.[Предлагаемый мониторинг и контроль], Аспекты.Исполнитель, Аспекты.[Дата создания], Аспекты.[Начало действия], Аспекты.[Конец действия] ";
CT=CT+" FROM Аспекты;";
}
else
{
 //Одно подразделение

 CT="INSERT INTO TempAspects ( [Номер аспекта], Подразделение, Ситуация, [Вид территории], Деятельность, Специальность, Аспект, Воздействие, G, O, R, S, T, L, N, Z, Значимость, [Наименование значимости], [Проявление воздействия], [Тяжесть последствий], Приоритетность, Приор, [Выполняющиеся мероприятия], [Предлагаемые мероприятия], [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Исполнитель, [Дата создания], [Начало действия], [Конец действия] ) ";
CT=CT+" SELECT Аспекты.[Номер аспекта], Аспекты.Подразделение, Аспекты.Ситуация, Аспекты.[Вид территории], Аспекты.Деятельность, Аспекты.Специальность, Аспекты.Аспект, Аспекты.Воздействие, Аспекты.G, Аспекты.O, Аспекты.R, Аспекты.S, Аспекты.T, Аспекты.L, Аспекты.N, Аспекты.Z, Аспекты.Значимость, [Наименование значимости], Аспекты.[Проявление воздействия], Аспекты.[Тяжесть последствий], Аспекты.Приоритетность, Приор, Аспекты.[Выполняющиеся мероприятия], Аспекты.[Предлагаемые мероприятия], Аспекты.[Мониторинг и контроль], Аспекты.[Предлагаемый мониторинг и контроль], Аспекты.Исполнитель, Аспекты.[Дата создания], Аспекты.[Начало действия], Аспекты.[Конец действия] ";
CT=CT+" FROM Аспекты WHERE (((Аспекты.Подразделение)="+IntToStr(Podr)+"));";
}

Comm->CommandText=CT;
Comm->Execute();

MP<TADODataSet>Temp(Form1);
Temp->Connection=Database;
Temp->CommandText="Select * From TempAspects";
Temp->Active=true;

for(Temp->First();!Temp->Eof;Temp->Next())
{
 int Num=Temp->FieldByName("Номер аспекта")->Value;

MP<TDataSource>DS(Form1);
DS->DataSet=Temp;
DS->Enabled=true;

MP<TDBMemo>TDBM(Form1);
TDBM->Parent=Form1;
TDBM->Visible=false;
TDBM->Width=Width1;
TDBM->DataSource=DS;
TDBM->DataField="Выполняющиеся мероприятия";

TStrings* TS=TDBM->Lines;
int N=TS->Count;

MP<TADODataSet>Memo1(Form1);
Memo1->Connection=Database;
Memo1->CommandText="Select * From Memo1";
Memo1->Active=true;

for(int i=0;i<N;i++)
{
String S=TS->Strings[i];
Memo1->Insert();
Memo1->FieldByName("Num")->Value=Num;
Memo1->FieldByName("NumStr")->Value=i;
Memo1->FieldByName("Text")->Value=S;
Memo1->Post();
}

MP<TDBMemo>TDBM2(Form1);
TDBM2->Parent=Form1;
TDBM2->Visible=false;
TDBM2->Width=Width2;
TDBM2->DataSource=DS;
TDBM2->DataField="Предлагаемые мероприятия";

TStrings* TS2=TDBM2->Lines;
int N2=TS2->Count;

MP<TADODataSet>Memo2(Form1);
Memo2->Connection=Database;
Memo2->CommandText="Select * From Memo2";
Memo2->Active=true;

for(int i=0;i<N2;i++)
{
String S=TS2->Strings[i];
Memo2->Insert();
Memo2->FieldByName("Num")->Value=Num;
Memo2->FieldByName("NumStr")->Value=i;
Memo2->FieldByName("Text")->Value=S;
Memo2->Post();
}

MP<TDBMemo>TDBM3(Form1);
TDBM3->Parent=Form1;
TDBM3->Visible=false;
TDBM3->Width=Width3;
TDBM3->DataSource=DS;
TDBM3->DataField="Мониторинг и контроль";

TStrings* TS3=TDBM3->Lines;
int N3=TS3->Count;

MP<TADODataSet>Memo3(Form1);
Memo3->Connection=Database;
Memo3->CommandText="Select * From Memo3";
Memo3->Active=true;

for(int i=0;i<N3;i++)
{
String S=TS3->Strings[i];
Memo3->Insert();
Memo3->FieldByName("Num")->Value=Num;
Memo3->FieldByName("NumStr")->Value=i;
Memo3->FieldByName("Text")->Value=S;
Memo3->Post();
}

MP<TDBMemo>TDBM4(Form1);
TDBM4->Parent=Form1;
TDBM4->Visible=false;
TDBM4->Width=Width4;
TDBM4->DataSource=DS;
TDBM4->DataField="Предлагаемый мониторинг и контроль";

TStrings* TS4=TDBM4->Lines;
int N4=TS4->Count;

MP<TADODataSet>Memo4(Form1);
Memo4->Connection=Database;
Memo4->CommandText="Select * From Memo4";
Memo4->Active=true;

for(int i=0;i<N4;i++)
{
String S=TS4->Strings[i];
Memo4->Insert();
Memo4->FieldByName("Num")->Value=Num;
Memo4->FieldByName("NumStr")->Value=i;
Memo4->FieldByName("Text")->Value=S;
Memo4->Post();
}
}
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка подготовки к чтению аспектов", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase+" Подразделение: "+IntToStr(Podr));

Form1->Block=-1;
}
}
//*************************************************
void Client::PrepareMergeAspects(String NameDatabase, long NumLogin, long W1, long W2, long W3, long W4)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Пользователи", "Подготовка к объединению аспектов","IDC="+IntToStr(IDC)+" NumLogin="+IntToStr(NumLogin));
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Пользователи", "Подготовка к объединению аспектов", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase+" Номер логина: "+IntToStr(NumLogin));

try
{
TADOConnection *Database=GetDBConnection(NameDatabase);
MP<TADODataSet>Temp(Form1);
Temp->Connection=Database;
Temp->CommandText="Select * From TempAspects order by [Номер аспекта]";
Temp->Active=true;

MP<TADODataSet>Memo(Form1);
Memo->Connection=Database;


for(Temp->First();!Temp->Eof;Temp->Next())
{
 int TempNum=Temp->FieldByName("Номер аспекта")->Value;

MP<TMemo>M(Form1);
M->Visible=false;
M->Parent=Form1;
M->Width=1009;

TStrings* TT=M->Lines;
TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo1 Where Num="+IntToStr(TempNum)+" Order by Num, NumStr";
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("Выполняющиеся мероприятия")->Assign(TT);
Temp->Post();


TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo2 Where Num="+IntToStr(TempNum)+" Order by Num, NumStr";
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("Предлагаемые мероприятия")->Assign(TT);
Temp->Post();


TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo3 Where Num="+IntToStr(TempNum)+" Order by Num, NumStr";
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("Мониторинг и контроль")->Assign(TT);
Temp->Post();

TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo4 Where Num="+IntToStr(TempNum)+" Order by Num, NumStr";
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("Предлагаемый мониторинг и контроль")->Assign(TT);
Temp->Post();
}

MergeAspects(NameDatabase, NumLogin);
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка подготовки к объединению аспектов", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase+" NumLogin: "+IntToStr(NumLogin));

Form1->Block=-1;
}
}
//***************************************************

void Client::PrepareMergeAspectsH(String NameDatabase, long NumLogin, long W1, long W2, long W3, long W4)
{
//DiaryEvent->WriteEvent(Now(), this->NameComp, this->Login, "Пользователи", "Подготовка к объединению аспектов","IDC="+IntToStr(IDC)+" NumLogin="+IntToStr(NumLogin));
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Пользователи", "Подготовка к объединению рисков", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase+" Номер логина: "+IntToStr(NumLogin));

try
{
TADOConnection *Database=GetDBConnection(NameDatabase);
MP<TADODataSet>Temp(Form1);
Temp->Connection=Database;
Temp->CommandText="Select * From TempAspects order by [Номер аспекта]";
Temp->Active=true;

MP<TADODataSet>Memo(Form1);
Memo->Connection=Database;


for(Temp->First();!Temp->Eof;Temp->Next())
{
 int TempNum=Temp->FieldByName("Номер аспекта")->Value;

MP<TMemo>M(Form1);
M->Visible=false;
M->Parent=Form1;
M->Width=1009;

TStrings* TT=M->Lines;
TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo1 Where Num="+IntToStr(TempNum)+" Order by Num, NumStr";
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("Выполняющиеся мероприятия")->Assign(TT);
Temp->Post();


TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo2 Where Num="+IntToStr(TempNum)+" Order by Num, NumStr";
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("Предлагаемые мероприятия")->Assign(TT);
Temp->Post();


TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo3 Where Num="+IntToStr(TempNum)+" Order by Num, NumStr";
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("Мониторинг и контроль")->Assign(TT);
Temp->Post();

TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo4 Where Num="+IntToStr(TempNum)+" Order by Num, NumStr";
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("Предлагаемый мониторинг и контроль")->Assign(TT);
Temp->Post();
}

MergeAspectsH(NameDatabase, NumLogin);
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка подготовки к объединению опасностей", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase+" NumLogin: "+IntToStr(NumLogin));

Form1->Block=-1;
}
}
//***************************************************

void Client::MergeAspects(String NameDatabase, long NumLogin)
{
try
{
TADOConnection *Database=GetDBConnection(NameDatabase);
MP<TADODataSet>TempPodr(Form1);
TempPodr->Connection=Database;
TempPodr->CommandText="Select * From TempПодразделения";
TempPodr->Active=true;

MP<TADODataSet>TempAsp(Form1);
TempAsp->Connection=Database;
TempAsp->CommandText="select * From TempAspects";
TempAsp->Active=true;

MP<TADODataSet>Asp(Form1);
Asp->Connection=Database;
Asp->CommandText="SELECT Аспекты.*, ObslOtdel.Login FROM (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel WHERE (((ObslOtdel.Login)="+IntToStr(NumLogin)+"));";
Asp->Active=true;

for(TempAsp->First();!TempAsp->Eof;TempAsp->Next())
{
 int N=TempAsp->FieldByName("Подразделение")->Value;

 if(TempPodr->Locate("номер подразделения",N,SO))
 {
 int NumPodr=TempPodr->FieldByName("ServerNum")->Value;

 TempAsp->Edit();
 TempAsp->FieldByName("Подразделение")->Value=NumPodr;
 TempAsp->Post();
 }
 else
 {
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Сбой", "Сбой объединения аспектов (не найдено подразделение)", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase+" NumLogin: "+IntToStr(NumLogin));

 }
}

/*
MP<TADOCommand>Comm(Form1);
Comm->Connection=Database;
Comm->CommandText="DELETE Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.Num)="+IntToStr(NumLogin)+"));";
Comm->Execute();

String ST="INSERT INTO Аспекты ( Подразделение, Ситуация, [Вид территории], Деятельность, Специальность, Аспект, Воздействие, G, O, R, S, T, L, N, Z, Значимость, [Проявление воздействия], [Тяжесть последствий], Приоритетность, [Выполняющиеся мероприятия], [Предлагаемые мероприятия], [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Исполнитель, [Дата создания], [Начало действия], [Конец действия] ) ";

ST=ST+" SELECT TempAspects.Подразделение, TempAspects.Ситуация, TempAspects.[Вид территории], TempAspects.Деятельность, TempAspects.Специальность, TempAspects.Аспект, TempAspects.Воздействие, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.Значимость, TempAspects.[Проявление воздействия], TempAspects.[Тяжесть последствий], TempAspects.Приоритетность, TempAspects.[Выполняющиеся мероприятия], TempAspects.[Предлагаемые мероприятия], TempAspects.[Мониторинг и контроль], TempAspects.[Предлагаемый мониторинг и контроль], TempAspects.Исполнитель, TempAspects.[Дата создания], TempAspects.[Начало действия], TempAspects.[Конец действия] ";
ST=ST+" FROM TempAspects;";
Comm->CommandText=ST;
Comm->Execute();
*/
MP<TADOCommand>Comm(Form1);
Comm->Connection=Database;
Comm->CommandText="UPDATE (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel SET Аспекты.Del = False WHERE (((ObslOtdel.Login)="+IntToStr(NumLogin)+"));";
Comm->Execute();

for(Asp->First();!Asp->Eof;Asp->Next())
{
 int N=Asp->FieldByName("Номер аспекта")->Value;

 if(TempAsp->Locate("ServerNum", N, SO))
 {
  Asp->Edit();
  Asp->FieldByName("Подразделение")->Value=TempAsp->FieldByName("Подразделение")->Value;
  Asp->FieldByName("Ситуация")->Value=TempAsp->FieldByName("Ситуация")->Value;
  Asp->FieldByName("Вид территории")->Value=TempAsp->FieldByName("Вид территории")->Value;
  Asp->FieldByName("Деятельность")->Value=TempAsp->FieldByName("Деятельность")->Value;
  Asp->FieldByName("Специальность")->Value=TempAsp->FieldByName("Специальность")->Value;
  Asp->FieldByName("Аспект")->Value=TempAsp->FieldByName("Аспект")->Value;
  Asp->FieldByName("Воздействие")->Value=TempAsp->FieldByName("Воздействие")->Value;
  Asp->FieldByName("G")->Value=TempAsp->FieldByName("G")->Value;
  Asp->FieldByName("O")->Value=TempAsp->FieldByName("O")->Value;
  Asp->FieldByName("R")->Value=TempAsp->FieldByName("R")->Value;
  Asp->FieldByName("S")->Value=TempAsp->FieldByName("S")->Value;
  Asp->FieldByName("T")->Value=TempAsp->FieldByName("T")->Value;
  Asp->FieldByName("L")->Value=TempAsp->FieldByName("L")->Value;
  Asp->FieldByName("N")->Value=TempAsp->FieldByName("N")->Value;
  Asp->FieldByName("Z")->Value=TempAsp->FieldByName("Z")->Value;
  Asp->FieldByName("Значимость")->Value=TempAsp->FieldByName("Значимость")->Value;
  Asp->FieldByName("Проявление воздействия")->Value=TempAsp->FieldByName("Проявление воздействия")->Value;
  Asp->FieldByName("Тяжесть последствий")->Value=TempAsp->FieldByName("Тяжесть последствий")->Value;
  Asp->FieldByName("Приоритетность")->Value=TempAsp->FieldByName("Приоритетность")->Value;
  Asp->FieldByName("Выполняющиеся мероприятия")->Value=TempAsp->FieldByName("Выполняющиеся мероприятия")->Value;
  Asp->FieldByName("Предлагаемые мероприятия")->Value=TempAsp->FieldByName("Предлагаемые мероприятия")->Value;
  Asp->FieldByName("Мониторинг и контроль")->Value=TempAsp->FieldByName("Мониторинг и контроль")->Value;
  Asp->FieldByName("Предлагаемый мониторинг и контроль")->Value=TempAsp->FieldByName("Предлагаемый мониторинг и контроль")->Value;
  //Asp->FieldByName("Исполнитель")->Value=TempAsp->FieldByName("Исполнитель")->Value;
  Asp->FieldByName("Дата создания")->Value=TempAsp->FieldByName("Дата создания")->Value;
  Asp->FieldByName("Начало действия")->Value=TempAsp->FieldByName("Начало действия")->Value;
  Asp->FieldByName("Конец действия")->Value=TempAsp->FieldByName("Конец действия")->Value;
  Asp->Post();

  TempAsp->Delete();
 }
 else
 {
  Asp->Edit();
  Asp->FieldByName("Del")->Value=true;
  Asp->Post();
 }
}

Comm->CommandText="DELETE Аспекты.*, Аспекты.Del, ObslOtdel.Login FROM (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel WHERE (((Аспекты.Del)=True) AND ((ObslOtdel.Login)="+IntToStr(NumLogin)+"));";
Comm->Execute();

for(TempAsp->First();!TempAsp->Eof;TempAsp->Next())
{
  Asp->Insert();
  Asp->FieldByName("Подразделение")->Value=TempAsp->FieldByName("Подразделение")->Value;
  Asp->FieldByName("Ситуация")->Value=TempAsp->FieldByName("Ситуация")->Value;
  Asp->FieldByName("Вид территории")->Value=TempAsp->FieldByName("Вид территории")->Value;
  Asp->FieldByName("Деятельность")->Value=TempAsp->FieldByName("Деятельность")->Value;
  Asp->FieldByName("Специальность")->Value=TempAsp->FieldByName("Специальность")->Value;
  Asp->FieldByName("Аспект")->Value=TempAsp->FieldByName("Аспект")->Value;
  Asp->FieldByName("Воздействие")->Value=TempAsp->FieldByName("Воздействие")->Value;
  Asp->FieldByName("G")->Value=TempAsp->FieldByName("G")->Value;
  Asp->FieldByName("O")->Value=TempAsp->FieldByName("O")->Value;
  Asp->FieldByName("R")->Value=TempAsp->FieldByName("R")->Value;
  Asp->FieldByName("S")->Value=TempAsp->FieldByName("S")->Value;
  Asp->FieldByName("T")->Value=TempAsp->FieldByName("T")->Value;
  Asp->FieldByName("L")->Value=TempAsp->FieldByName("L")->Value;
  Asp->FieldByName("N")->Value=TempAsp->FieldByName("N")->Value;
  Asp->FieldByName("Z")->Value=TempAsp->FieldByName("Z")->Value;
  Asp->FieldByName("Значимость")->Value=TempAsp->FieldByName("Значимость")->Value;
  Asp->FieldByName("Проявление воздействия")->Value=TempAsp->FieldByName("Проявление воздействия")->Value;
  Asp->FieldByName("Тяжесть последствий")->Value=TempAsp->FieldByName("Тяжесть последствий")->Value;
  Asp->FieldByName("Приоритетность")->Value=TempAsp->FieldByName("Приоритетность")->Value;
  Asp->FieldByName("Выполняющиеся мероприятия")->Value=TempAsp->FieldByName("Выполняющиеся мероприятия")->Value;
  Asp->FieldByName("Предлагаемые мероприятия")->Value=TempAsp->FieldByName("Предлагаемые мероприятия")->Value;
  Asp->FieldByName("Мониторинг и контроль")->Value=TempAsp->FieldByName("Мониторинг и контроль")->Value;
  Asp->FieldByName("Предлагаемый мониторинг и контроль")->Value=TempAsp->FieldByName("Предлагаемый мониторинг и контроль")->Value;
  //Asp->FieldByName("Исполнитель")->Value=TempAsp->FieldByName("Исполнитель")->Value;
  Asp->FieldByName("Дата создания")->Value=TempAsp->FieldByName("Дата создания")->Value;
  Asp->FieldByName("Начало действия")->Value=TempAsp->FieldByName("Начало действия")->Value;
  Asp->FieldByName("Конец действия")->Value=TempAsp->FieldByName("Конец действия")->Value;
  Asp->Post();

  Asp->Active=false;
  Asp->Active=true;
  Asp->Last();

  TempAsp->Edit();
  TempAsp->FieldByName("ServerNum")->Value=Asp->FieldByName("Номер аспекта")->Value;
  TempAsp->Post();
}

Comm->CommandText="Delete * From TempПодразделения";
Comm->Execute();
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка объединения аспектов", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase+" NumLogin: "+IntToStr(NumLogin));

Form1->Block=-1;
}
}
//****************************************************
void Client::MergeAspectsH(String NameDatabase, long NumLogin)
{
try
{
TADOConnection *Database=GetDBConnection(NameDatabase);
MP<TADODataSet>TempPodr(Form1);
TempPodr->Connection=Database;
TempPodr->CommandText="Select * From TempПодразделения";
TempPodr->Active=true;

MP<TADODataSet>TempAsp(Form1);
TempAsp->Connection=Database;
TempAsp->CommandText="select * From TempAspects";
TempAsp->Active=true;

MP<TADODataSet>Asp(Form1);
Asp->Connection=Database;
Asp->CommandText="SELECT Аспекты.*, ObslOtdel.Login FROM (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel WHERE (((ObslOtdel.Login)="+IntToStr(NumLogin)+"));";
Asp->Active=true;

for(TempAsp->First();!TempAsp->Eof;TempAsp->Next())
{
 int N=TempAsp->FieldByName("Подразделение")->Value;

 if(TempPodr->Locate("номер подразделения",N,SO))
 {
 int NumPodr=TempPodr->FieldByName("ServerNum")->Value;

 TempAsp->Edit();
 TempAsp->FieldByName("Подразделение")->Value=NumPodr;
 TempAsp->Post();
 }
 else
 {
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Сбой", "Сбой объединения аспектов (не найдено подразделение) Hazards", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase+" NumLogin: "+IntToStr(NumLogin));

 }
}

/*
MP<TADOCommand>Comm(Form1);
Comm->Connection=Database;
Comm->CommandText="DELETE Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.Num)="+IntToStr(NumLogin)+"));";
Comm->Execute();

String ST="INSERT INTO Аспекты ( Подразделение, Ситуация, [Вид территории], Деятельность, Специальность, Аспект, Воздействие, G, O, R, S, T, L, N, Z, Значимость, [Проявление воздействия], [Тяжесть последствий], Приоритетность, [Выполняющиеся мероприятия], [Предлагаемые мероприятия], [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Исполнитель, [Дата создания], [Начало действия], [Конец действия] ) ";

ST=ST+" SELECT TempAspects.Подразделение, TempAspects.Ситуация, TempAspects.[Вид территории], TempAspects.Деятельность, TempAspects.Специальность, TempAspects.Аспект, TempAspects.Воздействие, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.Значимость, TempAspects.[Проявление воздействия], TempAspects.[Тяжесть последствий], TempAspects.Приоритетность, TempAspects.[Выполняющиеся мероприятия], TempAspects.[Предлагаемые мероприятия], TempAspects.[Мониторинг и контроль], TempAspects.[Предлагаемый мониторинг и контроль], TempAspects.Исполнитель, TempAspects.[Дата создания], TempAspects.[Начало действия], TempAspects.[Конец действия] ";
ST=ST+" FROM TempAspects;";
Comm->CommandText=ST;
Comm->Execute();
*/
MP<TADOCommand>Comm(Form1);
Comm->Connection=Database;
Comm->CommandText="UPDATE (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel SET Аспекты.Del = False WHERE (((ObslOtdel.Login)="+IntToStr(NumLogin)+"));";
Comm->Execute();

for(Asp->First();!Asp->Eof;Asp->Next())
{
 int N=Asp->FieldByName("Номер аспекта")->Value;

 if(TempAsp->Locate("ServerNum", N, SO))
 {
  Asp->Edit();
  Asp->FieldByName("Подразделение")->Value=TempAsp->FieldByName("Подразделение")->Value;
  Asp->FieldByName("Ситуация")->Value=TempAsp->FieldByName("Ситуация")->Value;
  Asp->FieldByName("Вид территории")->Value=TempAsp->FieldByName("Вид территории")->Value;
  Asp->FieldByName("Деятельность")->Value=TempAsp->FieldByName("Деятельность")->Value;
  Asp->FieldByName("Специальность")->Value=TempAsp->FieldByName("Специальность")->Value;
  Asp->FieldByName("Аспект")->Value=TempAsp->FieldByName("Аспект")->Value;
  Asp->FieldByName("Воздействие")->Value=TempAsp->FieldByName("Воздействие")->Value;
  Asp->FieldByName("G")->Value=TempAsp->FieldByName("G")->Value;
  Asp->FieldByName("O")->Value=TempAsp->FieldByName("O")->Value;
  Asp->FieldByName("R")->Value=TempAsp->FieldByName("R")->Value;
  Asp->FieldByName("S")->Value=TempAsp->FieldByName("S")->Value;
  Asp->FieldByName("T")->Value=TempAsp->FieldByName("T")->Value;
  Asp->FieldByName("L")->Value=TempAsp->FieldByName("L")->Value;
  Asp->FieldByName("N")->Value=TempAsp->FieldByName("N")->Value;
  Asp->FieldByName("Z")->Value=TempAsp->FieldByName("Z")->Value;
  Asp->FieldByName("Значимость")->Value=TempAsp->FieldByName("Значимость")->Value;
  Asp->FieldByName("Наименование значимости")->Value=TempAsp->FieldByName("Наименование значимости")->Value;
  Asp->FieldByName("Проявление воздействия")->Value=TempAsp->FieldByName("Проявление воздействия")->Value;
  Asp->FieldByName("Тяжесть последствий")->Value=TempAsp->FieldByName("Тяжесть последствий")->Value;
  Asp->FieldByName("Приоритетность")->Value=TempAsp->FieldByName("Приоритетность")->Value;
  Asp->FieldByName("Приор")->Value=TempAsp->FieldByName("Приор")->Value;
  Asp->FieldByName("Выполняющиеся мероприятия")->Value=TempAsp->FieldByName("Выполняющиеся мероприятия")->Value;
  Asp->FieldByName("Предлагаемые мероприятия")->Value=TempAsp->FieldByName("Предлагаемые мероприятия")->Value;
  Asp->FieldByName("Мониторинг и контроль")->Value=TempAsp->FieldByName("Мониторинг и контроль")->Value;
  Asp->FieldByName("Предлагаемый мониторинг и контроль")->Value=TempAsp->FieldByName("Предлагаемый мониторинг и контроль")->Value;
  //Asp->FieldByName("Исполнитель")->Value=TempAsp->FieldByName("Исполнитель")->Value;
  Asp->FieldByName("Дата создания")->Value=TempAsp->FieldByName("Дата создания")->Value;
  Asp->FieldByName("Начало действия")->Value=TempAsp->FieldByName("Начало действия")->Value;
  Asp->FieldByName("Конец действия")->Value=TempAsp->FieldByName("Конец действия")->Value;
  Asp->Post();

  TempAsp->Delete();
 }
 else
 {
  Asp->Edit();
  Asp->FieldByName("Del")->Value=true;
  Asp->Post();
 }
}

Comm->CommandText="DELETE Аспекты.*, Аспекты.Del, ObslOtdel.Login FROM (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel WHERE (((Аспекты.Del)=True) AND ((ObslOtdel.Login)="+IntToStr(NumLogin)+"));";
Comm->Execute();

for(TempAsp->First();!TempAsp->Eof;TempAsp->Next())
{
  Asp->Insert();
  Asp->FieldByName("Подразделение")->Value=TempAsp->FieldByName("Подразделение")->Value;
  Asp->FieldByName("Ситуация")->Value=TempAsp->FieldByName("Ситуация")->Value;
  Asp->FieldByName("Вид территории")->Value=TempAsp->FieldByName("Вид территории")->Value;
  Asp->FieldByName("Деятельность")->Value=TempAsp->FieldByName("Деятельность")->Value;
  Asp->FieldByName("Специальность")->Value=TempAsp->FieldByName("Специальность")->Value;
  Asp->FieldByName("Аспект")->Value=TempAsp->FieldByName("Аспект")->Value;
  Asp->FieldByName("Воздействие")->Value=TempAsp->FieldByName("Воздействие")->Value;
  Asp->FieldByName("G")->Value=TempAsp->FieldByName("G")->Value;
  Asp->FieldByName("O")->Value=TempAsp->FieldByName("O")->Value;
  Asp->FieldByName("R")->Value=TempAsp->FieldByName("R")->Value;
  Asp->FieldByName("S")->Value=TempAsp->FieldByName("S")->Value;
  Asp->FieldByName("T")->Value=TempAsp->FieldByName("T")->Value;
  Asp->FieldByName("L")->Value=TempAsp->FieldByName("L")->Value;
  Asp->FieldByName("N")->Value=TempAsp->FieldByName("N")->Value;
  Asp->FieldByName("Z")->Value=TempAsp->FieldByName("Z")->Value;
  Asp->FieldByName("Значимость")->Value=TempAsp->FieldByName("Значимость")->Value;
  Asp->FieldByName("Наименование значимости")->Value=TempAsp->FieldByName("Наименование значимости")->Value;
  Asp->FieldByName("Проявление воздействия")->Value=TempAsp->FieldByName("Проявление воздействия")->Value;
  Asp->FieldByName("Тяжесть последствий")->Value=TempAsp->FieldByName("Тяжесть последствий")->Value;
  Asp->FieldByName("Приоритетность")->Value=TempAsp->FieldByName("Приоритетность")->Value;
  Asp->FieldByName("Приор")->Value=TempAsp->FieldByName("Приор")->Value;
  Asp->FieldByName("Выполняющиеся мероприятия")->Value=TempAsp->FieldByName("Выполняющиеся мероприятия")->Value;
  Asp->FieldByName("Предлагаемые мероприятия")->Value=TempAsp->FieldByName("Предлагаемые мероприятия")->Value;
  Asp->FieldByName("Мониторинг и контроль")->Value=TempAsp->FieldByName("Мониторинг и контроль")->Value;
  Asp->FieldByName("Предлагаемый мониторинг и контроль")->Value=TempAsp->FieldByName("Предлагаемый мониторинг и контроль")->Value;
  //Asp->FieldByName("Исполнитель")->Value=TempAsp->FieldByName("Исполнитель")->Value;
  Asp->FieldByName("Дата создания")->Value=TempAsp->FieldByName("Дата создания")->Value;
  Asp->FieldByName("Начало действия")->Value=TempAsp->FieldByName("Начало действия")->Value;
  Asp->FieldByName("Конец действия")->Value=TempAsp->FieldByName("Конец действия")->Value;
  Asp->Post();

  Asp->Active=false;
  Asp->Active=true;
  Asp->Last();

  TempAsp->Edit();
  TempAsp->FieldByName("ServerNum")->Value=Asp->FieldByName("Номер аспекта")->Value;
  TempAsp->Post();
}

Comm->CommandText="Delete * From TempПодразделения";
Comm->Execute();
}
catch(...)
{
DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка объединения опасностей", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase+" NumLogin: "+IntToStr(NumLogin));

Form1->Block=-1;
}
}
//****************************************************

String Client::GetLogin()
{
return this->Login;
}
//*****************************************************
String Client::GetComputer()
{
 return this->pNameComp;
}
//****************************************************
//////////////////////////////////////////////////
mForm::mForm()
{

}
/////////////////////////////////////////////////
mForm::mForm(long IDF, String Name)
{
//mForm();
pIDF=IDF;
pName=Name;
LastIDT=1;
}
////////////////////////////////////////////////
mForm::~mForm()
{
//vector<mTable*>::const_iterator IT=VTable.begin();
for(unsigned int i=0;i<VTable.size();i++)
{
 delete VTable[i];
}
VTable.clear();
}
/////////////////////////////////////////////////
mTable* mForm::GetTable(long IDT)
{
for(unsigned int i=0;i<VTable.size();i++)
{
 if(VTable[i]->IDT()==IDT)
 {
  return VTable[i];
 }
}
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Служебная ошибка", "Ошибка получения таблицы", "IDТ="+IntToStr(IDT));

return NULL;
}
/////////////////////////////////////////////////
long mForm::IDF()
{
return pIDF;
}
////////////////////////////////////////////////
String mForm::Name()
{
return pName;
}
////////////////////////////////////////////////
long mForm::CreateTable(TADOConnection* Conn)
{
long IDT=LastIDT;

mTable* T=new mTable(LastIDT, Conn->Owner);
T->SetConnect(Conn);
T->setIDT(IDT);
LastIDT++;
VTable.push_back(T);
return IDT;
}
////////////////////////////////////////////////
bool mForm::DeleteTable(long IDT)
{
vector<mTable*>::iterator IT=VTable.begin();
for(unsigned int i=0;i<VTable.size();i++)
{
 if(VTable[i]->IDT()==IDT)
 {
  delete VTable[i];
  VTable.erase(IT);
  return true;
 }
 IT++;
}
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "Служебная ошибка", "Ошибка удаления таблицы из формы (не найдена таблица)", "IDТ="+IntToStr(IDT));

return false;
}
////////////////////////////////////////////
void mForm::SetCommandText(String Text, long IDT)
{
GetTable(IDT)->SetCommandText(Text);
}
////////////////////////////////////////////
String mForm::GetCommandText(long IDT)
{
return GetTable(IDT)->GetCommandText();
}
/////////////////////////////////////////////
void mForm::Active(bool Value, long IDT)
{
 GetTable(IDT)->Active(Value);
}
////////////////////////////////////////////
void mForm::Moving(long Comand, long IDT)
{
 GetTable(IDT)->Moving(Comand);
}
////////////////////////////////////////////
bool mForm::BOF_EOF(long Comand, long IDT)
{
return GetTable(IDT)->BOF_EOF(Comand);
}
////////////////////////////////////////////
Variant mForm::FieldByName(String FieldName, long IDT)
{
return GetTable(IDT)->FieldByName(FieldName);
}
/////////////////////////////////////////////
void mForm::FieldByName(Variant Value, String FieldName, long IDT)
{
 GetTable(IDT)->FieldByName(Value, FieldName);
}
/////////////////////////////////////////////
void mForm::EditingRecord(long Comand, long IDT)
{
GetTable(IDT)->EditingRecord(Comand);
}
////////////////////////////////////////////
long mForm::RecordCount(long IDT)
{
return GetTable(IDT)->RecordCount();
}
////////////////////////////////////////////
long mForm::RecNo(long IDT)
{
return GetTable(IDT)->RecNo();
}
//////////////////////////////////////////
void mForm::Move(long Step, long IDT)
{
GetTable(IDT)->Move(Step);
}
////////////////////////////////////////////
bool mForm::Locate(long IDT, String FieldName, String Key, TLocateOptions SearchOptions)
{
return GetTable(IDT)->Locate(FieldName, Key, SearchOptions);
}
/////////////////////////////////////////////
void mForm::Delete(long IDT)
{
GetTable(IDT)->Delete();
}
////////////////////////////////////////////
Variant mForm::Fields(long NumField, long IDT)
{
return GetTable(IDT)->Fields(NumField);
}
//////////////////////////////////////////////
void mForm::Fields(Variant Value, long NumField, long IDT)
{
GetTable(IDT)->Fields(Value, NumField);
}
//////////////////////////////////////////////
long mForm::FieldsCount(long IDT)
{
return GetTable(IDT)->FieldsCount();
}
/////////////////////////////////////////////
//++++++++++++++++++++++++++++++++++++++
mTable::mTable(long IDT, TComponent *Owner)
{
pOwner=Owner;
Table=new TADODataSet(Form1);
}
//+++++++++++++++++++++++++++++++++++++++
mTable::~mTable()
{
delete Table;
//delete pOwner;
}
//+++++++++++++++++++++++++++++++++++++++
long mTable::IDT()
{
return pIDT;
}
//++++++++++++++++++++++++++++++++++++++++
void mTable::setIDT(long IDT)
{
pIDT=IDT;
}
//+++++++++++++++++++++++++++++++++++++++++
void mTable::SetConnect(TADOConnection* Conn)
{
Table->Connection=Conn;
}
//+++++++++++++++++++++++++++++++++++++++++
void mTable::SetCommandText(String Text)
{
Table->CommandText=Text;
}
//+++++++++++++++++++++++++++++++++++++++++
String mTable::GetCommandText()
{
return Table->CommandText;
}
//++++++++++++++++++++++++++++++++++++++++++
void mTable::Active(bool Value)
{
Table->Active=Value;
}
//++++++++++++++++++++++++++++++++++++++++++
void mTable::Moving(long Comand)
{
switch (Comand)
{
 case 0:
 {
 Table->First();
 break;
 }
 case 1:
 {
 Table->Last();
 break;
 }
 case 2:
 {
 Table->Prior();
 break;
 }
 case 3:
 {
 Table->Next();
 break;
 }

}

}
//++++++++++++++++++++++++++++++++++++++++++++
bool mTable::BOF_EOF(long Comand)
{
switch (Comand)
{
 case 0:
 {
 return Table->Bof;
 }
 case 1:
 {
 return Table->Eof;
 }
}
return false;
}
//++++++++++++++++++++++++++++++++++++++++++++++
Variant mTable::FieldByName(String FieldName)
{
return Table->FieldByName(FieldName)->Value;
}
//+++++++++++++++++++++++++++++++++++++++++++++++
void mTable::FieldByName(Variant Value, String FieldName)
{
Table->FieldByName(FieldName)->Value=Value;
}
//+++++++++++++++++++++++++++++++++++++++++++++++
void mTable::EditingRecord(long Comand)
{
switch (Comand)
{
 case 0:
 {
 Table->Insert();
 break;
 }
 case 1:
 {
 Table->Edit();
 break;
 }
 case 2:
 {
 Table->Post();
 break;
 }
}
}
//++++++++++++++++++++++++++++++++++++++++++++++
long mTable::RecordCount()
{
return Table->RecordCount;
}
//++++++++++++++++++++++++++++++++++++++++++++++
long mTable::RecNo()
{
return Table->RecNo;
}
//++++++++++++++++++++++++++++++++++++++++++++++
void mTable::Move(long Step)
{
Table->MoveBy(Step);
}
//++++++++++++++++++++++++++++++++++++++++++++
bool mTable::Locate(String FieldName, String Key, TLocateOptions SearchOptions)
{
return Table->Locate(FieldName, Key, SearchOptions);
}
//+++++++++++++++++++++++++++++++++++++++++++++
void mTable::Delete()
{
Table->Delete();
}
//+++++++++++++++++++++++++++++++++++++++++++++++
Variant mTable::Fields(long NumField)
{
return Table->FieldList->Fields[NumField]->Value;
}
//+++++++++++++++++++++++++++++++++++++++++++++++
void mTable::Fields(Variant Value, long NumField)
{
Table->FieldList->Fields[NumField]->Value=Value;
}
//+++++++++++++++++++++++++++++++++++++++++++++++++
long  mTable::FieldsCount()
{
return Table->FieldList->Count;
}
//++++++++++++++++++++++++++++++++++++++++++++++++++
//.................................................
Diary::Diary(TComponent* Owner, String PatchDiary)
{
this->Form=(TForm*)Owner;
DiaryBasePath=PatchDiary;

DiaryBase=new TADOConnection(Owner);

DiaryBase->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source="+PatchDiary+";Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False";
DiaryBase->LoginPrompt=false;
DiaryBase->Connected=true;
}
//................................................
Diary::~Diary()
{
delete DiaryBase;
}
//.................................................
void Diary::WriteEvent(TDateTime DT, String Comp, String Login, String Type, String Name, String Prim)
{
MP<TADODataSet>TypeOp(Form);
TypeOp->Connection=DiaryBase;
TypeOp->CommandText="Select * From TypeOp order by Num";
TypeOp->Active=true;

int NumType;
if(TypeOp->Locate("NameType",Type, SO))
{
//Найдено, прочитать
NumType=TypeOp->FieldByName("Num")->Value;
}
else
{
//Не найдено, записать
TypeOp->Insert();
TypeOp->FieldByName("NameType")->Value=Type;
TypeOp->Post();
TypeOp->Active=false;
TypeOp->Active=true;
TypeOp->Last();
NumType=TypeOp->FieldByName("Num")->Value;
}

MP<TADODataSet>Operation(Form);
Operation->Connection=DiaryBase;
Operation->CommandText="Select * From Operations order by Num";
Operation->Active=true;

/*
Variant locvalues[] = {EDep->Text, EFam->Text};
Table1->Locate("Dep;Fam", VarArrayOf(locvalues,1),      SearchOptions<<loPartialKey<<loCaseInsensitive);
*/
int NumOp;
Variant locvalues[] = {Name, NumType};
if(Operation->Locate("NameOperation;Type",VarArrayOf(locvalues,1),SO))
{
//найдено, читать
NumOp=Operation->FieldByName("Num")->Value;
}
else
{
//Ненайдено, записать
Operation->Insert();
Operation->FieldByName("NameOperation")->Value=Name;
Operation->FieldByName("Type")->Value=NumType;
Operation->Post();
Operation->Active=false;
Operation->Active=true;
Operation->Last();
NumOp=Operation->FieldByName("Num")->Value;
}

MP<TADODataSet>Event(Form);
Event->Connection=DiaryBase;
Event->CommandText="Select * From Events order by Num";
Event->Active=true;

Event->Insert();
Event->FieldByName("Date_Time")->Value=DT;
Event->FieldByName("Comp")->Value=Comp;
Event->FieldByName("Login")->Value=Login;
Event->FieldByName("Operation")->Value=NumOp;
Event->FieldByName("Prim")->Value=Prim;
Event->Post();
}
//................................................
void Diary::WriteEvent(TDateTime DT, String Comp, String Login, String Type, String Name)
{
WriteEvent(DT, Comp, Login, Type, Name, "");
}
//.................................................

