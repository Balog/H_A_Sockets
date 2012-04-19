//---------------------------------------------------------------------------


#pragma hdrstop

#include "ClientClass.h"
#include "AdmMain.h"
//---------------------------------------------------------------------------

#pragma package(smart_init)
String VarToStr(Variant V)
{
String S="";
if(!V.IsNull())
{
S=V;
}
return S;
}
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
long VarToInt(Variant V)
{
long S=0;
if(!V.IsNull())
{
S=V;
}
return S;
}
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
bool VarToBool(Variant V)
{
bool S=false;
if(!V.IsNull())
{
S=V;
}
return S;
}
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Client::Client()
{
Login="��� �� ���������";
VerTimer=new TTimer(Main);
VerTimer->Interval=3000;
VerTimer->OnTimer=VTimer;
VerTimer->Enabled=false;
VForms.clear();

S_Error=true;
}
//-----------------------------------------
void __fastcall Client::VTimer(TObject *Sender)
{
if(Server.intVal!=0)
{
Server.OleProcedure("VerifyConnect", IDClient);
}
}
//-----------------------------------------
void Client::Connect(String Server_Name, String Login)
{
/*
        char buffer[32];
	DWORD size;
	size=sizeof(buffer);
	GetComputerName(buffer,&size);
        CompName=buffer;
*/
try
{
CompName=CName();
if(Server_Name=="")
{


Server_Name=CompName;
}

ServerName=Server_Name;

ClassID = Comobj::StringToGUID("Server.Clients");
Server  = (IUnknown*)CreateRemoteComObject(ServerName, ClassID);


long IDC=0;
int OK=false;



//� �������� ������� ������� �� ����������� � ���������� � ������ ��������
//��������� �������� IDC

Server.OleProcedure("Connect", CompName.c_str(), Login.c_str(), &IDC, &OK);
IDClient=IDC;
}
catch(...)
{
S_Error=false;
Server=0;
}

}
//------------------------------------------
int Client::Disconnect()
{
int OK=true;
//������� �� ���������� � �������� �� ������ ��������
if(Server.intVal!=0)
{
Server.OleProcedure("Disconnect_IDC", IDC(), &OK);
}
//�������� ���������� ����������
VerTimer->Enabled=false;

return OK; //���� true
}
//-------------------------------------------

Client::~Client()
{
Disconnect();
for(unsigned int i=0;i<VForms.size();i++)
{
delete VForms[i];
}
VerTimer->Enabled=false;
delete VerTimer;


for(unsigned int i=0; i<VIDDB.size();i++)
{
delete VIDDB[i];
}






//Server=Unassigned;
//Server=0;
}
//-------------------------------------------
long Client::IDC()
{
 return IDClient;
}
//--------------------------------------------
String Client::CName()
{
String CompName;
        char buffer[32];
	DWORD size;
	size=sizeof(buffer);
	GetComputerName(buffer,&size);
        CompName=buffer;

        return CompName;
}
//--------------------------------------------
Variant Client::GetServer()
{
return Server;
}
//--------------------------------------------
bool Client::AddDatabase(String Name)
{
long pIDDB=0;
if(Server.intVal!=0)
{
Server.OleProcedure("AddDatabase", IDClient, Name.c_str(), &pIDDB);
}
//IDDB=pIDDB;
DBItem* DBI=new DBItem();
DBI->Num=pIDDB;
DBI->Name=Name;
VIDDB.push_back(DBI);
return pIDDB!=0;
}
//--------------------------------------------
bool Client::GetDatabaseConnect(long IDC, int NumIDDB)
{
int Connect=false;
if(Server.intVal!=0)
{
Connect=Server.OlePropertyGet("DatabaseConnect", IDClient, VIDDB[NumIDDB]->Num);
}
return Connect;
}
//--------------------------------------------
void Client::SetDatabaseConnect(long IDC, int NumIDDB, bool Connect)
{
if(Server.intVal!=0)
{
Server.OlePropertySet("DatabaseConnect", IDC, VIDDB[NumIDDB]->Num, Connect);
}
}
//-------------------------------------------
void Client::SetCommandText(unsigned int NumIDDB, String Text)
{
if(Server.intVal!=0)
{
Server.OlePropertySet("SetCommandText", this->IDC(), VIDDB[NumIDDB]->Num, WideString(Text.c_str()).c_bstr());
}
}
//--------------------------------------------
void Client::SetCommandText(String NameDB, String Text)
{
if(Server.intVal!=0)
{
int NumIDDB;
for(unsigned int i=0;i<VIDDB.size();i++)
{
DBItem* DBI=VIDDB[i];
if(DBI->Name==NameDB)
{
NumIDDB=i;
break;
}
}

Server.OlePropertySet("SetCommandText", this->IDC(), VIDDB[NumIDDB]->Num, WideString(Text.c_str()).c_bstr());

}
}
//--------------------------------------------
void Client::CommandExec(unsigned int NumIDDB)
{
int Ret=false;
if(Server.intVal!=0)
{
int i=0;
for(;i<200;i++)
{
Server.OleProcedure("CommandExec", this->IDC(), VIDDB[NumIDDB]->Num, &Ret);
if(Ret)
{
break;
}
else
{
Sleep(500);
}
}
if(i>=200)
{
ShowMessage("������� �� ���������.\r������ ������������\r����� ���������� ������� �������.");
}
}
}
//-------------------------------------------
void Client::CommandExec(String NameDB)
{

int Ret=false;
if(Server.intVal!=0)
{
int i=0;
for(;i<200;i++)
{
int NumIDDB;
for(unsigned int i=0;i<VIDDB.size();i++)
{
DBItem* DBI=VIDDB[i];
if(DBI->Name==NameDB)
{
NumIDDB=i;
break;
}
}
Server.OleProcedure("CommandExec", this->IDC(), VIDDB[NumIDDB]->Num, &Ret);

if(Ret)
{
break;
}
else
{
Sleep(500);
}
}
if(i>=200)
{
ShowMessage("������� �� ���������.\r������ ������������\r����� ���������� ������� �������.");
}
}

}
//-------------------------------------------
bool Client::MergeLogins(String NameDB)
{
long OK=false;
long Wait=false;
if(Server.intVal!=0)
{
int i=0;
for(;i<200;i++)
{
int NumIDDB;

for(unsigned int i=0;i<VIDDB.size();i++)
{
DBItem* DBI=VIDDB[i];
if(DBI->Name==NameDB)
{
NumIDDB=i;
break;
}
}

Server.OleProcedure("MergeLogins", this->IDC(),  VIDDB[NumIDDB]->Num, &OK, &Wait);

if(!Wait)
{
break;
}
else
{
Sleep(500);
}
}
if(i>=200)
{
ShowMessage("������� �� ���������.\r������ ������������\r����� ���������� ������� �������.");
}
}

return OK;
}
//--------------------------------------------
Form* Client::RegForm(TForm* F)
{
Form* Frm=new Form(ServerName);
Frm->RegForm(IDClient,F->Name);
Frm->NameForm=F->Name;
VForms.push_back(Frm);

return Frm;
}
//--------------------------------------------
int Client::IDF(TForm* F)
{
int Ret=0;
for(unsigned int i=0;i<VForms.size();i++)
{
if(VForms[i]->NameForm==F->Name)
{
 Ret=VForms[i]->IDF();
 break;
}

}
return Ret;
}
//---------------------------------------------
void Client::Stop()
{
/*
Sleep(1000);
for(unsigned int i=0;i<VForms.size();i++)
{
VForms[i]->Stop();
}

Disconnect();
Server=Unassigned;
Server=0;
*/

Sleep(1000);


for(unsigned int i=0;i<VForms.size();i++)
{
VForms[i]->Stop();
}
Disconnect();
Server=Unassigned;
Server=0;
}
//--------------------------------------------
void Client::Start()
{

if(Server.intVal==0)
{
Connect(ServerName, this->Login);
//������������ ����������� � ����� ������ �� ������� ��� ������
for(unsigned int i=0;i<VIDDB.size();i++)
{
//int N=
long pIDDB=0;
if(Server.intVal!=0)
{
Server.OleProcedure("AddDatabase", IDClient, VIDDB[i]->Name.c_str(), &pIDDB);
}
VIDDB[i]->Num=pIDDB;
SetDatabaseConnect(IDClient, i, true);
}
//Server  = (IUnknown*)CreateRemoteComObject(ServerName, ClassID);
for(unsigned int i=0;i<VForms.size();i++)
{
VForms[i]->Start(IDClient);
}
VerTimer->Enabled=S_Error;
}





}
//--------------------------------------------
void Client::SetLogin(String Login)
{
this->Login=Login;
}
//--------------------------------------------
Table* Client::CreateTable(TForm* F, String ServerName, int NumIDDB)
{
Table* Ret;
DBItem* DBI=VIDDB[NumIDDB];
int IDDB=DBI->Num;
for(unsigned int i=0;i<VForms.size();i++)
{
if(VForms[i]->NameForm==F->Name)
{
Ret=VForms[i]->CreateTable(ServerName, IDC(), IDDB) ;
break;
}

}
return Ret;
}
//------------------------------------------
Table* Client::CreateTable(TForm* F, String ServerName, String DBName)
{
Table* Ret;
int IDDB;
for(unsigned int i=0;i<VIDDB.size();i++)
{
DBItem* DBI=VIDDB[i];
if(DBI->Name==DBName)
{
IDDB=DBI->Num;
break;
}
}

for(unsigned int i=0;i<VForms.size();i++)
{
if(VForms[i]->NameForm==F->Name)
{
Ret=VForms[i]->CreateTable(ServerName, IDC(), IDDB) ;
break;
}

}

return Ret;
}
//-------------------------------------------
void Client::DeleteTable(TForm* F, Table* T)
{
IVF=VForms.begin();
for(unsigned int i=0;i<VForms.size();i++)
{
if(VForms[i]->NameForm==F->Name)
{
VForms[i]->DeleteTable(T);
break;
}
IVF++;
}
}
//-------------------------------------------
bool Client::NetError()
{
bool B=false;
if(Server.intVal!=0)
{
B=Server.OlePropertyGet("NetError");
}
return B;
}
//-----------------------------------------------
void Client::LoadTable(Table* FromCopy, TADODataSet *ToCopy)
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
ShowMessage("����� ����� ���������� ������ �� ���������");
}
}
//---------------------------------------------------------------
void Client::LoadTable(TADODataSet* FromCopy, Table* ToCopy)
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
ShowMessage("����� ����� ���������� ������ �� ���������");
}
}
//---------------------------------------------------------------
int Client::VerifyTable(TADODataSet* FromCopy, Table* ToCopy)
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
int Client::VerifyTable(Table* FromCopy, TADODataSet* ToCopy)
{
return VerifyTable(ToCopy, FromCopy);
}
//---------------------------------------------------------------
bool Client::PrepareLoadMetod(String NameDatabase)
{
int Ret=false;
if(Server.intVal!=0)
{
Server.OleProcedure("PrepareReadMetod", this->IDC(), NameDatabase.c_str(), &Ret);
}
return Ret;
}
//---------------------------------------------------------------
void Client::MergeNodeBranch(String NameDatabase1, String NameNode, String NameBranch, String NameDatabase2, String NameTable, String NameField)
{
if(Server.intVal!=0)
{
Server.OleProcedure("MergeNodeBranch", this->IDC(), NameDatabase1.c_str(), NameNode.c_str(), NameBranch.c_str(), NameDatabase2.c_str(), NameTable.c_str(), NameField.c_str());
}
}
//---------------------------------------------------------------
void Client::WriteDiaryEvent(String Type, String Name, String Prim)
{
if(Server.intVal==0)
{
this->Start();
Server.OleProcedure("WriteDiary", CompName.c_str(), Login.c_str(), Type.c_str(), Name.c_str(), Prim.c_str());
this->Stop();
}
else
{
Server.OleProcedure("WriteDiary", CompName.c_str(), Login.c_str(), Type.c_str(), Name.c_str(), Prim.c_str());
}
}
//-------------------------------------------------------------------
int Client::GetLicenseCount(String DBName)
{
int Res=-1;
if(Server.intVal!=0)
{
Res=Server.OlePropertyGet("LicenseCount", DBName.c_str());
}
return Res;
}
//------------------------------------------------------------------
//********************************************
Form::Form()
{

}
//*********************************************
Form::Form(String Server_Name)
{
try
{
String CompName=CName();
if(Server_Name=="")
{


Server_Name=CompName;
}

ServerName=Server_Name;


ClassID = Comobj::StringToGUID("Server.Form");
Server  = (IUnknown*)CreateRemoteComObject(ServerName, ClassID);
}
catch(...)
{
Server=0;
}
}
//**********************************************
String Form::CName()
{
String CompName;
        char buffer[32];
	DWORD size;
	size=sizeof(buffer);
	GetComputerName(buffer,&size);
        CompName=buffer;

        return CompName;
}
//***********************************************
Form::~Form()
{
int OK=false;
if(Server.intVal!=0)
{
Server.OleProcedure("UnRegForm_IDF",IDClient, IDForm, &OK);
}

for(unsigned int i=0; i<VTable.size();i++)
{
 delete VTable[i];
}
Server=Unassigned;
Server=0;
}
//*************************************************
int Form::RegForm(long IDC, String NameForm)
{
int IDF=0;
int OK=false;
if(Server.intVal!=0)
{
Server.OleProcedure("RegForm", IDC, NameForm.c_str(), &IDF, &OK);
}
IDForm=IDF;
IDClient=IDC;
return IDForm;
}
//*************************************************
int Form::IDF()
{
return this->IDForm;
}

//**************************************************
Table* Form::CreateTable(String ServerName, int IDC, int IDDB)
{
Table* Tab;
Tab=new Table(ServerName);
Tab->CreateTable(IDC, IDDB, IDForm);

VTable.push_back(Tab);

return Tab;
}
//*************************************************
void Form::DeleteTable(Table* T)
{
vector<Table*>::iterator IVT=VTable.begin();
for(unsigned int i=0; VTable.size();i++)
{
 if(VTable[i]==T)
 {
  delete VTable[i];
  VTable.erase(IVT);
  break;
 }
 IVT++;
}
}
//**************************************************
void Form::Stop()
{
/*
int OK=false;
for(unsigned int i=0; VTable.size();i++)
{
VTable[i]->Stop();
}
if(Server.intVal!=0)
{
Server.OleProcedure("UnRegForm_IDF",IDClient, IDForm, &OK);
}
Server=Unassigned;
Server=0;
*/
int OK=false;
for(unsigned int i=0; VTable.size();i++)
{
VTable[i]->Stop();
}
if(Server.intVal!=0)
{
Server.OleProcedure("UnRegForm_IDF",IDClient, IDForm, &OK);
}
Server=Unassigned;
Server=0;
}
//**************************************************
void Form::Start(int IDC)
{
Server  = (IUnknown*)CreateRemoteComObject(ServerName, ClassID);
RegForm(IDC, NameForm);
for(unsigned int i=0; VTable.size();i++)
{
VTable[i]->Start();
}
}
//**************************************************
////////////////////////////////////////////////////
Table::Table()
{

}
////////////////////////////////////////////////////
Table::Table(String Server_Name)
{
try
{
String CompName=CName();
if(Server_Name=="")
{


Server_Name=CompName;
}

ServerName=Server_Name;


ClassID = Comobj::StringToGUID("Server.Tables");
Server  = (IUnknown*)CreateRemoteComObject(ServerName, ClassID);
}
catch(...)
{
Server=0;
}
}
////////////////////////////////////////////////////
Table::~Table()
{
DeleteTable();
Server=Unassigned;
Server=0;
}
/////////////////////////////////////////////////////
String Table::CName()
{
String CompName;
        char buffer[32];
	DWORD size;
	size=sizeof(buffer);
	GetComputerName(buffer,&size);
        CompName=buffer;

        return CompName;
}
////////////////////////////////////////////////////
int Table::CreateTable(long IDC, long IDDB, long IDF)
{
IDClient=IDC;
IDForm=IDF;
IDDataBase=IDDB;
long IDT=0;
if(Server.intVal!=0)
{
Server.OleProcedure("CreateTable", IDC, IDDB, IDF, &IDT);
}
IDTable=IDT;
return IDT;
}
///////////////////////////////////////////////////
bool Table::DeleteTable()
{
int OK=false;
if(Server.intVal!=0)
{
Server.OleProcedure("DeleteTable", IDClient, IDForm, IDTable, &OK);
}
return OK;
}
///////////////////////////////////////////////////
void Table::Stop()
{
/*
Server=Unassigned;
Server=0;
*/
try
{
Server=Unassigned;
Server=0;
}
catch(...)
{
}

}
///////////////////////////////////////////////////
void Table::Start()
{
Server  = (IUnknown*)CreateRemoteComObject(ServerName, ClassID);
}
///////////////////////////////////////////////////
void Table::SetCommandText(String Text)
{
if(Server.intVal!=0)
{
Server.OlePropertySet("CommandText", IDClient, IDForm, IDTable, Text.c_str());
}
}
//////////////////////////////////////////////////
String Table::GetCommandText()
{
String Text="";
if(Server.intVal!=0)
{
Text=Server.OlePropertyGet("CommandText",IDClient, IDForm, IDTable);
}
return Text;
}
//////////////////////////////////////////////////
bool Table::Active(bool Value)
{
int Ret=false;
if(Server.intVal!=0)
{


int i=0;
for(;i<200;i++)
{
Server.OleProcedure("Active", Value, IDClient, IDForm, IDTable, &Ret);
if(Ret)
{
break;
}
else
{
Sleep(500);
}
}
if(i>=200)
{
ShowMessage("������� �� ���������.\r������ ������������\r����� ���������� ������� �������.");
}





}
return Ret==1;
}
///////////////////////////////////////////////////
void Table::First()
{
if(Server.intVal!=0)
{
Server.OleProcedure("Moving", 0, IDClient, IDForm, IDTable);
}
}
///////////////////////////////////////////////////
void Table::Last()
{
if(Server.intVal!=0)
{
Server.OleProcedure("Moving", 1, IDClient, IDForm, IDTable);
}
}
///////////////////////////////////////////////////
void Table::Prior()
{
if(Server.intVal!=0)
{
Server.OleProcedure("Moving", 2, IDClient, IDForm, IDTable);
}
}
///////////////////////////////////////////////////
void Table::Next()
{
if(Server.intVal!=0)
{
Server.OleProcedure("Moving", 3, IDClient, IDForm, IDTable);
}
}
///////////////////////////////////////////////////
bool Table::bof()
{
if(Server.intVal!=0)
{
return Server.OlePropertyGet("BOF_EOF", 0, IDClient, IDForm, IDTable);
}
else
{
 return false;
}
}
///////////////////////////////////////////////////
bool Table::eof()
{
if(Server.intVal!=0)
{
return Server.OlePropertyGet("BOF_EOF", 1, IDClient, IDForm, IDTable);
}
else
{
return false;
}
}
///////////////////////////////////////////////////
Variant Table::FieldByName(String FieldName)
{
if(Server.intVal!=0)
{
return Server.OlePropertyGet("FieldByName_Variant", WideString(FieldName.c_str()).c_bstr(), IDClient, IDForm, IDTable);
}
else
{
return Null;
}
}
////////////////////////////////////////////////////
void Table::FieldByName (Variant Value, String FieldName)
{

Variant N=Value;
if(Server.intVal!=0)
{
Server.OlePropertySet("FieldByName_Variant", WideString(FieldName.c_str()).c_bstr(),IDClient, IDForm, IDTable, WideString(((String)N).c_str()).c_bstr());//WideString(N.c_str()).c_bstr()
}
}
/////////////////////////////////////////////////////
void Table::Insert()
{
if(Server.intVal!=0)
{
Server.OleProcedure("EditingRecord", 0, IDClient, IDForm, IDTable);
}
}
///////////////////////////////////////////////////
void Table::Edit()
{
if(Server.intVal!=0)
{
Server.OleProcedure("EditingRecord", 1, IDClient, IDForm, IDTable);
}
}
///////////////////////////////////////////////////
void Table::Post()
{
if(Server.intVal!=0)
{
Server.OleProcedure("EditingRecord", 2, IDClient, IDForm, IDTable);
}
}
///////////////////////////////////////////////////
long Table::RecordCount()
{
if(Server.intVal!=0)
{
return Server.OlePropertyGet("RecordCount", IDClient, IDForm, IDTable);
}
else
{
return 0;
}
}
///////////////////////////////////////////////////
long Table::RecNo()
{
if(Server.intVal!=0)
{
return Server.OlePropertyGet("RecNo", IDClient, IDForm, IDTable);
}
else
{
return 0;
}
}
///////////////////////////////////////////////////
void Table::MoveBy(long Step)
{
if(Server.intVal!=0)
{
Server.OleProcedure("Move", Step, IDClient, IDForm, IDTable);
}
}
//////////////////////////////////////////////////////
bool Table::Locate(String FieldName, String Key, long SearchParam)
{
if(Server.intVal!=0)
{
return Server.OlePropertyGet("Locate", IDClient, IDForm, IDTable, WideString(FieldName.c_str()).c_bstr(), WideString(Key.c_str()).c_bstr(), SearchParam);
}
else
{
return false;
}
}
//////////////////////////////////////////////////////
void Table::Delete()
{
if(Server.intVal!=0)
{
Server.OleProcedure("Delete", IDClient, IDForm, IDTable);
}
}
///////////////////////////////////////////////////////
Variant Table::Fields(long NumField)
{
if(Server.intVal!=0)
{
return Server.OlePropertyGet("Fields", NumField, IDClient, IDForm, IDTable);
}
else
{
return Null;
}
}
////////////////////////////////////////////////////////
void Table::Fields(Variant Value, long NumField)
{
if(Server.intVal!=0)
{
Server.OlePropertySet("Fields", NumField, IDClient, IDForm, IDTable,  WideString(((String)Value).c_str()).c_bstr());
}
}
////////////////////////////////////////////////////////
long Table::FieldsCount()
{
if(Server.intVal!=0)
{
return Server.OlePropertyGet("FieldsCount", IDClient, IDForm, IDTable);
}
else
{
return 0;
}
}
/////////////////////////////////////////////////////////

