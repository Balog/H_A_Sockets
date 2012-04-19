//---------------------------------------------------------------------------


#pragma hdrstop

#include "ServerClasses.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

mClients::mClients()
{
LastIDC=1;
}

mClients::mClients(TComponent* Owner)
{
//mClients();
LastIDC=1;

pOwner=Owner;
}
//--------------------------------------------
mClients::~mClients()
{
for(unsigned int i=0;i<VClients.size();i++)
{
delete VClients[i];
}
VClients.clear();


}
//--------------------------------------------
long mClients::Connect(String NameComp)
{

long IDC=LastIDC;
LastIDC++;

Client *C=new Client(IDC, NameComp);
VClients.push_back(C);


if (FCountClientChanged) CountClientChanged (Sender) ;

return IDC;
}
//--------------------------------------------
bool mClients::Disconnect(long IDC)
{
bool Ret=false;
IVC=VClients.begin();
for(unsigned int i=0;i<VClients.size();i++)
{
 if(VClients[i]->IDC()==IDC)
 {
 Ret=true;
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
  delete VClients[i];
  VClients.erase(IVC);
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
GetClient(IDC)->AddDB(DatabasePath, DatabaseName, &IDDB, pOwner);
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

  return GetClient(IDC)->CreateTable(IDDB, IDF);

}
//----------------------------------------------
bool mClients::DeleteTable(long IDC, long IDF, long IDT)
{

  return GetClient(IDC)->DeleteTable(IDF, IDT);

}
//-------------------------------------------------
void mClients::SetCommandText(String Text, long IDC, long IDF, long IDT)
{
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
GetClient(IDC)->Delete(IDF, IDT);
}
//-------------------------------------------
//********************************************
Client::Client()
{
LastIDF=1;
LastIDDB=1;
}
//********************************************
Client::Client(TObject *S)
{
//Client();
LastIDF=1;
LastIDDB=1;
Sender=S;
}
///******************************************
Client::~Client()
{
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
}
//*********************************************
Client::Client(long idc, String NC)
{
LastIDF=1;
LastIDDB=1;
pNameComp=NC;
pIDC=idc;

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
 }
 IFC++;
}
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
DatabaseClient* DBC=new DatabaseClient();

DBC->Database=new TADOConnection (Owner);
DBC->Command=new TADOCommand(Owner);
String P=DBPath+"\\"+Name+".mdb";
P="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+P+";Persist Security Info=False";

DBC->Database->ConnectionString=P;
DBC->Database->LoginPrompt=false;
DBC->Database->Connected=true;

DBC->Command->Connection=DBC->Database;
*IDDB=LastIDDB;

DBC->IDDB=LastIDDB;
LastIDDB++;
VDatabase.push_back(DBC);


return true;
}
//*************************************************
bool Client::GetDatabaseConnect(long IDDB)
{
for(unsigned int i=0;i<VDatabase.size();i++)
{
 if(VDatabase[i]->IDDB==IDDB)
 {
  return VDatabase[i]->Database->Connected;
 }
}
return false;
}
//**********************************************
bool Client::SetDatabaseConnect(long IDDB, bool Connect)
{
for(unsigned int i=0;i<VDatabase.size();i++)
{
 if(VDatabase[i]->IDDB==IDDB)
 {
  VDatabase[i]->Database->Connected=Connect;
  return true;
 }
}
return false;
}
//**********************************************
long Client::CreateTable(long IDDB, long IDF)
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

  return GetForm(IDF)->CreateTable( Conn);

}
else
{
 return NULL;
}

}
//**********************************************
bool Client::DeleteTable(long IDF, long IDT)
{

  return GetForm(IDF)->DeleteTable(IDT);

}
//**********************************************

void Client::SetCommandText(String Text,long IDF, long IDT)
{
GetForm(IDF)->SetCommandText(Text, IDT);
}
//**********************************************
String Client::GetCommandText(long IDF, long IDT)
{
return GetForm(IDF)->GetCommandText(IDT);
}
//**********************************************
void Client::Active(bool Value, long IDF, long IDT)
{
GetForm(IDF)->Active(Value, IDT);
}
//***********************************************
void Client::Moving(long Comand, long IDF, long IDT)
{
GetForm(IDF)->Moving(Comand, IDT);
}
//************************************************
bool Client::BOF_EOF (long Comand, long IDF, long IDT)
{
return GetForm(IDF)->BOF_EOF(Comand, IDT);
}
//************************************************
Variant Client::FieldByName(String FieldName, long IDF, long IDT)
{
return GetForm(IDF)->FieldByName(FieldName, IDT);
}
//*************************************************
void Client::FieldByName(Variant Value, String FieldName, long IDF, long IDT)
{
GetForm(IDF)->FieldByName(Value, FieldName, IDT);
}
//*************************************************

void Client::EditingRecord(long Comand, long IDF, long IDT)
{
GetForm(IDF)->EditingRecord(Comand, IDT);
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
//++++++++++++++++++++++++++++++++++++++
mTable::mTable(long IDT, TComponent *Owner)
{
pOwner=Owner;
Table=new TADODataSet(pOwner);
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
