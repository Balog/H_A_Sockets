//---------------------------------------------------------------------------
#include <vcl.h>
#include <vector>
#include <ADODB.hpp>
#include <DB.hpp>
#include "UServer.h"

using namespace std;
#ifndef ServerClassesH
#define ServerClassesH

//---------------------------------------------------------------------------

struct DatabaseClient
{
long IDDB;
TADOConnection *Database;
TADOCommand* Command;
String FileName;
};
//--------------------------------------------------------------------------
class Client;
class mForm;
class mTable;
class Diary;

class mClients
{
private:
vector<Client*>VClients;
vector<Client*>::iterator IVC;

long LastIDC;
TObject *Sender;
TComponent* pOwner;
String DatabasePath;
String DiaryBase;



TNotifyEvent FCountClientChanged;

void __fastcall ClientTimer(TObject *Sender);

public:
mClients();
mClients(TComponent* Owner, String DiaryBase);
~mClients();
long Connect(String CompName, String Login);
bool Disconnect(long IDC);
bool Disconnect(String Name);
bool DisconnectAll();
long ClientCount();
String GetName(int Number);
int GetIDC(int Number);
Diary* DiaryEvent;
String NameComp;
String Login;

Client* GetClient(long IDC);
Client* GetClient(int Number);
void SetDatabasePatc(String Path);
long AddDatabase(long IDC, String DatabaseName);
bool GetDatabaseConnect(long IDC, long IDDB);
bool SetDatabaseConnect(long IDC, long IDDB, bool Connect);
TADOCommand* Command(long IDC, long IDDB);
bool MergeLogins(long IDC, long IDDB);
void VerifyConnect(long IDC);

long CreateTable(long IDC, long IDDB, long IDF);
bool DeleteTable(long IDC, long IDF, long IDT);

void SetCommandText(String Text, long IDC, long IDF, long IDT);
String GetCommandText(long IDC, long IDF, long IDT);
void Active(bool Value, long IDC, long IDF, long IDT);
void Moving(long Comand, long IDC, long IDF, long IDT);
bool BOF_EOF(long Comand, long IDC, long IDF, long IDT);
Variant FieldByName(String FieldName, long IDC, long IDF, long IDT);
void FieldByName(Variant Value, String FieldName, long IDC, long IDF, long IDT);
void EditingRecord(long Comand, long IDC, long IDF, long IDT);

long RecordCount(long IDC, long IDF, long IDT);
long RecNo(long IDC, long IDF, long IDT);
void Move(long Step, long IDC, long IDF, long IDT);
bool Locate(long IDC, long IDF, long IDT, String FieldName, String Key, long SearchParam);
void Delete(long IDC, long IDF, long IDT);
Variant Fields(long NumField, long IDC, long IDF, long IDT);
void Fields(Variant Value, long NumField, long IDC, long IDF, long IDT);
long FieldsCount(long IDC, long IDF, long IDT);

bool PrepareReadMetod(long IDC, String NameDatabase);
bool MergeNodeBranch(long IDC, String NameDatabase1, String NameNode, String NameBranch, String NameDatabase2, String NameTable, String NameField, String NameKey, String Name);
bool MergeNodeBranch(long IDC, String NameDatabase1, String NameNode, String NameBranch);
void MergeZn(long IDC, String Database1, String Database2);
void SaveMetod(long IDC, String Database);
void SavePodr(long IDC, String DB1);
void SaveSit(long IDC, String DB1, String DB2);
void MergeAspectsMainSpec(long IDC);
void MergeAspectsMainSpecH(long IDC);
void PrepareLoadAspects( long IDC, String NameDatabase, long Podr, long Width1, long Width2, long Width3, long Width4);
void PrepareLoadAspectsH( long IDC, String NameDatabase, long Podr, long Width1, long Width2, long Width3, long Width4);
void PrepareMergeAspects(String NameDatabase, long IDC, long NumLogin, long W1, long W2, long W3, long W4);
void PrepareMergeAspectsH(String NameDatabase, long IDC, long NumLogin, long W1, long W2, long W3, long W4);
//__published:
__property TNotifyEvent CountClientChanged = {read= FCountClientChanged, write= FCountClientChanged};
//virtual CountClientChanged (mClients *Sender);
void WriteDiaryEvent(String Comp, String Login, String Type, String Name, String Prim);


};
//**************************************************
class Client
{
private:
TLocateOptions SO;
String pNameComp;
long pIDC;
vector<mForm*>VForm;
vector<mForm*>::iterator IFC;
long LastIDF;
long LastIDDB;
vector<DatabaseClient*>VDatabase;
TControl* pOwner;
String Login;
String DiaryBase;
Diary* DiaryEvent;
public:
Client();
Client(long IDC, String NameComp, String Login, TControl *Owner, String DiaryBase);
Client(TComponent *S);
~Client();
TTimer* ClTimer;

TObject *Sender;
long FormCount();
Client* SearchClient(long IDC);

mForm* GetForm(long IDF);
mForm* GetForm(int Number);

long IDC();
String NameComp();

long RegForm(String Name);
bool UnRegForm(long IDF);
bool UnRegForm(String Name);
bool UnRegAll();

bool AddDB(String DBPath, String Name, long *IDDB, TComponent* Owner);
bool GetDatabaseConnect(long IDDB);
bool SetDatabaseConnect(long IDDB, bool Connect);
TADOCommand* Command(long IDDB);
bool MergeLogins(long IDDB);

long CreateTable(long IDDB, long IDF);
bool DeleteTable(long IDF, long IDT);

void SetCommandText(String Text,long IDF, long IDT);
String GetCommandText(long IDF, long IDT);
void Active(bool Value, long IDF, long IDT);
void Moving(long Comand, long IDF, long IDT);
bool BOF_EOF (long Comand, long IDF, long IDT);
Variant FieldByName(String FieldName, long IDF, long IDT);
void FieldByName(Variant Value, String FieldName, long IDF, long IDT);
void EditingRecord(long Comand, long IDF, long IDT);

long RecordCount(long IDF, long IDT);
long RecNo(long IDF, long IDT);
void Move(long Step, long IDF, long IDT);
bool Locate(long IDF, long IDT, String FieldName, String Key, TLocateOptions SearchOptions);
void Delete(long IDF, long IDT);
Variant Fields(long NumField, long IDF, long IDT);
void Fields(Variant Value, long NumField, long IDF, long IDT);
long FieldsCount(long IDF, long IDT);
bool PrepareReadMetod(String NameDatabase);

bool MergeNodeBranch(String NameDatabase1, String NameNode, String NameBranch, String NameDatabase2, String NameTable, String NameField, String NameKey, String Name);
bool MergeNodeBranch(String NameDatabase1, String NameNode, String NameBranch);
void MergeZn(String Database1, String Database2);
void SaveMetod(String Database);
void SavePodr(String DB1);
void SaveSit(String DB1, String DB2);
TADOConnection* GetDBConnection(String NameDatabase);
void MergeAspectsMainSpec();
void MergeAspectsMainSpecH();
void PrepareLoadAspects(String NameDatabase, long Podr, long Width1, long Width2, long Width3, long Width4);
void PrepareLoadAspectsH(String NameDatabase, long Podr, long Width1, long Width2, long Width3, long Width4);
void PrepareMergeAspects(String NameDatabase, long NumLogin, long W1, long W2, long W3, long W4);
void PrepareMergeAspectsH(String NameDatabase, long NumLogin, long W1, long W2, long W3, long W4);
void MergeAspects(String NameDatabase, long NumLogin);
void MergeAspectsH(String NameDatabase, long NumLogin);
String GetLogin();
String GetComputer();

};
///////////////////////////////////////////////////////
class mForm
{
private:
long pIDF;
long LastIDT;
String pName;
vector<mTable*>VTable;

public:
mForm();
mForm(long IDF, String Name);
~mForm();
long IDF();
String Name();

mTable* GetTable(long IDT);
int CountForms(long IDC);
long CreateTable(TADOConnection* Conn);
bool DeleteTable(long IDT);

void SetCommandText(String Text, long IDT);
String GetCommandText(long IDT);
void Active(bool Value, long IDT);
void Moving(long Comand, long IDT);
bool BOF_EOF(long Comand, long IDT);
Variant FieldByName(String FieldName, long IDT);
void FieldByName(Variant Value, String FieldName, long IDT);
void EditingRecord(long Comand, long IDT);

long RecordCount(long IDT);
long RecNo(long IDT);
void Move(long Step, long IDT);
bool Locate(long IDT, String FieldName, String Key, TLocateOptions SearchOptions);
void Delete(long IDT);
Variant Fields(long NumField, long IDT);
void Fields(Variant Value, long NumField, long IDT);
long FieldsCount(long IDT);
};
//////////////////////////////////////////////////////

//+++++++++++++++++++++++++++++++++++++++++++++++++++++
class mTable
{
private:
long pIDT;
TADODataSet* Table;
TComponent* pOwner;
public:
mTable(long IDT, TComponent *Owner);
~mTable();
long IDT();
void setIDT(long IDT);
void SetConnect(TADOConnection* Conn);

void SetCommandText(String Text);
String GetCommandText();
void Active(bool Value);
void Moving(long Comand);
bool BOF_EOF(long Comand);
Variant FieldByName(String FieldName);
void FieldByName(Variant Value, String FieldName);
void EditingRecord(long Comand);

long RecordCount();
long RecNo();
void Move(long Step);
bool Locate(String FieldName, String Key, TLocateOptions SearchOptions);
void Delete();
Variant Fields(long NumField);
void Fields(Variant Value, long NumField);
long FieldsCount();
};
//...........................................................
class Diary
{
private:
String DiaryBasePath;
TForm* Form;
TADOConnection *DiaryBase;
TLocateOptions SO;
public:
Diary(TComponent* Owner, String PatchDiary);
~Diary();
void WriteEvent(TDateTime DT, String Comp, String Login, String Type, String Name);
void WriteEvent(TDateTime DT, String Comp, String Login, String Type, String Name, String Prim);

};
#endif
