//---------------------------------------------------------------------------
#include <vcl.h>
#include <vector>
#include <ADODB.hpp>
#include <DB.hpp>

using namespace std;
#ifndef ServerClassesH
#define ServerClassesH
//---------------------------------------------------------------------------
struct DatabaseClient
{
long IDDB;
TADOConnection *Database;
TADOCommand* Command;
};
//--------------------------------------------------------------------------
class Client;
class mForm;
class mTable;

class mClients
{
private:
vector<Client*>VClients;
vector<Client*>::iterator IVC;

long LastIDC;
TObject *Sender;
TComponent* pOwner;
String DatabasePath;

TNotifyEvent FCountClientChanged;

public:
mClients();
mClients(TComponent* Owner);
~mClients();
long Connect(String CompName);
bool Disconnect(long IDC);
bool Disconnect(String Name);
bool DisconnectAll();
long ClientCount();
String GetName(int Number);
int GetIDC(int Number);

Client* GetClient(long IDC);
Client* GetClient(int Number);
void SetDatabasePatc(String Path);
long AddDatabase(long IDC, String DatabaseName);
bool GetDatabaseConnect(long IDC, long IDDB);
bool SetDatabaseConnect(long IDC, long IDDB, bool Connect);

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
//__published:
__property TNotifyEvent CountClientChanged = {read= FCountClientChanged, write= FCountClientChanged};
//virtual CountClientChanged (mClients *Sender);
};
//**************************************************
class Client
{
private:
String pNameComp;
long pIDC;
vector<mForm*>VForm;
vector<mForm*>::iterator IFC;
long LastIDF;
long LastIDDB;
vector<DatabaseClient*>VDatabase;

public:
Client();
Client(long IDC, String NameComp);
Client(TObject *S);
~Client();

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
};

#endif
