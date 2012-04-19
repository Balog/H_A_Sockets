//---------------------------------------------------------------------------
#include <vcl.h>
#include "utilcls.h"
#include <vector>
#include <ADODB.hpp>
#include <DB.hpp>
#include <ComCtrls.hpp>

#ifndef ClientClassH
#define ClientClassH
//---------------------------------------------------------------------------
using namespace std;
struct DBItem
{
 int Num;
 String Name;
 String IDName;
};

class Form;
class Table;
//****************************************
class Client
{
friend Form;
private:
String ServerName;
String CompName;
long IDClient;
//long IDDB;
vector<DBItem*>VIDDB;
Variant Server;
String CName();
TTimer* VerTimer;
void __fastcall VTimer(TObject *Sender);
String Login;

GUID ClassID;

vector<Form*>VForms;
vector<Form*>::iterator IVF;

public:
Client();
~Client();

void Stop();
void Start();
void SetLogin(String Login);
bool S_Error;
void Connect(String Server_Name, String Login);
int Disconnect();
long IDC();
Variant GetServer();

Form* RegForm(TForm* Form);
int IDF(TForm* F);
Table* CreateTable(TForm* F, String ServerName, int NumIDDB);
Table* CreateTable(TForm* F, String ServerName, String DBName);
void DeleteTable(TForm* F, Table* T);

bool AddDatabase(String Name, String IDName);
bool GetDatabaseConnect(long IDC, int NumIDDB);
void SetDatabaseConnect(long IDC, int NumIDDB, bool Connect);
void SetCommandText(unsigned int NumIDDB, String Text);
void SetCommandText(String NameDB, String Text);
void CommandExec(unsigned int NumIDDB);
void CommandExec(String NameDB);
bool MergeLogins(int NumIDDB);
bool NetError();

void LoadTable(Table* FromCopy, TADODataSet* ToCopy);
void LoadTable(TADODataSet* FromCopy, Table* ToCopy);
void LoadTable(Table* FromCopy, TADODataSet *ToCopy, TLabel* L, TProgressBar* PB);
void LoadTable(TADODataSet* FromCopy, Table* ToCopy, TLabel* L, TProgressBar* PB);

//0 - все совпадает
//1 - не совпадает содержимое полей
//2 - не сопадает число записей
//3 - не совпадает число полей
int VerifyTable(TADODataSet* LTable, Table* RTable);
int VerifyTable(Table* LTable, TADODataSet* RTable);

bool PrepareLoadMetod(String NameDatabase);

void MergeNodeBranch(String NameDatabase1, String NameNode, String NameBranch, String NameDatabase2, String NameTable, String NameField, String Key, String Name);
void MergeNodeBranch(String NameDatabase1, String NameNode, String NameBranch);
void MergeZn(String Database1, String Database2);
void SaveMetod(String Database);
void SavePodr(String DB1);
void SaveSit(String DB1,String DB2);
void MergeAspectsMainSpec();
void PrepareLoadAspects(int Podr, int Memo1Width, int Memo2Width, int Memo3Width, int Memo4Width);

bool SetBlock(int Value);
bool GetBlock();
void PrepareMergeAspects(int NumLogin, int W1, int W2, int W3, int W4);
void WriteDiaryEvent(String Type, String Name, String Prim);
int GetLicenseCount(String DBName);
};
//****************************************
class Form
{
friend Table;
private:
int IDClient;
int IDForm;
Variant Server;
GUID ClassID;
String CName();
String ServerName;
vector<Table*>VTable;


public:
Form();
Form(String CompName);
~Form();

String NameForm;
int RegForm(long IDC, String Name);
bool UnRegForm(long IDC, long IDF);
int IDF();

void Stop();
void Start(int IDC);

Table* CreateTable(String ServerName, int IDC, int IDDB);
void DeleteTable(Table* T);

};
//****************************************
class Table
{
private:
String CName();
String ServerName;
Variant Server;
GUID ClassID;

int IDClient;
int IDDataBase;
int IDForm;
int IDTable;

public:

Table();
Table(String ServerName);
~Table();

void Stop();
void Start();

int CreateTable(long IDC, long IDDB, long IDF);
bool DeleteTable();

void SetCommandText(String Text);
String GetCommandText();
bool Active(bool Value);



void First();
void Last();
void Prior();
void Next();

bool BOF_EOF(long comand);
bool bof();
bool eof();
Variant FieldByName(String FieldName);
void FieldByName (Variant Value, String FieldName);
void Insert();
void Edit();
void Post();

long RecordCount();
long RecNo();
void MoveBy(long Step);
bool Locate(String FileName, String Key, long SearchParam);
void Delete();

Variant Fields(long NumField);
void Fields(Variant Value, long NumField);
long FieldsCount();
};
//****************************************
 #endif
