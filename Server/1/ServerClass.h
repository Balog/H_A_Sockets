//---------------------------------------------------------------------------
#include <vcl.h>
#include <vector>
#include <ADODB.hpp>
#include <DB.hpp>
#include <ScktComp.hpp>
//#include "MDBConnector.h"
using namespace std;
#ifndef ServerClassH
#define ServerClassH
//---------------------------------------------------------------------------
struct BaseItem
{
String Name;
String Patch;
bool MainSpec;
String FileName;
int LicCount;
TADOConnection *Database;
//TADOCommand* Command;
};
//------------------------------------------------------------------------
/*
struct DatabaseClient
{
long IDDB;
TADOConnection *Database;
TADOCommand* Command;
String FileName;
};
*/
//-----------------------------------------------------------------------
class Client;
class mForm;
class mTable;
class Diary;
//-----------------------------------------------------------------------
class Clients
{
 private:

String DatabasePath;
String DiaryBase;




 public:
Clients(TComponent* Owner);
~Clients();
Diary* DiaryEvent;
vector<Client*>VClients;
vector<Client*>::iterator IVC;
vector<BaseItem>VBases;
TComponent* pOwner;
void WriteDiaryEvent(String Comp, String Login, String Type, String Name, String Prim);
void WriteDiaryEvent(String Comp, String Login, String Type, String Name);
void ConnectDiary(String DiaryPatch);
void Clients::SearchDubl(String IP, String Login, String AppPatch, TCustomWinSocket *S);
};
//-------------------------------------------------------------------------
class Client
{
 private:
vector<mForm*>VForm;
vector<mForm*>::iterator IFC;
Clients* Parent;
TADOConnection* GetDatabase(String NameDB);
String TableToStr(String NameDB, String SQLText);
void DecodeTable(String NameDB, String ServerSQL, String Text);
void MergeLogins(String NameDB);

 public:
Client(Clients*);
~Client();
TCustomWinSocket *Socket;
String IP;
String AppPatch;
int LastCommand;
bool SDubl;

void CommandExec(int Comm, vector<String>);

String Login;

};
//-------------------------------------------------------------------------
class mForm
{
 public:
 mForm();
 ~mForm();

 int IDF;
 String NameForm;
 private:

};
//-------------------------------------------------------------------------
class Diary
{
private:
String DiaryBasePath;
TForm* Form;
TADOConnection *DiaryBase;
TLocateOptions SO;
TComponent* Owner;
public:
Diary(TComponent* Owner, String Patch);
~Diary();
void WriteEvent(TDateTime DT, String Comp, String Login, String Type, String Name);
void WriteEvent(TDateTime DT, String Comp, String Login, String Type, String Name, String Prim);

};
#endif
