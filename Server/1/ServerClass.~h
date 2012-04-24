//---------------------------------------------------------------------------
#include <vcl.h>
#include <vector>
#include <ADODB.hpp>
#include <DB.hpp>
#include <ScktComp.hpp>
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
TADOCommand* Command;
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

class Clients
{
 private:

String DatabasePath;
String DiaryBase;



 public:
Clients(TComponent* Owner, String DiaryBase);
~Clients();
vector<Client*>VClients;
vector<Client*>::iterator IVC;
vector<BaseItem>VBases;
TComponent* pOwner;

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
 public:
Client(Clients*);
~Client();
TCustomWinSocket *Socket;
String IP;
String AppPatch;
int LastCommand;

void CommandExec(int Comm, vector<String>);
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


 public:


};
#endif
