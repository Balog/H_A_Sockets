//---------------------------------------------------------------------------
#include <vector>
#include <ADODB.hpp>
#include <DB.hpp>
#include <ScktComp.hpp>
using namespace std;
#ifndef ServerClassH
#define ServerClassH
//---------------------------------------------------------------------------
struct DatabaseClient
{
long IDDB;
TADOConnection *Database;
TADOCommand* Command;
String FileName;
};
//-----------------------------------------------------------------------
class Client;
class mForm;
class mTable;
class Diary;

class Clients
{
 private:
TComponent* pOwner;
String DatabasePath;
String DiaryBase;



 public:
Clients(TComponent* Owner, String DiaryBase);
~Clients();
vector<Client*>VClients;
vector<Client*>::iterator IVC;
};
//-------------------------------------------------------------------------
class Client
{
 private:


 public:
Client();
~Client();
TCustomWinSocket *Socket;
String IP;


};
//-------------------------------------------------------------------------
class mForm
{
 public:

 private:

};
//-------------------------------------------------------------------------
class Diary
{
 private:


 public:


};
#endif
