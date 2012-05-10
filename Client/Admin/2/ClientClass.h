//---------------------------------------------------------------------------
#include <vcl.h>
#include "utilcls.h"
#include <vector>
#include <ADODB.hpp>
#include <DB.hpp>
#include <ScktComp.hpp>
#include <ActnList.hpp>
#include <ActnMan.hpp>
#include "MDBConnector.h"
using namespace std;
#ifndef ClientClassH
#define ClientClassH
//---------------------------------------------------------------------------
struct DBItem
{
 int Num;
 String Name;
};
//****************************************
struct CDBItem
{
int Num;
String Name;
String ServerDB;
int NumDatabase;
int NumLicense;
};
//****************************************
struct Trig
{
 int Var;
 int Max;
 String TrueAction;
 String FalseAction;
};
//****************************************
class Form;
class Table;
//****************************************
struct WaitAction
{
int WaitCommand;
int NextCommand;
vector<String>ParamComm;
};
//******************************************
class Client
{
public:
Client(TClientSocket *ClientSocket, TActionManager *ActMan, TForm *Form);
~Client();

void Connect(String ServerName, int Port);
void CommandExec(int Comm, vector<String>);
Form* RegForm(String FormName);

WaitAction Act;
vector<Form*>VForms;
vector<Form*>::iterator IVF;

void StartAction(String NameAction);
void ConnectDatabase(String Name,int Number,  bool Connect);
void ReadTable(String NameDB, String ServerSQL, String ClientSQL);
void DecodeTable(String NameDB, String ClientSQL, String Text);

vector<CDBItem>VDB;
vector<Trig>VTrigger;
void ActTrigger(int NumTrigger);

int GetIDDBName(String Name);
void RegisterDatabase(String Name, int NumLic);

MDBConnector* Database;
MDBConnector* Diary;
void LoginResult(String Login, String Pass, bool Ok);
int GetLicenseCount(String DBName);
void WriteDiaryEvent(String Type, String Name, String Prim);
void WriteDiaryEvent(String Type, String Name);
String Login;
void WriteTable(String Database, String ClientSQLText, String ServerSQLText);
private:
String GetIP();
TClientSocket *Socket;

TActionManager *ActionManager;
TForm *Owner;
void VerifyLicense(String NameDB);
String IP;
String TableToStr(String SQLText);

};
//****************************************
class Form
{
public:
Form();
~Form();
String FormName;
int IDF;

void RegForm(String FormName);
private:

};
//****************************************
class Table
{
public:

private:

};
//****************************************
#endif
