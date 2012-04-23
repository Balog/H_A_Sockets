//---------------------------------------------------------------------------
#include <vcl.h>
#include "utilcls.h"
#include <vector>
#include <ADODB.hpp>
#include <DB.hpp>
#include <ScktComp.hpp>
#include <ActnList.hpp>
#include <ActnMan.hpp>
using namespace std;
#ifndef ClientClassH
#define ClientClassH
//---------------------------------------------------------------------------
struct DBItem
{
 int Num;
 String Name;
};

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
Client(TClientSocket *ClientSocket, TActionManager *ActMan);
~Client();

void Connect(String ServerName, int Port);
void CommandExec(int Comm, vector<String>);
Form* RegForm(String FormName);

WaitAction Act;
vector<Form*>VForms;
vector<Form*>::iterator IVF;

void StartAction(String NameAction);
void ConnectDatabase(String Name, bool Connect);
void ReadTable(String NameDB, String ServerSQL, String ClientSQL); 
private:
String GetIP();
TClientSocket *Socket;

TActionManager *ActionManager;




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