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
private:
String GetIP();
TClientSocket *Socket;
WaitAction Act;
TActionManager *ActionManager;

vector<Form*>VForms;
vector<Form*>::iterator IVF;

void StartAction(String NameAction);
};
//****************************************
class Form
{
public:
Form();
~Form();
String FormName;

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
