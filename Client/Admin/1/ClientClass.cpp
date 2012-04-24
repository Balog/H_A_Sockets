//---------------------------------------------------------------------------


#pragma hdrstop

#include "ClientClass.h"
#include "MasterPointer.h"
//---------------------------------------------------------------------------

#pragma package(smart_init)
Client::Client(TClientSocket *ClientSocket, TActionManager *ActMan, TForm *Form)
{
Socket=ClientSocket;
ActionManager=ActMan;
Owner=Form;
}
//********************************************************************
Client::~Client()
{

}
//********************************************************************
void Client::Connect(String ServerName, int Port)
{
//Ожидается ответ на команду 1

Act.WaitCommand=1;
Socket->Address=ServerName;
Socket->Host=ServerName;
Socket->Port=Port;
Socket->Active=true;
}
//********************************************************************
void Client::CommandExec(int Comm, vector<String>Parameters)
{
if(Act.WaitCommand==Comm | Act.WaitCommand==0)
{
switch(Comm)
{
 case 1:
 {

  //Отвечаем на команду 1
  String Mess="Command:1;2|"+GetIP()+"|"+Application->ExeName;
  //Необходимо запомнить что ожидается команда номер 2
  Act.WaitCommand=2;
  Socket->Socket->SendText(Mess);
  break;
 }
 case 2:
 {
 //После сигнала что сервер готов к дальнейшей после идентификации клиента работе регистрируем формы
 //Запускаем действие - регистрация главной формы.
 //Далее ожидается еще регистрация формы
 Act.WaitCommand=Act.NextCommand;
 StartAction("RegForm_Form1");

 break;
 }
 case 3:
 {


 //Читаем записи о дальнейших действиях
 String Action=Act.ParamComm[0];

 Act.ParamComm.clear();

 Act.WaitCommand=0;
 String NumPrevForm=Parameters[0];
 Act.ParamComm.push_back(NumPrevForm);
 StartAction(Action);
 break;
 }
 case 4:
 {
 String Action=Act.ParamComm[0];
 StartAction(Action);
 break;
 }
 case 5:
 {
 //ShowMessage(Parameters[0]);
 DecodeTable(Act.ParamComm[0],Act.ParamComm[1],Parameters[0]);
 }
}
}
else
{
 ShowMessage("Ожидалась команда "+IntToStr(Act.WaitCommand)+" а пришел ответ на команду "+IntToStr(Comm));
}
}
//*********************************************************************
String Client::GetIP()
{
    String addr="";

    const int WSVer = 0x101;
    WSAData wsaData;
    hostent *h;
    char Buf[128];
    if (WSAStartup(WSVer, &wsaData) == 0)
    {
    if (gethostname(&Buf[0], 128) == 0)
    {
    h = gethostbyname(&Buf[0]);
    if (h != NULL)
    {
    //ShowMessage(inet_ntoa (*(reinterpret_cast<in_addr *>(*(h->h_addr_list)))));
    addr=inet_ntoa (*(reinterpret_cast<in_addr *>(*(h->h_addr_list))));

    }
    else MessageBox(0,"Вы не в сети. И IP адреса у вас нет.",0,0);
    }
    WSACleanup;
    }


    return addr;
}
//**************************************************************************
Form* Client::RegForm(String FormName)
{
Form* F=new Form();
F->RegForm(FormName);
VForms.push_back(F);
//Процедура регистрации формы на сервере, команда 3
  String Mess="Command:3;1|"+FormName;
  //Необходимо запомнить что ожидается команда номер 3
  Act.WaitCommand=3;
  Socket->Socket->SendText(Mess);

return F;
}
//****************************************************************************
void Client::StartAction(String NameAction)
{
for(int i=0;i<ActionManager->ActionCount;i++)
{
 if(ActionManager->Actions[i]->Name==NameAction)
 {
  ActionManager->Actions[i]->Execute();
 }
}
}
//***************************************************************************
void Client::ConnectDatabase(String Name, bool Connect)
{
String C;
if(Connect)
{
C="true";
}
else
{
C="false";
}

String Mess="Command:4;2|"+Name+"|"+C+"|";
Socket->Socket->SendText(Mess);

}
//*************************************************************************
void Client::ReadTable(String NameDB, String ServerSQL, String ClientSQL)
{
 Act.ParamComm.clear();
 Act.WaitCommand=5;
 Act.ParamComm.push_back(NameDB);
 Act.ParamComm.push_back(ClientSQL);

 Socket->Socket->SendText("Command:5;3|"+NameDB+"|"+ServerSQL+"|");
}
//*************************************************************************
void Client::DecodeTable(String NameDB, String ClientSQL, String Text)
{
 MP<TADODataSet>Tab(Owner);
 
}
//*************************************************************************
/////////////////////////////////////////////////////////////////////////////
Form::Form()
{
IDF=-1;
FormName="-";
}
//////////////////////////////////////////////////////////////////////////////
Form::~Form()
{

}
/////////////////////////////////////////////////////////////////////////////
void Form::RegForm(String FormName)
{
this->FormName=FormName;

}
//////////////////////////////////////////////////////////////////////////////

