//---------------------------------------------------------------------------


#pragma hdrstop

#include "ClientClass.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)
Client::Client(TClientSocket *ClientSocket, TActionManager *ActMan)
{
Socket=ClientSocket;
ActionManager=ActMan;
}
//********************************************************************
Client::~Client()
{

}
//********************************************************************
void Client::Connect(String ServerName, int Port)
{
//��������� ����� �� ������� 1

Act.WaitCommand=1;
Socket->Address=ServerName;
Socket->Host=ServerName;
Socket->Port=Port;
Socket->Active=true;
}
//********************************************************************
void Client::CommandExec(int Comm, vector<String>Parameters)
{
if(Act.WaitCommand==Comm)
{
switch(Comm)
{
 case 1:
 {

  //�������� �� ������� 1
  String Mess="Command:1;2|"+GetIP()+"|"+Application->ExeName;
  //���������� ��������� ��� ��������� ������� ����� 2
  Act.WaitCommand=2;
  Socket->Socket->SendText(Mess);
  break;
 }
 case 2:
 {
 StartAction("RegForm");
 //ActionManager->Actions[0]->Execute();
 break;
 }
}
}
else
{
 ShowMessage("��������� ������� "+IntToStr(Act.WaitCommand)+" � ������ ����� �� ������� "+IntToStr(Comm));
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
    else MessageBox(0,"�� �� � ����. � IP ������ � ��� ���.",0,0);
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
//��������� ����������� ����� �� �������, ������� 3
  String Mess="Command:3;1|"+FormName;
  //���������� ��������� ��� ��������� ������� ����� 3
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
/////////////////////////////////////////////////////////////////////////////
Form::Form()
{

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

