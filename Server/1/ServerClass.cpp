//---------------------------------------------------------------------------


#pragma hdrstop

#include "ServerClass.h"

//---------------------------------------------------------------------------

#pragma package(smart_init)

Clients::Clients(TComponent* Owner, String DiaryBase)
{
pOwner=Owner;
this->DiaryBase=DiaryBase;
}
//-------------------------------------------------------
Clients::~Clients()
{
for(unsigned int i=0;i<VClients.size();i++)
{
delete VClients[i];
}
VClients.clear();
}
//-------------------------------------------------------
//********************************************************
Client::Client()
{

}
//**************************************************
Client::~Client()
{

}
//***************************************************
void Client::CommandExec(int Comm, vector<String>Parameters)
{
if(LastCommand==Comm | LastCommand==0)
{
 switch(Comm)
 {
  case 1:
  {
   //����� �� ������� ������� IP � �������
   this->IP=Parameters[0];
   this->AppPatch=Parameters[1];
   //�������� ������� ���������� � ���������� ������
   //����� �� ���� �� ������������
   LastCommand=0;
   this->Socket->SendText("Command:2;0|");
  break;
  }
  case 3:
  {
  //������������ ����������� ����� �� �������
   ShowMessage(Parameters[0]);
   break;
  }
 }

}
}
//---------------------------------------------------------------------------