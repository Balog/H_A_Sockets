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