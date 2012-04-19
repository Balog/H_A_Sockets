// CLIENTSIMPL : Implementation of TClientsImpl (CoClass: Clients, Interface: IClients)

#include <vcl.h>
#pragma hdrstop

#include "CLIENTSIMPL.H"
#include "UServer.h"
#include "SysUtils.hpp"
#include "ServerClasses.h"
/////////////////////////////////////////////////////////////////////////////
// TClientsImpl

STDMETHODIMP TClientsImpl::Connect(BSTR CLName, BSTR Login, long* IDC,
  int* OK)
{
*IDC=Form1->MyClients->Connect(CLName, Login);
*OK=true;
  return S_OK;
}

STDMETHODIMP TClientsImpl::Disconnect_IDC(long* IDC, int* OK)
{
*OK=Form1->MyClients->Disconnect(*IDC);
  return S_OK;
}

STDMETHODIMP TClientsImpl::Disconnect_Name(BSTR CLName, int* OK)
{
*OK=Form1->MyClients->Disconnect(CLName);
  return S_OK;
}

STDMETHODIMP TClientsImpl::DisconnectAll(int* OK)
{
*OK=Form1->MyClients->DisconnectAll();
  return S_OK;
}

STDMETHODIMP TClientsImpl::get_CountClients(long* Count)
{
  try
  {
  *Count=Form1->MyClients->ClientCount();
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_IClients);
  }
  return S_OK;
};


STDMETHODIMP TClientsImpl::AddDatabase(long IDC, BSTR Name, long* IDDB)
{
*IDDB=Form1->MyClients->AddDatabase(IDC, Name);
  return S_OK;
}

STDMETHODIMP TClientsImpl::get_DatabaseConnect(long IDC, long IDDB,
  int* Connect)
{
  try
  {
*Connect=Form1->MyClients->GetDatabaseConnect(IDC, IDDB);
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_IClients);
  }
  return S_OK;
};


STDMETHODIMP TClientsImpl::set_DatabaseConnect(long IDC, long IDDB,
  int Connect)
{
  try
  {
  Form1->MyClients->SetDatabaseConnect(IDC, IDDB, Connect);
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_IClients);
  }
  return S_OK;
};



STDMETHODIMP TClientsImpl::set_SetCommandText(long IDC, long IDDB,
  BSTR Param3)
{
  try
  {
  Form1->MyClients->Command(IDC, IDDB)->CommandText=Param3;
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_IClients);
  }
  return S_OK;
};



STDMETHODIMP TClientsImpl::CommandExec(long IDC, long IDDB, int* OK)
{
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
Form1->MyClients->Command(IDC, IDDB)->Execute();
Form1->UpdateTempLogin();
*OK=true;
Form1->Block=-1;
}
else
{
*OK=false;
}


  return S_OK;
}


STDMETHODIMP TClientsImpl::MergeLogins(long IDC, long IDDB, int* OK,
  int* Wait)
{
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
*OK=Form1->MyClients->MergeLogins(IDC, IDDB);
Form1->UpdateTempLogin();
Form1->Block=-1;
*Wait=false;
}
else
{
*Wait=true;
}


  return S_OK;
}


STDMETHODIMP TClientsImpl::VerifyConnect(long IDC)
{
Form1->MyClients->VerifyConnect(IDC);
  return S_OK;
}


STDMETHODIMP TClientsImpl::get_NetError(VARIANT_BOOL* Value)
{
  try
  {
*Value=Form1->Net_Error;
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_IClients);
  }
  return S_OK;
};



STDMETHODIMP TClientsImpl::PrepareReadMetod(long IDC, BSTR NameDatabase,
  int* Ret)
{
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
*Ret=Form1->MyClients->PrepareReadMetod(IDC, NameDatabase);
Form1->Block=-1;
}
else
{
*Ret=false;
}
  return S_OK;
}


STDMETHODIMP TClientsImpl::MergeNodeBranch(long IDC, BSTR NameDatabase1,
  BSTR NameNode, BSTR NameBranch, BSTR NameDatabase2, BSTR NameTable,
  BSTR NameField, BSTR NameKey, BSTR Name, int* Ret)
{
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
*Ret=Form1->MyClients->MergeNodeBranch(IDC, NameDatabase1, NameNode, NameBranch, NameDatabase2, NameTable, NameField, NameKey, Name);
//ShowMessage("MergeNodeBranch "+IntToStr(IDC));
Form1->Block=-1;
}
else
{
*Ret=false;
}
  return S_OK;
}


STDMETHODIMP TClientsImpl::MergeNodeBranch_Sm(long IDC, BSTR NameDatabase1,
  BSTR NameNode, BSTR NameBranch, int* Ret)
{
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
*Ret=Form1->MyClients->MergeNodeBranch(IDC, NameDatabase1, NameNode, NameBranch);
Form1->Block=-1;
}
else
{
*Ret=false;
}

  return S_OK;
}


STDMETHODIMP TClientsImpl::MergeZn(long IDC, BSTR Database1,
  BSTR Database2, int* Ret)
{
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
Form1->MyClients->MergeZn(IDC, Database1, Database2);
*Ret=true;
Form1->Block=-1;
}
else
{
*Ret=false;
}
  return S_OK;
}


STDMETHODIMP TClientsImpl::SaveMetod(long IDC, BSTR Database, int* Ret)
{
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
Form1->MyClients->SaveMetod(IDC, Database);
*Ret=true;
Form1->Block=-1;
}
else
{
*Ret=false;
}


  return S_OK;
}


STDMETHODIMP TClientsImpl::SavePodr(long IDC, BSTR DB1, int* Ret)
{
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
Form1->MyClients->SavePodr(IDC, DB1);
*Ret=true;
Form1->Block=-1;
}
else
{
*Ret=false;
}

  return S_OK;
}


STDMETHODIMP TClientsImpl::SaveSit(long IDC, BSTR DB1, BSTR DB2, int* Ret)
{
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
Form1->MyClients->SaveSit(IDC, DB1, DB2);
*Ret=true;
Form1->Block=-1;
}
else
{
*Ret=false;
}

  return S_OK;
}


STDMETHODIMP TClientsImpl::Visible()
{
Form1->Visible=true;
  return S_OK;
}


STDMETHODIMP TClientsImpl::MergeAspectsMainSpec(long IDC, int* Ret)
{
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
Form1->MyClients->MergeAspectsMainSpec(IDC);
*Ret=true;
Form1->Block=-1;
}
else
{
*Ret=false;
}

  return S_OK;
}


STDMETHODIMP TClientsImpl::PrepareLoadAspects(BSTR NameDatabase, long IDC,
  long Podr, long Width1, long Width2, long Width3, long Width4, int* Ret)
{
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
Form1->MyClients->PrepareLoadAspects(IDC, NameDatabase, Podr, Width1, Width2, Width3, Width4);
*Ret=true;
Form1->Block=-1;
}
else
{
*Ret=false;
}

  return S_OK;
}


STDMETHODIMP TClientsImpl::get_Block(long IDC, int* Value)
{
  try
  {
  *Value=Form1->Block;
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_IClients);
  }
  return S_OK;
};


STDMETHODIMP TClientsImpl::set_Block(long IDC, int* Value)
{
  try
  {
  if(*Value==1 &  Form1->Block<0)
  {
  Form1->Block=IDC;
  }
  if(*Value==0 & Form1->Block==IDC)
  {
  Form1->Block=-1;
  }
  *Value=Form1->Block;
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_IClients);
  }
  return S_OK;
};



STDMETHODIMP TClientsImpl::PrepareMergeAspects(BSTR NameDatabase, long IDC,
  long NumLogin, long W1, long W2, long W3, long W4, int* Ret)
{
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
Form1->MyClients->PrepareMergeAspects(NameDatabase, IDC, NumLogin, W1, W2, W3, W4);
*Ret=true;
Form1->Block=-1;
}
else
{
*Ret=false;
}
  return S_OK;
}


STDMETHODIMP TClientsImpl::WriteDiary(BSTR Comp, BSTR Login, BSTR Type,
  BSTR Name, BSTR Prim)
{
Form1->MyClients->WriteDiaryEvent(Comp, Login, Type, Name, Prim);
  return S_OK;
}


STDMETHODIMP TClientsImpl::get_LicenseCount(BSTR DBName, long* Value)
{
  try
  {
  *Value=Form1->GetLicenseCount(DBName);
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_IClients);
  }
  return S_OK;
};



STDMETHODIMP TClientsImpl::PrepareMergeAspectsH(BSTR NameDatabase,
  long IDC, long NumLogin, long W1, long W2, long W3, long W4, int* Ret)
{ 
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
Form1->MyClients->PrepareMergeAspectsH(NameDatabase, IDC, NumLogin, W1, W2, W3, W4);
*Ret=true;
Form1->Block=-1;
}
else
{
*Ret=false;
}
  return S_OK;
}


STDMETHODIMP TClientsImpl::PrepareLoadAspectsH(BSTR NameDatabase, long IDC,
  long Podr, long Width1, long Width2, long Width3, long Width4, int* Ret)
{ 
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
Form1->MyClients->PrepareLoadAspectsH(IDC, NameDatabase, Podr, Width1, Width2, Width3, Width4);
*Ret=true;
Form1->Block=-1;
}
else
{
*Ret=false;
}

  return S_OK;
}


STDMETHODIMP TClientsImpl::MergeAspectsMainSpecH(long IDC, int* Ret)
{ 
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
Form1->MyClients->MergeAspectsMainSpecH(IDC);
*Ret=true;
Form1->Block=-1;
}
else
{
*Ret=false;
}

  return S_OK;
}


