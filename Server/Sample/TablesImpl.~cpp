// TABLESIMPL : Implementation of TTablesImpl (CoClass: Tables, Interface: ITables)

#include <vcl.h>
#pragma hdrstop

#include "TABLESIMPL.H"
#include "UServer.h"
#include "ServerClasses.h"

/////////////////////////////////////////////////////////////////////////////
// TTablesImpl

STDMETHODIMP TTablesImpl::CreateTable(long IDC, long IDDB, long IDF,
  long* IDT)
{
*IDT=Form1->MyClients->CreateTable(IDC, IDDB, IDF);
return S_OK;
}

STDMETHODIMP TTablesImpl::DeleteTable(long IDC, long IDF, long IDT,
  int* OK)
{
*OK=Form1->MyClients->DeleteTable(IDC, IDF, IDT);
  return S_OK;
}


STDMETHODIMP TTablesImpl::Active(VARIANT_BOOL Value, long IDC, long IDF,
  long IDT, int* OK)
{
if(Form1->Block<0 | Form1->Block==IDC)
{
Form1->Block=IDC;
Form1->MyClients->Active(Value, IDC, IDF, IDT);
Form1->UpdateTempLogin();
*OK= true;
Form1->Block=-1;
}
else
{
*OK=false;
}


  return S_OK;
}


STDMETHODIMP TTablesImpl::get_CommandText(long IDC, long IDF, long IDT,
  BSTR* Text)
{
  try
  {
  String S=Form1->MyClients->GetCommandText(IDC, IDF, IDT);
  //*Value=WideString(S.c_str()).c_bstr()
  *Text=WideString(S.c_str()).c_bstr();
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_ITables);
  }
  return S_OK;
};


STDMETHODIMP TTablesImpl::set_CommandText(long IDC, long IDF, long IDT,
  BSTR Text)
{
  try
  {
  Form1->MyClients->SetCommandText(Text, IDC, IDF, IDT);
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_ITables);
  }
  return S_OK;
};



STDMETHODIMP TTablesImpl::Moving(long Comand, long IDC, long IDF, long IDT)
{
Form1->MyClients->Moving(Comand, IDC, IDF, IDT);
  return S_OK;
}


STDMETHODIMP TTablesImpl::get_BOF_EOF(long Comand, long IDC, long IDF,
  long IDT, VARIANT_BOOL* Value)
{
  try
  {
  *Value=Form1->MyClients->BOF_EOF(Comand, IDC, IDF, IDT);
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_ITables);
  }
  return S_OK;
};



STDMETHODIMP TTablesImpl::get_FieldByName_Variant(BSTR FieldName, long IDC,
  long IDF, long IDT, VARIANT* Value)
{
  try
  {
//  *Value=Form1->Tab->FieldByName(FieldName)->Value;
*Value=Form1->MyClients->FieldByName(FieldName, IDC, IDF, IDT);
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_ITables);
  }
  return S_OK;
};


STDMETHODIMP TTablesImpl::set_FieldByName_Variant(BSTR FieldName, long IDC,
  long IDF, long IDT, VARIANT Value)
{
  try
  {
//  Form1->Tab->FieldByName(FieldName)->Value=(Variant)Value;
Form1->MyClients->FieldByName((Variant)Value, FieldName, IDC, IDF, IDT);
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_ITables);
  }
  return S_OK;
};



STDMETHODIMP TTablesImpl::EditingRecord(long Comand, long IDC, long IDF,
  long IDT)
{
Form1->MyClients->EditingRecord(Comand, IDC, IDF, IDT);
  return S_OK;
}


STDMETHODIMP TTablesImpl::get_RecordCount(long IDC, long IDF, long IDT,
  long* Value)
{
  try
  {
  *Value=Form1->MyClients->RecordCount(IDC, IDF, IDT);
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_ITables);
  }
  return S_OK;
};



STDMETHODIMP TTablesImpl::get_RecNo(long IDC, long IDF, long IDT,
  long* Value)
{
  try
  {
  *Value=Form1->MyClients->RecNo(IDC, IDF, IDT);
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_ITables);
  }
  return S_OK;
};



STDMETHODIMP TTablesImpl::Move(long IDC, long IDF, long IDT, long Step)
{
Form1->MyClients->Move(Step, IDC, IDF, IDT);
  return S_OK;
}


STDMETHODIMP TTablesImpl::get_Locate(long IDC, long IDF, long IDT,
  BSTR FieldName, BSTR Key, long SearchParam, VARIANT_BOOL* OK)
{
  try
  {
  *OK=Form1->MyClients->Locate(IDC, IDF, IDT, FieldName, Key, SearchParam);
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_ITables);
  }
  return S_OK;
};



STDMETHODIMP TTablesImpl::Delete(long IDC, long IDF, long IDT)
{
Form1->MyClients->Delete(IDC, IDF, IDT);
  return S_OK;
}


STDMETHODIMP TTablesImpl::get_Fields(long NumField, long IDC, long IDF,
  long IDT, VARIANT* Value)
{
  try
  {
*Value=Form1->MyClients->Fields(NumField, IDC, IDF, IDT);
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_ITables);
  }
  return S_OK;
};


STDMETHODIMP TTablesImpl::set_Fields(long NumField, long IDC, long IDF,
  long IDT, VARIANT Value)
{
  try
  {
Form1->MyClients->Fields((Variant)Value, NumField, IDC, IDF, IDT);
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_ITables);
  }
  return S_OK;
};



STDMETHODIMP TTablesImpl::get_FieldsCount(long IDC, long IDF, long IDT,
  long* Value)
{
  try
  {
*Value=Form1->MyClients->FieldsCount(IDC, IDF, IDT);
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_ITables);
  }
  return S_OK;
};




