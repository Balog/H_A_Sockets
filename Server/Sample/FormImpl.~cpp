// FORMIMPL : Implementation of TFormImpl (CoClass: Form, Interface: IForm)

#include <vcl.h>
#pragma hdrstop

#include "FORMIMPL.H"
#include "CLIENTSIMPL.H"
#include "UServer.h"
/////////////////////////////////////////////////////////////////////////////
// TFormImpl

STDMETHODIMP TFormImpl::RegForm(long IDC, BSTR NameForm, long* IDF,
  int* OK)
{
*IDF=Form1->MyClients->GetClient(IDC)->RegForm(NameForm);
*OK=true;
//Form1->Label2->Caption=NameForm;
  return S_OK;
}


STDMETHODIMP TFormImpl::UnregForm_IDF(long IDC, long IDF, int* OK)
{

*OK=Form1->MyClients->GetClient(IDC)->UnRegForm(IDF);
  return S_OK;
}

STDMETHODIMP TFormImpl::get_CountForm(long IDC, long* Count)
{
  try
  {
  *Count=Form1->MyClients->GetClient(IDC)->FormCount();
  }
  catch(Exception &e)
  {
    return Error(e.Message.c_str(), IID_IForm);
  }
  return S_OK;
};


STDMETHODIMP TFormImpl::UnregAll(long IDC, int* OK)
{
*OK=Form1->MyClients->GetClient(IDC)->UnRegAll();
  return S_OK;
}

STDMETHODIMP TFormImpl::UnregForm_Name(long IDC, BSTR NameForm, int* OK)
{
*OK=Form1->MyClients->GetClient(IDC)->UnRegForm(NameForm);
  return S_OK;
}

