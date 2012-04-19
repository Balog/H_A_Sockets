// FORMIMPL.H : Declaration of the TFormImpl

#ifndef FormImplH
#define FormImplH

#define ATL_APARTMENT_THREADED

#include "Server_TLB.H"


/////////////////////////////////////////////////////////////////////////////
// TFormImpl     Implements IForm, default interface of Form
// ThreadingModel : Apartment
// Dual Interface : TRUE
// Event Support  : FALSE
// Default ProgID : Server.Form
// Description    : 
/////////////////////////////////////////////////////////////////////////////
class ATL_NO_VTABLE TFormImpl : 
  public CComObjectRootEx<CComSingleThreadModel>,
  public CComCoClass<TFormImpl, &CLSID_Form>,
  public IDispatchImpl<IForm, &IID_IForm, &LIBID_Server>
{
public:
  TFormImpl()
  {
  }

  // Data used when registering Object 
  //
  DECLARE_THREADING_MODEL(otApartment);
  DECLARE_PROGID("Server.Form");
  DECLARE_DESCRIPTION("");

  // Function invoked to (un)register object
  //
  static HRESULT WINAPI UpdateRegistry(BOOL bRegister)
  {
    TTypedComServerRegistrarT<TFormImpl> 
    regObj(GetObjectCLSID(), GetProgID(), GetDescription());
    return regObj.UpdateRegistry(bRegister);
  }


BEGIN_COM_MAP(TFormImpl)
  COM_INTERFACE_ENTRY(IForm)
  COM_INTERFACE_ENTRY2(IDispatch, IForm)
END_COM_MAP()

// IForm
public:
 
  STDMETHOD(RegForm(long IDC, BSTR NameForm, long* IDF, int* OK));
  STDMETHOD(UnregForm_IDF(long IDC, long IDF, int* OK));
  STDMETHOD(get_CountForm(long IDC, long* Count));
  STDMETHOD(UnregAll(long IDC, int* OK));
  STDMETHOD(UnregForm_Name(long IDC, BSTR NameForm, int* OK));
};

#endif //FormImplH
