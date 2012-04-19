// CLIENTSIMPL.H : Declaration of the TClientsImpl

#ifndef ClientsImplH
#define ClientsImplH

#define ATL_APARTMENT_THREADED

#include "Server_TLB.H"


/////////////////////////////////////////////////////////////////////////////
// TClientsImpl     Implements IClients, default interface of Clients
// ThreadingModel : Apartment
// Dual Interface : TRUE
// Event Support  : FALSE
// Default ProgID : Server.Clients
// Description    : 
/////////////////////////////////////////////////////////////////////////////
class ATL_NO_VTABLE TClientsImpl : 
  public CComObjectRootEx<CComSingleThreadModel>,
  public CComCoClass<TClientsImpl, &CLSID_Clients>,
  public IDispatchImpl<IClients, &IID_IClients, &LIBID_Server>
{
public:
  TClientsImpl()
  {
  }

  // Data used when registering Object 
  //
  DECLARE_THREADING_MODEL(otApartment);
  DECLARE_PROGID("Server.Clients");
  DECLARE_DESCRIPTION("");

  // Function invoked to (un)register object
  //
  static HRESULT WINAPI UpdateRegistry(BOOL bRegister)
  {
    TTypedComServerRegistrarT<TClientsImpl> 
    regObj(GetObjectCLSID(), GetProgID(), GetDescription());
    return regObj.UpdateRegistry(bRegister);
  }


BEGIN_COM_MAP(TClientsImpl)
  COM_INTERFACE_ENTRY(IClients)
  COM_INTERFACE_ENTRY2(IDispatch, IClients)
END_COM_MAP()

// IClients
public:
 
  STDMETHOD(Connect(BSTR CLName, BSTR Login, long* IDC, int* OK));
  STDMETHOD(Disconnect_IDC(long* IDC, int* OK));
  STDMETHOD(Disconnect_Name(BSTR CLName, int* OK));
  STDMETHOD(DisconnectAll(int* OK));
  STDMETHOD(get_CountClients(long* Count));
  STDMETHOD(AddDatabase(long IDC, BSTR Name, long* IDDB));
  STDMETHOD(get_DatabaseConnect(long IDC, long IDDB, int* Connect));
  STDMETHOD(set_DatabaseConnect(long IDC, long IDDB, int Connect));
  STDMETHOD(set_SetCommandText(long IDC, long IDDB, BSTR Param3));
  STDMETHOD(CommandExec(long IDC, long IDDB, int* OK));
  STDMETHOD(MergeLogins(long IDC, long IDDB, int* OK, int* Wait));
  STDMETHOD(VerifyConnect(long IDC));
  STDMETHOD(get_NetError(VARIANT_BOOL* Value));
  STDMETHOD(PrepareReadMetod(long IDC, BSTR NameDatabase, int* Ret));
  STDMETHOD(MergeNodeBranch(long IDC, BSTR NameDatabase1, BSTR NameNode,
      BSTR NameBranch, BSTR NameDatabase2, BSTR NameTable, BSTR NameField,
      BSTR NameKey, BSTR Name, int* Ret));
  STDMETHOD(MergeNodeBranch_Sm(long IDC, BSTR NameDatabase1, BSTR NameNode,
      BSTR NameBranch, int* Ret));
  STDMETHOD(MergeZn(long IDC, BSTR Database1, BSTR Database2, int* Ret));
  STDMETHOD(SaveMetod(long IDC, BSTR Database, int* Ret));
  STDMETHOD(SavePodr(long IDC, BSTR DB1, int* Ret));
  STDMETHOD(SaveSit(long IDC, BSTR DB1, BSTR DB2, int* Ret));
  STDMETHOD(Visible());
  STDMETHOD(MergeAspectsMainSpec(long IDC, int* Ret));
  STDMETHOD(PrepareLoadAspects(BSTR NameDatabase, long IDC, long Podr,
      long Width1, long Width2, long Width3, long Width4, int* Ret));
  STDMETHOD(get_Block(long IDC, int* Value));
  STDMETHOD(set_Block(long IDC, int* Value));
  STDMETHOD(PrepareMergeAspects(BSTR NameDatabase, long IDC, long NumLogin,
      long W1, long W2, long W3, long W4, int* Ret));
  STDMETHOD(WriteDiary(BSTR Comp, BSTR Login, BSTR Type, BSTR Name,
      BSTR Prim));
  STDMETHOD(get_LicenseCount(BSTR DBName, long* Value));
  STDMETHOD(PrepareMergeAspectsH(BSTR NameDatabase, long IDC, long NumLogin,
      long W1, long W2, long W3, long W4, int* Ret));
  STDMETHOD(PrepareLoadAspectsH(BSTR NameDatabase, long IDC, long Podr,
      long Width1, long Width2, long Width3, long Width4, int* Ret));
  STDMETHOD(MergeAspectsMainSpecH(long IDC, int* Ret));
};

#endif //ClientsImplH
