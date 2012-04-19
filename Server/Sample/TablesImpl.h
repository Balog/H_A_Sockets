// TABLESIMPL.H : Declaration of the TTablesImpl

#ifndef TablesImplH
#define TablesImplH

#define ATL_APARTMENT_THREADED

#include "Server_TLB.H"


/////////////////////////////////////////////////////////////////////////////
// TTablesImpl     Implements ITables, default interface of Tables
// ThreadingModel : Apartment
// Dual Interface : TRUE
// Event Support  : FALSE
// Default ProgID : Server.Tables
// Description    : 
/////////////////////////////////////////////////////////////////////////////
class ATL_NO_VTABLE TTablesImpl : 
  public CComObjectRootEx<CComSingleThreadModel>,
  public CComCoClass<TTablesImpl, &CLSID_Tables>,
  public IDispatchImpl<ITables, &IID_ITables, &LIBID_Server>
{
public:
  TTablesImpl()
  {
  }

  // Data used when registering Object 
  //
  DECLARE_THREADING_MODEL(otApartment);
  DECLARE_PROGID("Server.Tables");
  DECLARE_DESCRIPTION("");

  // Function invoked to (un)register object
  //
  static HRESULT WINAPI UpdateRegistry(BOOL bRegister)
  {
    TTypedComServerRegistrarT<TTablesImpl> 
    regObj(GetObjectCLSID(), GetProgID(), GetDescription());
    return regObj.UpdateRegistry(bRegister);
  }


BEGIN_COM_MAP(TTablesImpl)
  COM_INTERFACE_ENTRY(ITables)
  COM_INTERFACE_ENTRY2(IDispatch, ITables)
END_COM_MAP()

// ITables
public:
 
  STDMETHOD(CreateTable(long IDC, long IDDB, long IDF, long* IDT));
  STDMETHOD(DeleteTable(long IDC, long IDF, long IDT, int* OK));
  STDMETHOD(Active(VARIANT_BOOL Value, long IDC, long IDF, long IDT,
      int* OK));
  STDMETHOD(get_CommandText(long IDC, long IDF, long IDT, BSTR* Text));
  STDMETHOD(set_CommandText(long IDC, long IDF, long IDT, BSTR Text));
  STDMETHOD(Moving(long Comand, long IDC, long IDF, long IDT));
  STDMETHOD(get_BOF_EOF(long Comand, long IDC, long IDF, long IDT,
      VARIANT_BOOL* Value));
  STDMETHOD(get_FieldByName_Variant(BSTR FieldName, long IDC, long IDF,
      long IDT, VARIANT* Value));
  STDMETHOD(set_FieldByName_Variant(BSTR FieldName, long IDC, long IDF,
      long IDT, VARIANT Value));
  STDMETHOD(EditingRecord(long Comand, long IDC, long IDF, long IDT));
  STDMETHOD(get_RecordCount(long IDC, long IDF, long IDT, long* Value));
  STDMETHOD(get_RecNo(long IDC, long IDF, long IDT, long* Value));
  STDMETHOD(Move(long IDC, long IDF, long IDT, long Step));
  STDMETHOD(get_Locate(long IDC, long IDF, long IDT, BSTR FieldName,
      BSTR Key, long SearchParam, VARIANT_BOOL* OK));
  STDMETHOD(Delete(long IDC, long IDF, long IDT));
  STDMETHOD(get_Fields(long NumField, long IDC, long IDF, long IDT,
      VARIANT* Value));
  STDMETHOD(set_Fields(long NumField, long IDC, long IDF, long IDT,
      VARIANT Value));
  STDMETHOD(get_FieldsCount(long IDC, long IDF, long IDT, long* Value));
};

#endif //TablesImplH
