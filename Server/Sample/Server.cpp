//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//---------------------------------------------------------------------------
#include <atl\atlmod.h>
#include <atl\atlmod.h>
#include <atl\atlmod.h>
#include <atl\atlmod.h>
#include "ClientsImpl.h"

#include "TablesImpl.h"
#include "FormImpl.h"
USEFORM("UServer.cpp", Form1);
USEFORM("EditLogin.cpp", EditLogins);
USEFORM("Diary.cpp", FDiary);
//---------------------------------------------------------------------------
TComModule _ProjectModule(0 /*InitATLServer*/);
TComModule &_Module = _ProjectModule;

// The ATL Object map holds an array of _ATL_OBJMAP_ENTRY structures that
// described the objects of your OLE server. The MAP is handed to your
// project's CComModule-derived _Module object via the Init method.
//
BEGIN_OBJECT_MAP(ObjectMap)
  OBJECT_ENTRY(CLSID_Clients, TClientsImpl)

  OBJECT_ENTRY(CLSID_Tables, TTablesImpl)
  OBJECT_ENTRY(CLSID_Form, TFormImpl)
END_OBJECT_MAP()
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {
                  CreateMutex(NULL, true, "ServerMutex");
                  int Res = GetLastError();
                  if(Res == ERROR_ALREADY_EXISTS)
                  {
                  String CompName;
                  char buffer[32];
                  DWORD size;
                  size=sizeof(buffer);
                  GetComputerName(buffer,&size);
                  CompName=buffer;

                   CoInitialize(NULL);

                  GUID ClassID = Comobj::StringToGUID("Server.Clients");
                  Variant Server  = (IUnknown*)CreateRemoteComObject(CompName, ClassID);

                  Server.OleProcedure("Visible");

                  Application->Terminate();
                  }
                  if(Res == ERROR_INVALID_HANDLE)
                  {
                  ShowMessage("������ � ����� mutex object");
                  Application->Terminate();
                  }
                 if(Res==0)
                 {
                 Application->Initialize();
                 Application->CreateForm(__classid(TForm1), &Form1);
                 Application->CreateForm(__classid(TEditLogins), &EditLogins);
                 Application->CreateForm(__classid(TFDiary), &FDiary);
                 Application->Run();
                 }
        }
        catch (Exception &exception)
        {
                 Application->ShowException(&exception);
        }
        catch (...)
        {
                 try
                 {
                         throw Exception("");
                 }
                 catch (Exception &exception)
                 {
                         Application->ShowException(&exception);
                 }
        }
        return 0;
}
//---------------------------------------------------------------------------
