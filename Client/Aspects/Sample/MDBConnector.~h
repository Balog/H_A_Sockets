//---------------------------------------------------------------------------
#include <vcl.h>
#include <ADODB.hpp>
#ifndef MDBConnectorH
#define MDBConnectorH
//---------------------------------------------------------------------------
typedef void __fastcall (__closure *TArchTimer)(System::TObject *Sender, int Time);

class MDBConnector : public TADOConnection
{
private:
String pPatchDB;
String pNameDB;
String pPathBackup;
String pNamebackUp;
bool pAutoPackDB;
TComponent* pOwner;

String CreateName();
int FileAge(String File);

TArchTimer FArchTimer;
String GetDBPatch();
String GetDBName();
__published:
__property TArchTimer OnArchTimer = {read= FArchTimer, write= FArchTimer};

public:
__property String NameDB={read= GetDBName};
__property String PatchDB={read= GetDBPatch};
MDBConnector(String PatchDB, String NameDB, String PathBackup, bool AutoConnect, bool AutoPackDB, TComponent* Owner);
MDBConnector(String PatchDB, String NameDB, TComponent* Owner);
__fastcall ~MDBConnector();

void ClearArchive(int Age);

bool PackDB();
bool Backup();
bool Backup(String PatchBackup);
bool Backup(String PatchBackup, String NameBackUp);

void ConnectDB();
void DisconnectDB();

void SetPatchBackUp(String Patch);


};
#endif
 