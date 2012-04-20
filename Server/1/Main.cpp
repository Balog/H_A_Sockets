//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Main.h"
#include "inifiles.hpp";
#include "MasterPointer.h"
#include "CodeText.h"
#include "MDBConnector.h"

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
Block=-1;

PageControl1->ActivePageIndex=0;

String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Server.ini");
int NumBase=Ini->ReadInteger("Main","NumDatabases",0);
int DiaryStoreDays=Ini->ReadInteger("Main","DiaryStoreDays",0);
int Days=Ini->ReadInteger("Main","StoreArchive",0);
int Port=Ini->ReadInteger("Main","Port",2000);
ServerSocket->Port=Port;

Cl=new Clients(this, "");


ServerSocket->Active=true;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ServerSocketClientConnect(TObject *Sender,
      TCustomWinSocket *Socket)
{
Client *C=new Client();
C->Socket=Socket;
Cl->VClients.push_back(C);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ServerSocketClientDisconnect(TObject *Sender,
      TCustomWinSocket *Socket)
{
Cl->IVC=Cl->VClients.begin();
for(unsigned int i=0;i<Cl->VClients.size();i++)
{
 if(Cl->VClients[i]->Socket==Socket)
 {
  Cl->VClients.erase(Cl->IVC);
 }
 Cl->IVC++;
}
}
//---------------------------------------------------------------------------
