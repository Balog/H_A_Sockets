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

Path=ExtractFilePath(Application->ExeName);
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
C->LastCommand=1;
Cl->VClients.push_back(C);

//Передача команды на запрос IP адреса
String Mess="Command:1;0|";
/*
for(unsigned int i=0;i<Cl->VClients.size();i++)
{
 if(Socket==Cl->VClients[i]->Socket)
 {
  Cl->VClients[i]->Socket->SendText(Mess);
 }
}
*/
Socket->SendText(Mess);
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
void __fastcall TForm1::ServerSocketClientError(TObject *Sender,
      TCustomWinSocket *Socket, TErrorEvent ErrorEvent, int &ErrorCode)
{
//ErrorCode=0;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ServerSocketClientRead(TObject *Sender,
      TCustomWinSocket *Socket)
{
String Mess=Socket->ReceiveText();
if(Mess.Length()!=0)
{
//Выделение параметров
int N0=Mess.Pos(":");
int N=Mess.Pos(";");
int N1=Mess.Pos("|");
//String Num=RText.SubString(9,1);
int Comm, NumPar;
String Nick;
String SS=Mess.SubString(N0+1,N-N0-1);
Comm=StrToInt(SS);
NumPar=StrToInt(Mess.SubString(N+1,N1-N-1));
String S=Mess.SubString(N1+1,Mess.Length());
Parameters.clear();
for(int i=0;i<NumPar;i++)
{
String Par;
 int M=S.Pos("|");
 if(M>0)
 {
 Par=S.SubString(1,M-1);
 }
 else
 {
 Par=S;
 }
 S=S.SubString(M+1,S.Length());
 Parameters.push_back(Par);


}
for(unsigned int i=0;i<Cl->VClients.size();i++)
{
 if(Cl->VClients[i]->Socket==Socket)
 {
  Cl->VClients[i]->CommandExec(Comm, Parameters);
 }
 Cl->IVC++;
}
}
}
//---------------------------------------------------------------------------

