//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Main.h"
#include "inifiles.hpp";
#include "CodeText.h"
#include "MasterPointer.h"
using namespace std;
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Admin.ini");
String Server=Ini->ReadString("Main","Server","localhost");
int Port=Ini->ReadInteger("Main","Port",2000);
int AbsPathAspect=Ini->ReadInteger("Main","AbsPathAspect",1);
String AspectBase=Ini->ReadString("Main","AspectBase","");
if(AbsPathAspect==0)
{
AdminDatabase=Path+AspectBase;
}
else
{
AdminDatabase= AspectBase;
}

Database=new MDBConnector(ExtractFilePath(AdminDatabase), ExtractFileName(AdminDatabase), this);
Database->SetPatchBackUp("Archive");


MP<TADODataSet>Roles(this);
Roles->CommandText="Select * from Roles order by Num";
Roles->Connection=Database;
Roles->Active=true;
Role->Clear();
for(Roles->First();!Roles->Eof;Roles->Next())
{
Role->Items->Add(Roles->FieldByName("Name")->AsString);
}
Role->ItemIndex=0;

int Days=Ini->ReadInteger("Main","StoreArchive",0);
Database->ClearArchive(Days);
Database->PackDB();
Database->Backup("Archive");

ClientSocket->Address=Server;
ClientSocket->Host=Server;
ClientSocket->Port=Port;
ClientSocket->Active=true;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ClientSocketRead(TObject *Sender,
      TCustomWinSocket *Socket)
{
//
String Mess=Socket->ReceiveText();
if (Mess.Length()!=0)
{
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
 int M=S.Pos(",");
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
CommandExec(Comm, Parameters);
//}
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ClientSocketWrite(TObject *Sender,
      TCustomWinSocket *Socket)
{
//
}
//---------------------------------------------------------------------------
void TForm1::CommandExec(int Comm, vector<String>Parameters)
{
switch(Comm)
{
 case 1:
 {
  //ShowMessage(GetIP());
  String Mess="Command:1;1|"+GetIP();
  ClientSocket->Socket->SendText(Mess);
  break;
 }
 case 2:
 {

 break;
 }
}
}
//------------------------------------------------------------------------
String TForm1::GetIP()
{
    String addr="";

    const int WSVer = 0x101;
    WSAData wsaData;
    hostent *h;
    char Buf[128];
    if (WSAStartup(WSVer, &wsaData) == 0)
    {
    if (gethostname(&Buf[0], 128) == 0)
    {
    h = gethostbyname(&Buf[0]);
    if (h != NULL)
    {
    //ShowMessage(inet_ntoa (*(reinterpret_cast<in_addr *>(*(h->h_addr_list)))));
    addr=inet_ntoa (*(reinterpret_cast<in_addr *>(*(h->h_addr_list))));

    }
    else MessageBox(0,"Вы не в сети. И IP адреса у вас нет.",0,0);
    }
    WSACleanup;
    }


    return addr;
}
//-------------------------------------------------------------------------
