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
MClient->CommandExec(Comm, Parameters);
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
void __fastcall TForm1::FormShow(TObject *Sender)
{
Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Admin.ini");
String Server=Ini->ReadString("Main","Server","localhost");
int Port=Ini->ReadInteger("Main","Port",2000);
MClient=new Client(ClientSocket, ActionManager1);
MClient->Connect(Server, Port);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::RegFormExecute(TObject *Sender)
{
//ShowMessage(this->Name);

MClient->RegForm(this->Name);
}
//---------------------------------------------------------------------------

