//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Main.h"
#include "inifiles.hpp";
#include "CodeText.h"
#include "MasterPointer.h"
#include "PassForm.h"
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

VDB.clear();
CBDatabase->Clear();
ServerName=Ini->ReadString("Server","ServerName","");
int Num=Ini->ReadInteger("Server","NumServerBase",0);
for(int i=0;i<Num;i++)
{
String S="Base"+IntToStr(i+1);
String ServerBase=Ini->ReadString(S,"ServerDatabase","");
String Name=Ini->ReadString(S,"Name","");

//MClient->AddDatabase(ServerBase);

CDBItem DBI;
DBI.Name=Name;
DBI.ServerDB=ServerBase;
DBI.Num=i;
DBI.NumDatabase=Ini->ReadInteger(S,"NumDatabase",-1);
if(DBI.NumDatabase>=0)
{
CBDatabase->Items->Add(Name);
}

VDB.push_back(DBI);
}
CBDatabase->ItemIndex=0;

MClient=new Client(ClientSocket, ActionManager1);
MClient->Connect(Server, Port);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::RegForm_Form1Execute(TObject *Sender)
{
MClient->Act.WaitCommand=3;
MClient->Act.NextCommand=3;
String Par="PostRegForm_Form1";
MClient->Act.ParamComm.clear();
MClient->Act.ParamComm.push_back(Par);

MClient->RegForm(this->Name);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::RegForm_Form2Execute(TObject *Sender)
{
MClient->Act.NextCommand=4;
String Par="PostRegForm_Form2";
MClient->Act.ParamComm.clear();
MClient->Act.ParamComm.push_back(Par);

MClient->RegForm(Pass->Name);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::PostRegForm_Form1Execute(TObject *Sender)
{
int max=MClient->VForms.size()-1;
MClient->VForms[max]->IDF=StrToInt(MClient->Act.ParamComm[0]);

MClient->StartAction("RegForm_Form2");


}
//---------------------------------------------------------------------------

void __fastcall TForm1::PostRegForm_Form2Execute(TObject *Sender)
{
int max=MClient->VForms.size()-1;
MClient->VForms[max]->IDF=StrToInt(MClient->Act.ParamComm[0]);

MClient->StartAction("AspectsConnect");
}
//---------------------------------------------------------------------------


void __fastcall TForm1::AspectsConnectExecute(TObject *Sender)
{
MClient->Act.NextCommand=4;
String Par="DiaryConnect";
MClient->Act.ParamComm.clear();
MClient->Act.ParamComm.push_back(Par);

MClient->ConnectDatabase("Аспекты", true);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DiaryConnectExecute(TObject *Sender)
{
MClient->Act.NextCommand=0;
String Par="LoadLogins";
MClient->Act.ParamComm.clear();
MClient->Act.ParamComm.push_back(Par);

MClient->ConnectDatabase("Diary", true);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::LoadLoginsExecute(TObject *Sender)
{
//
Pass->ShowModal();
}
//---------------------------------------------------------------------------

