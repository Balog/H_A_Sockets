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


Cl->VBases.clear();
Base->Clear();
for(int i=0;i<NumBase;i++)
{
 String NameSect="Base"+IntToStr(i+1);
 BaseItem BI;

 BI.Name=Ini->ReadString(NameSect,"Name","");

int MS=Ini->ReadInteger(NameSect,"MainSpec",0);
if(MS==0)
{
BI.MainSpec=false;
 Base->Items->Add(BI.Name);

 String Lic=Ini->ReadString(NameSect,"License","");

 FILE *F;
String File=ExtractFilePath(Application->ExeName)+Lic;
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),"�� ���������", "��� �� ���������", "��������", "�������� ���� ��������", File);

if ((F = fopen(File.c_str(), "rt")) == NULL)
{
 //ShowMessage("���� �������� �� ������� �������");
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),"�� ���������", "��� �� ���������" "������ ��������", "���� ��������� �������", File);

 BI.LicCount=0;
}
else
{
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),"�� ���������", "��� �� ���������" "��������", "���� �������� ������", File); //ShowMessage("���� �������� ������");
 char S[256];
 fgets(S, 257, F);
 String Text=S;
 int LC=AnalizLic(Text, BI.Name);
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),"�� ���������", "��� �� ���������" "��������", "���� �������� �����������", "DB: "+BI.Name+" NumUsers="+IntToStr(LC)); //ShowMessage("���� �������� ������");

 BI.LicCount=LC;

}

}
else
{
BI.MainSpec=true;
}
int AbsPath=Ini->ReadInteger(NameSect, "AbsPath",0);

if(AbsPath==0)
{
 BI.Patch=Path+Ini->ReadString(NameSect,"Base","");
}
else
{
 BI.Patch= Ini->ReadString(NameSect,"Base","");
}
String FN=ExtractFileName(BI.Patch);
String FN1=FN.SubString(0,FN.Length()-ExtractFileExt(BI.Patch).Length());
BI.FileName=FN1;


MDBConnector* ADOConn=new MDBConnector(ExtractFilePath(BI.Patch), FN1, this);
ADOConn->SetPatchBackUp("Archive");
BI.Database=ADOConn;
//BI.Database->Connected=true;
Cl->VBases.push_back(BI);

ADOConn->ClearArchive(Days);
ADOConn->PackDB();
ADOConn->Backup("Archive");

if(BI.Name=="Diary")
{
MP<TADOCommand>Comm(this);
Comm->Connection=ADOConn;
Comm->CommandText="DELETE Now()-[Date_Time] AS Days FROM Events WHERE (((Now()-[Date_Time])>"+IntToStr(DiaryStoreDays)+"));";
Comm->Execute();
ADOConn->Connected=false;
}
//delete ADOConn;
}

Base->ItemIndex=0;

String File=GetFileDatabase(Base->Text);
//int Num;
//int LicCount;
for(unsigned int i=0;i<Cl->VBases.size();i++)
{
 if(Cl->VBases[i].FileName==File)
 {
Database->Connected=false;
Database->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+Cl->VBases[i].Patch+";Persist Security Info=False";
Database->LoginPrompt=false;
//Database->Connected=true;

  break;
 }
}

/*

*/




ServerSocket->Active=true;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ServerSocketClientConnect(TObject *Sender,
      TCustomWinSocket *Socket)
{
Client *C=new Client(Cl);
C->Socket=Socket;
C->LastCommand=1;
Cl->VClients.push_back(C);

//�������� ������� �� ������ IP ������
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
//��������� ����������
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
int TForm1::AnalizLic(String Text, String Pass)
{
int Ret=0;
CodeText *CT=new CodeText();
String DText=CT->EnCrypt(Text,Pass, 256);
String DCopy=DText;
int Pos1=DText.AnsiPos("+");
String Key1=DText.SubString(1,Pos1-1);
DText=DText.SubString(Pos1+1,DText.Length());
int Pos2=DText.AnsiPos("+");
String Key2=DText.SubString(1,Pos2-1);

String Key1_2="+"+Key1+"+"+Key2;
int Pos3=DCopy.AnsiPos(Key1_2);

//ShowMessage("Pos3="+IntToStr(Pos3));
if(Pos3!=0)
{
//�������� �����
if(Key1=="�������������" & Key2=="������")
{
//������������� ��������
Ret=-1;
}
else
{
//����������� ��������
Ret=ReadRes(Key1, Key2);
}
}
else
{
Ret=0;
}
//ShowMessage("Ret="+IntToStr(Ret));
return Ret;
}
//---------------------------------------------------------------------------
String TForm1::GetFileDatabase(String NameDatabase)
{
String Ret="";
for(unsigned int i=0;i<Cl->VBases.size();i++)
{
 if(Cl->VBases[i].Name==NameDatabase)
 {
  Ret=Cl->VBases[i].FileName;
 }
}

return Ret;
}
//---------------------------------------------------------------------------
int TForm1::ReadRes(String Key1, String Key2)
{
 int N1=Propis(Key1);
 int N2=Propis(Key2);
 int Res=N1*10+N2;
 return Res;
}
//---------------------------------------------------------------------------
int TForm1::Propis(String S)
{
int Res;
if(S=="����")
{
  Res=0;
  return Res;
}
if(S=="����")
{
  Res=1;
  return Res;
}
if(S=="���")
{
  Res=2;
  return Res;
}
if(S=="���")
{
  Res=3;
  return Res;
}
if(S=="������")
{
  Res=4;
  return Res;
}
if(S=="����")
{
  Res=5;
  return Res;
}
if(S=="�����")
{
  Res=6;
  return Res;
}
if(S=="����")
{
  Res=7;
  return Res;
}
if(S=="������")
{
  Res=8;
  return Res;
}
if(S=="������")
{
  Res=9;
  return Res;
}

return Res;
}
//------------------------------------------------------------------------


void __fastcall TForm1::FormDestroy(TObject *Sender)
{


delete Cl;
}
//---------------------------------------------------------------------------

