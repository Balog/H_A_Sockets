//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Zastavka.h"
#include <comobj.hpp>
#include <fstream.h>
#include "MasterPointer.h"
#include "Main.h"
#include <FileCtrl.hpp>
#include "PassForm.h"
#include "Progress.h"
#include "MainForm.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TZast *Zast;

//---------------------------------------------------------------------------
__fastcall TZast::TZast(TComponent* Owner)
        : TForm(Owner)
{
/*
Saved=false;
Stop=false;
Start=false;
Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Admin.ini");

int AbsPathAspect=Ini->ReadInteger("Main","AbsPathAspect",1);
String AspectBase=Ini->ReadString("Main","AdminBase","");
String DiaryBase=Ini->ReadString("Main","DiaryBase","");
String AdminDatabase;
String DiaryDatabase;
if(AbsPathAspect==0)
{
AdminDatabase=Path+AspectBase;
DiaryDatabase=Path+DiaryBase;
}
else
{
AdminDatabase= AspectBase;
DiaryDatabase=DiaryBase;
}
MClient=new Client(ClientSocket, ActionManager1, this);
MClient->Database=new MDBConnector(ExtractFilePath(AdminDatabase), ExtractFileName(AdminDatabase), this);
MClient->Database->SetPatchBackUp("Archive");

int Days=Ini->ReadInteger("Main","StoreArchive",0);
MClient->Database->ClearArchive(Days);
MClient->Database->PackDB();
MClient->Database->Backup("Archive");

MClient->Diary=new MDBConnector(ExtractFilePath(DiaryDatabase), ExtractFileName(DiaryDatabase), this);
MClient->Diary->SetPatchBackUp("Archive");


MClient->Diary->ClearArchive(Days);
MClient->Diary->PackDB();
MClient->Diary->Backup("Archive");
*/
//Reg=false;
//Start=false;
MClient=new Client(ClientSocket, ActionManager1, this);
String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"NetAspects.ini");

int AbsPathMain=Ini->ReadInteger("Main","AbsPathMSpec",1);
String MSpecBase=Ini->ReadString("Main","MSpecBase","");
String MAspectBase=Ini->ReadString("Main","MAspBase","");
String UsrBase=Ini->ReadString("Main","UsrBase","");
if(AbsPathMain==0)
{
MainDatabase=Path+MSpecBase;
MainDatabase2=Path+MAspectBase;
MainDatabase3=Path+UsrBase;
}
else
{
MainDatabase= MSpecBase;
MainDatabase2=MAspectBase;
MainDatabase3=UsrBase;
}



ADOConn=new MDBConnector(ExtractFilePath(MainDatabase), ExtractFileName(MainDatabase), this);
ADOConn->SetPatchBackUp("Archive");
ADOConn->OnArchTimer=ArchTimer;

MClient->ADOConn=ADOConn;

ADOAspect=new MDBConnector(ExtractFilePath(MainDatabase2), ExtractFileName(MainDatabase2), this);
ADOAspect->SetPatchBackUp("Archive");
ADOAspect->OnArchTimer=ArchTimer;

MClient->ADOAspect=ADOAspect;

ADOUsrAspect=new MDBConnector(ExtractFilePath(MainDatabase3), ExtractFileName(MainDatabase3), this);
ADOUsrAspect->SetPatchBackUp("Archive");
ADOUsrAspect->OnArchTimer=ArchTimer;

MClient->ADOUsrAspect=ADOUsrAspect;
}
//---------------------------------------------------------------------------
void __fastcall TZast::Timer1Timer(TObject *Sender)
{

this->Hide();
//

this->BorderStyle=bsSingle;


Timer2->Enabled=false;
if(Start)
{
Timer1->Enabled=false;
//Pass->Show();
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ViewLogins");

 Zast->MClient->ReadTable("�������","Select Login, Code1, Code2, Role from Logins Where Role<>1", "Select Login, Code1, Code2, Role From TempLogins");


Timer1->Enabled=false;
}
else
{
Timer1->Interval=2000;
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::Image1Click(TObject *Sender)
{
this->Hide();
Timer1->Enabled=false;
this->BorderStyle=bsSingle;
Timer1->Enabled=false;
Timer1->Interval=100;
Timer1->Enabled=true;
}
//---------------------------------------------------------------------------
void __fastcall TZast::MyException(TObject *Sender,
                                    Exception *E)
{

}




//---------------------------------------------------

void __fastcall TZast::FormShow(TObject *Sender)
{
Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"NetAspects.ini");
Server=Ini->ReadString("Server","Server","localhost");
Port=Ini->ReadInteger("Server","Port",2000);

MClient->VDB.clear();
//Form1->CBDatabase->Clear();
//ServerName=Ini->ReadString("Server","ServerName","");
int Num=Ini->ReadInteger("Server","NumServerBase",0);
for(int i=0;i<Num;i++)
{
String S="Base"+IntToStr(i+1);
String ServerBase=Ini->ReadString(S,"ServerDatabase","");
String Name=Ini->ReadString(S,"Name","");

CDBItem DBI;
DBI.Name=Name;
DBI.ServerDB=ServerBase;
DBI.Num=i;
DBI.NumDatabase=Ini->ReadInteger(S,"NumDatabase",-1);

if(DBI.NumDatabase>=0)
{
//Form1->CBDatabase->Items->Add(Name);
MClient->VDB.push_back(DBI);
}


}
//Form1->CBDatabase->ItemIndex=0;
/*
MP<TADODataSet>Roles(this);
Roles->CommandText="Select * from Roles order by Num";
Roles->Connection=MClient->Database;
Roles->Active=true;
Form1->Role->Clear();
for(Roles->First();!Roles->Eof;Roles->Next())
{
Form1->Role->Items->Add(Roles->FieldByName("Name")->AsString);
}
Form1->Role->ItemIndex=0;
*/
HRGN MyRegion, MyRegion2, RoundReg, TempReg, MyRegion3, MyRegion4, MyRegion5, MyRegion6, MyRegion7;

MyRegion = CreateEllipticRgn(1, 1, 466, 466);

RoundReg = CreateEllipticRgn(74, 74, 395, 395);

TPoint R[9];
R[0]=Point(R1->Left,R1->Top);
R[1]=Point(R2->Left,R2->Top);
R[2]=Point(R3->Left,R3->Top);
R[3]=Point(R4->Left,R4->Top);
R[4]=Point(R5->Left,R5->Top);
R[5]=Point(R6->Left,R6->Top);
R[6]=Point(R7->Left,R7->Top);
R[7]=Point(R8->Left,R8->Top);
R[8]=Point(R9->Left,R9->Top);
R[9]=Point(R10->Left,R10->Top);

MyRegion3 = CreatePolygonRgn(R, 10, WINDING);

CombineRgn(MyRegion3, MyRegion3, RoundReg, RGN_AND);
CombineRgn(MyRegion, MyRegion, MyRegion3, RGN_XOR);

TPoint T[9];
T[0]=Point(T1->Left,T1->Top);
T[1]=Point(T2->Left,T2->Top);
T[2]=Point(T3->Left,T3->Top);
T[3]=Point(T4->Left,T4->Top);
T[4]=Point(T5->Left,T5->Top);
T[5]=Point(T6->Left,T6->Top);
T[6]=Point(T7->Left,T7->Top);
T[7]=Point(T8->Left,T8->Top);
T[8]=Point(T9->Left,T9->Top);
T[9]=Point(T10->Left,T10->Top);

MyRegion4 = CreatePolygonRgn(T, 10, WINDING);
CombineRgn(MyRegion4, MyRegion4, RoundReg, RGN_AND);
CombineRgn(MyRegion, MyRegion, MyRegion4, RGN_XOR);

TPoint Y[12];
Y[0]=Point(Y1->Left,Y1->Top);
Y[1]=Point(Y2->Left,Y2->Top);
Y[2]=Point(Y3->Left,Y3->Top);
Y[3]=Point(Y4->Left,Y4->Top);
Y[4]=Point(Y5->Left,Y5->Top);
Y[5]=Point(Y6->Left,Y6->Top);
Y[6]=Point(Y7->Left,Y7->Top);
Y[7]=Point(Y8->Left,Y8->Top);
Y[8]=Point(Y9->Left,Y9->Top);
Y[9]=Point(Y10->Left,Y10->Top);
Y[10]=Point(Y11->Left,Y11->Top);
Y[11]=Point(Y12->Left,Y12->Top);
Y[12]=Point(Y13->Left,Y13->Top);

MyRegion5 = CreatePolygonRgn(Y, 13, WINDING);
CombineRgn(MyRegion5, MyRegion5, RoundReg, RGN_AND);
CombineRgn(MyRegion, MyRegion, MyRegion5, RGN_XOR);

TPoint W[5];
W[0]=Point(W1->Left,W1->Top);
W[1]=Point(W2->Left,W2->Top);
W[2]=Point(W3->Left,W3->Top);
W[3]=Point(W4->Left,W4->Top);
W[4]=Point(W5->Left,W5->Top);
W[5]=Point(W6->Left,W6->Top);

MyRegion6 = CreatePolygonRgn(W, 6, WINDING);
CombineRgn(MyRegion6, MyRegion6, RoundReg, RGN_AND);
CombineRgn(MyRegion, MyRegion, MyRegion6, RGN_XOR);

TPoint U[13];
U[0]=Point(U1->Left,U1->Top);
U[1]=Point(U2->Left,U2->Top);
U[2]=Point(U3->Left,U3->Top);
U[3]=Point(U4->Left,U4->Top);
U[4]=Point(U5->Left,U5->Top);
U[5]=Point(U6->Left,U6->Top);
U[6]=Point(U7->Left,U7->Top);
U[7]=Point(U8->Left,U8->Top);
U[8]=Point(U9->Left,U9->Top);
U[9]=Point(U10->Left,U10->Top);
U[10]=Point(U11->Left,U11->Top);
U[11]=Point(U12->Left,U12->Top);
U[12]=Point(U13->Left,U13->Top);
U[13]=Point(U14->Left,U14->Top);

MyRegion7 = CreatePolygonRgn(U, 14, WINDING);
CombineRgn(MyRegion7, MyRegion7, RoundReg, RGN_AND);
CombineRgn(MyRegion, MyRegion, MyRegion7, RGN_XOR);


TPoint P[11];
P[0] = Point(StaticText1->Left,StaticText1->Top);
P[1] = Point(StaticText2->Left,StaticText2->Top);
P[2] = Point(StaticText3->Left,StaticText3->Top);
P[3] = Point(StaticText4->Left,StaticText4->Top);
P[4] = Point(StaticText5->Left,StaticText5->Top);
P[5] = Point(StaticText6->Left,StaticText6->Top);
P[6] = Point(StaticText7->Left,StaticText7->Top);
P[7] = Point(StaticText8->Left,StaticText8->Top);
P[8] = Point(StaticText8_1->Left,StaticText8_1->Top);
P[9] = Point(StaticText9->Left,StaticText9->Top);
P[10] = Point(StaticText10->Left,StaticText10->Top);
P[11] = Point(StaticText11->Left,StaticText11->Top);

MyRegion2 = CreatePolygonRgn(P, 12, ALTERNATE);
CombineRgn(MyRegion, MyRegion, MyRegion2, RGN_OR);

SetWindowRgn(Handle, MyRegion, true);

Timer2->Enabled=true;
}
//---------------------------------------------------------------------------


void TZast::CompactDB(TADOConnection * Conn, String B)
{
String CurDir=GetCurrentDir();
SetCurrentDir(ExtractFilePath(B));
String Pe=ExtractFileExt(B);
String P=B.SubString(1,B.Length()-Pe.Length());
AnsiString BP=P+".ldb";

if(FileExists(BP)==false)
{

Variant JE = Variant::CreateObject ("JRO.JetEngine");

WideString WSCurrConn;
WSCurrConn=Conn->ConnectionString;
try
 {

   String N1=ExtractFilePath(B)+"temp.mdb";
    DeleteFile(N1);
  JE.OleFunction ("CompactDatabase",  WSCurrConn,
   "Provider=Microsoft.Jet.OLEDB.4.0;Password="";Mode=ReadWrite;Data Source=temp.mdb");
  if(FileExists(N1))
   {
   String N=P+".dtb";
    CopyFile(N1.c_str(), N.c_str(), false);
    DeleteFile(N1);

   }
 }

catch ( ... )
{
SetCurrentDir(CurDir);
}
}
}
//-----------------------------------
void __fastcall TZast::ArchTimer(TObject *Sender, int N)
{

}
//----------------------------------
void __fastcall TZast::Timer2Timer(TObject *Sender)
{
Timer2->Enabled=false;
MClient->Connect(Server, Port);
Timer1->Enabled=true;
}
//---------------------------------------------------------------------------
void __fastcall TZast::FormDestroy(TObject *Sender)
{
delete MClient;
}
//---------------------------------------------------------------------------



void __fastcall TZast::ClientSocketRead(TObject *Sender,
      TCustomWinSocket *Socket)
{
//
String Mess="";
String Mess1="";
do
{
Sleep(100);
Mess1=Socket->ReceiveText();

Mess=Mess+Mess1;
}
while(Mess1.Length()!=0);

if (Mess.Length()!=0)
{
int N0=Mess.Pos(":");
int N=Mess.Pos(";");
int N1=Mess.Pos("|");

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
String Par1;
 int M=S.Pos("|");
 if(M>0)
 {
 Par=S.SubString(1,M-1);
 }
 else
 {
 Par=S;
 }
 int NLen=Par.Pos("#");
 int Len=StrToInt(Par.SubString(1,NLen-1));
 Par1=S.SubString(NLen+1,Len);
 S=S.SubString(M+1,S.Length());
 Parameters.push_back(Par1);
}
MClient->CommandExec(Comm, Parameters);
//}
}        
}
//---------------------------------------------------------------------------





void __fastcall TZast::PrepareConnectBaseExecute(TObject *Sender)
{

MClient->Act.NextCommand=4;
String Par="ConnectBase";
MClient->Act.ParamComm.clear();
String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"NetAspects.ini");
int MaxBase=MClient->VDB.size();

Trig T;
T.Var=0;
T.Max=MaxBase;
T.TrueAction="ConnectBase";
T.FalseAction="LoadLogins";
MClient->VTrigger.push_back(T);

MClient->ActTrigger(0);

}
//---------------------------------------------------------------------------

void __fastcall TZast::ConnectBaseExecute(TObject *Sender)
{

int TekNumBase=StrToInt(MClient->VTrigger[0].Var);
//MClient->Act.ParamComm.clear();
String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"NetAspects.ini");
int NumDatabase=MClient->VDB[TekNumBase].Num;
String NameDatabase=MClient->VDB[TekNumBase].Name;

/////////////////////////////////////////////////
//�������� ����������� ����������� ���
//�� Ini ����� � ���������� ����� ��������
//����� ������� 4
/////////////////////////////////////////////////
MClient->ConnectDatabase(NameDatabase, NumDatabase, true);

}
//---------------------------------------------------------------------------

void __fastcall TZast::LoadLoginsExecute(TObject *Sender)
{
Start=true;

}
//---------------------------------------------------------------------------

void __fastcall TZast::ViewLoginsExecute(TObject *Sender)
{
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->ADOAspect;
Tab->CommandText="Select Login, Code1, Code2 From TempLogins";
Tab->Active=true;

Pass->CbLogin->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
Pass->CbLogin->Items->Add(Tab->FieldByName("Login")->AsString);
}
Pass->CbLogin->ItemIndex=0;
Pass->EdPass->Text="";


Pass->ShowModal();
//Pass->ViewLogins();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ClientSocketError(TObject *Sender,
      TCustomWinSocket *Socket, TErrorEvent ErrorEvent, int &ErrorCode)
{
if(ErrorCode==10061 | ErrorCode==10060)
{
 this->Hide();
 ShowMessage("�� ������ ������!");
 ErrorCode=0;
 this->Close();
}
}
//---------------------------------------------------------------------------



void __fastcall TZast::BeginWorkExecute(TObject *Sender)
{
String Login=MClient->Act.ParamComm[0];
int Role=StrToInt(MClient->Act.ParamComm[1]);
if(Role==2)
{

}
else
{

}

switch (Role)
{
 case 2:
 {
  Documents->Show();
 break;
 }
 case 3:
 {
 Form1->Caption="������������: "+Login;
 Form1->Show();
 break;
 }
 case 4:
 {
 Form1->Caption="���������������� ������������: "+Login+"                ***������ �� ������ ���������***";
 Form1->Show();
 break;
 }
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::DeletePodrExecute(TObject *Sender)
{
int NumP=Documents->Podr->FieldByName("ServerNum")->AsInteger;

MP<TADODataSet>STab(this);
STab->Connection=ADOAspect;
STab->CommandText="SELECT �������.[����� �������], �������.������������� FROM ������� WHERE (((�������.�������������)="+IntToStr(NumP)+"));";
STab->Active=true;

int N=STab->RecordCount;


if(N==0)
{
Documents->Podr->Delete();
}
else
{
String S1;
int N1=N-((int)(N/10))*10;
switch(N1)
{
 case 1:
 {
 S1="�� ������� ������������ "+IntToStr(N)+" ������, ������������� ����� �������������.";
 break;
 }
 case 2: case 3: case 4:
 {
 S1="�� ������� ������������� "+IntToStr(N)+" �������, ������������� ����� �������������.";
 break;
 }
 default:
 {
 S1="�� ������� ������������� "+IntToStr(N)+" ��������, ������������� ����� �������������.";
 }
}

String S=S1+"\r��������� \"�������� ��������\" �������������� ���������� ��������� ������������� �� ��������";
Application->MessageBoxA(S.c_str(),"�������� �������������",MB_ICONEXCLAMATION);
}

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeMetodikaExecute(TObject *Sender)
{
Documents->Metod->Active=false;
Documents->Metod->Active=true;

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� �������� (��������)","");

ReadDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadMetodikaExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeMetodika");
Zast->MClient->ReadTable("Reference", "Select [�����], [��������] From ��������", "Select [�����], [��������] From ��������");

}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadDocExecute(TObject *Sender)
{
Str_RW S;
if(Documents->ReadWrite.size()!=0)
{
Documents->RW=Documents->ReadWrite.begin();
S=Documents->ReadWrite[0];
Documents->ReadWrite.erase(Documents->RW);
Prog->Label1->Caption=S.Text;
Prog->PB->Position=S.Num;
Prog->Repaint();
Sleep(1000);
MClient->StartAction(S.NameAction);
}
else
{
Prog->Close();
 ShowMessage("���������");
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadPodrazdExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergePodrazd");
Zast->MClient->ReadTable("�������", "Select �������������.[����� �������������], �������������.[�������� �������������] From ������������� order by [����� �������������]", "Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������] From Temp������������� order by [����� �������������]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergePodrazdExecute(TObject *Sender)
{
MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOAspect;
Podr->CommandText="Select * From �������������";
Podr->Active=true;

MP<TADODataSet>TempPodr(this);
TempPodr->Connection=Zast->ADOAspect;
TempPodr->CommandText="Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������] From Temp������������� order by [����� �������������]";
TempPodr->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="UPDATE ������������� SET �������������.Del = False;";
Comm->Execute();

for(Podr->First();!Podr->Eof;Podr->Next())
{
 int N=Podr->FieldByName("ServerNum")->Value;
 if(TempPodr->Locate("����� �������������",N,SO))
 {
  Podr->Edit();
  Podr->FieldByName("�������� �������������")->Value=TempPodr->FieldByName("�������� �������������")->Value;
  Podr->Post();

  TempPodr->Delete();
 }
 else
 {
  Podr->Edit();
  Podr->FieldByName("Del")->Value=true;
  Podr->Post();
 }
}
Comm->CommandText="DELETE * FROM ������������� WHERE (((�������������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ������������� ( ServerNum, [�������� �������������] ) SELECT Temp�������������.[����� �������������], Temp�������������.[�������� �������������] FROM Temp�������������;";
Comm->Execute();

Comm->CommandText="Delete * From Temp�������������";
Comm->Execute();

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ������������� (��������)","");

Documents->Podr->Active=false;
Documents->Podr->Active=true;

ReadDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadCritExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeCrit");
Zast->MClient->ReadTable("Reference", "Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[��������1], ����������.[��������], ����������.[��� �������], ����������.[���� �������], ����������.[����������� ����] From ���������� Order by [����� ����������];", "Select TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.[��������1], TempZn.[��������], TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] From TempZn Order by [����� ����������];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeCritExecute(TObject *Sender)
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
/*
Comm->CommandText="Delete * From TempZn";
Comm->Execute();
*/

MP<TADODataSet>TempZn(this);
TempZn->Connection=Zast->ADOConn;
TempZn->CommandText="Select TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.[��������1], TempZn.[��������], TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] From TempZn Order by [����� ����������]";
TempZn->Active=true;

Comm->CommandText="Delete * From ����������";
Comm->Execute();

Comm->CommandText="INSERT INTO ���������� ( [����� ����������], [������������ ����������], ��������1, ��������, [��� �������], [���� �������], [����������� ����] ) SELECT TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.��������1, TempZn.��������, TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] FROM TempZn;";
Comm->Execute();
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ��������� (��������)","");

Documents->ADODataSet1->Active=false;
Documents->ADODataSet1->Active=true;

ReadDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadSitExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeSit1");
Zast->MClient->ReadTable("Reference", "Select ��������.[����� ��������], ��������.[�������� ��������] From �������� order by [����� ��������];", "Select TempSit.[����� ��������], TempSit.[�������� ��������] From TempSit order by [����� ��������];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeSit1Execute(TObject *Sender)
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
Comm->CommandText="Delete * From ��������";
Comm->Execute();

Comm->CommandText="INSERT INTO �������� ( [����� ��������], [�������� ��������] ) SELECT TempSit.[����� ��������], TempSit.[�������� ��������] FROM TempSit;";
Comm->Execute();

//----
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * From TempSit";
Comm->Execute();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeSit2");
Zast->MClient->ReadTable("�������", "Select ��������.[����� ��������], ��������.[�������� ��������] From �������� Where �����=True order by [����� ��������];", "Select TempSit.[����� ��������], TempSit.[�������� ��������] From TempSit order by [����� ��������];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeSit2Execute(TObject *Sender)
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * From �������� where �����=true";
Comm->Execute();

Comm->CommandText="INSERT INTO �������� ( [����� ��������], [�������� ��������], ����� ) SELECT TempSit.[����� ��������], TempSit.[�������� ��������], True FROM TempSit;";
Comm->Execute();

Documents->Sit->Active=false;
Documents->Sit->Active=true;

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� �������� (��������)","");

ReadDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadVozd1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadVozd2");
Zast->MClient->ReadTable("Reference", "Select ����_3.[����� ����], ����_3.[��������], ����_3.[��������] From ����_3 Order by ��������, [����� ����];", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadVozd2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeVozd");
Zast->MClient->ReadTable("Reference", "Select �����_3.[����� �����], �����_3.[����� ��������], �����_3.[��������], �����_3.[�����] From �����_3 Order by [����� ��������], [����� �����];", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeVozdExecute(TObject *Sender)
{
Documents->MergeNode("����_3");
Documents->MergeBranch("�����_3");

Documents->LoadTab1();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeVozd2");
Zast->MClient->ReadTable("�������", "Select �����������.[����� �����������], �����������.[������������ �����������], �����������.[�����] From ����������� order by [����� �����������];", "Select TempVozd.[����� �����������], TempVozd.[������������ �����������], TempVozd.[�����] From TempVozd order by [����� �����������];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeVozd2Execute(TObject *Sender)
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="UPDATE ����������� SET �����������.Del = False;";
Comm->Execute();

MP<TADODataSet>TempVozd(this);
TempVozd->Connection=Zast->ADOAspect;
TempVozd->CommandText="Select TempVozd.[����� �����������], TempVozd.[������������ �����������], TempVozd.[�����] From TempVozd order by [����� �����������]";
TempVozd->Active=true;

MP<TADODataSet>Vozd(this);
Vozd->Connection=Zast->ADOAspect;
Vozd->CommandText="Select * From �����������";
Vozd->Active=true;

 for(Vozd->First();!Vozd->Eof;Vozd->Next())
 {
  int N=Vozd->FieldByName("����� �����������")->Value;

  if(TempVozd->Locate("����� �����������", N, SO))
  {
   Vozd->Edit();
   Vozd->FieldByName("������������ �����������")->Value=TempVozd->FieldByName("������������ �����������")->Value;
   Vozd->FieldByName("�����")->Value=TempVozd->FieldByName("�����")->Value;
   Vozd->Post();

   TempVozd->Delete();
  }
  else
  {
   Vozd->Edit();
   Vozd->FieldByName("Del")->Value=true;
   Vozd->Post();
  }
 }
Comm->CommandText="DELETE �����������.Del FROM ����������� WHERE (((�����������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ����������� ( [����� �����������], [������������ �����������], ����� ) SELECT TempVozd.[����� �����������], TempVozd.[������������ �����������], TempVozd.����� FROM TempVozd;";
Comm->Execute();

Comm->CommandText="DELETE TempVozd.* FROM TempVozd;";
Comm->Execute();

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ����������� (��������)","");

ReadDoc->Execute();
}
//---------------------------------------------------------------------------


void __fastcall TZast::ReadMeropr1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadMeropr2");
Zast->MClient->ReadTable("Reference", "Select ����_4.[����� ����], ����_4.[��������], ����_4.[��������] From ����_4 Order by ��������, [����� ����];", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");

}
//---------------------------------------------------------------------------
void __fastcall TZast::ReadMeropr2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeMeropr");
Zast->MClient->ReadTable("Reference", "Select �����_4.[����� �����], �����_4.[����� ��������], �����_4.[��������], �����_4.[�����] From �����_4 Order by [����� ��������], [����� �����];", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeMeroprExecute(TObject *Sender)
{
Documents->MergeNode("����_4");
Documents->MergeBranch("�����_4");

Documents->LoadTab2();

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ����������� (��������)","");

ReadDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadTerr1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadTerr2");
Zast->MClient->ReadTable("Reference", "Select ����_5.[����� ����], ����_5.[��������], ����_5.[��������] From ����_5 Order by ��������, [����� ����];", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");


}
//---------------------------------------------------------------------------
void __fastcall TZast::ReadTerr2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeTerr1");
Zast->MClient->ReadTable("Reference", "Select �����_5.[����� �����], �����_5.[����� ��������], �����_5.[��������], �����_5.[�����] From �����_5 Order by [����� ��������], [����� �����];", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");


}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeTerr1Execute(TObject *Sender)
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="UPDATE ���������� SET ����������.Del = False;";
Comm->Execute();

MP<TADODataSet>TempTerr(this);
TempTerr->Connection=Zast->ADOAspect;
TempTerr->CommandText="Select TempTerr.[����� ����������], TempTerr.[������������ ����������], TempTerr.[�����] From TempTerr order by [����� ����������]";
TempTerr->Active=true;

MP<TADODataSet>Terr(this);
Terr->Connection=Zast->ADOAspect;
Terr->CommandText="Select * From ����������";
Terr->Active=true;

 for(Terr->First();!Terr->Eof;Terr->Next())
 {
  int N=Terr->FieldByName("����� ����������")->Value;

  if(TempTerr->Locate("����� ����������", N, SO))
  {
   Terr->Edit();
   Terr->FieldByName("������������ ����������")->Value=TempTerr->FieldByName("������������ ����������")->Value;
   Terr->FieldByName("�����")->Value=TempTerr->FieldByName("�����")->Value;
   Terr->Post();

   TempTerr->Delete();
  }
  else
  {
   Terr->Edit();
   Terr->FieldByName("Del")->Value=true;
   Terr->Post();
  }
 }
Comm->CommandText="DELETE ����������.Del FROM ���������� WHERE (((����������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ���������� ( [����� ����������], [������������ ����������], ����� ) SELECT TempTerr.[����� ����������], TempTerr.[������������ ����������], TempTerr.����� FROM TempTerr;";
Comm->Execute();

Comm->CommandText="DELETE TempTerr.* FROM TempTerr;";
Comm->Execute();

Documents->MergeNode("����_5");
Documents->MergeBranch("�����_5");
Documents->LoadTab3();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeTerr2");
Zast->MClient->ReadTable("�������", "Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[�����] From ���������� order by [����� ����������];", "Select TempTerr.[����� ����������], TempTerr.[������������ ����������], TempTerr.[�����] From TempTerr order by [����� ����������];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeTerr2Execute(TObject *Sender)
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="UPDATE ���������� SET ����������.Del = False;";
Comm->Execute();

MP<TADODataSet>TempTerr(this);
TempTerr->Connection=Zast->ADOAspect;
TempTerr->CommandText="Select TempTerr.[����� ����������], TempTerr.[������������ ����������], TempTerr.[�����] From TempTerr order by [����� ����������]";
TempTerr->Active=true;

MP<TADODataSet>Terr(this);
Terr->Connection=Zast->ADOAspect;
Terr->CommandText="Select * From ����������";
Terr->Active=true;

 for(Terr->First();!Terr->Eof;Terr->Next())
 {
  int N=Terr->FieldByName("����� ����������")->Value;

  if(TempTerr->Locate("����� ����������", N, SO))
  {
   Terr->Edit();
   Terr->FieldByName("������������ ����������")->Value=TempTerr->FieldByName("������������ ����������")->Value;
   Terr->FieldByName("�����")->Value=TempTerr->FieldByName("�����")->Value;
   Terr->Post();

   TempTerr->Delete();
  }
  else
  {
   Terr->Edit();
   Terr->FieldByName("Del")->Value=true;
   Terr->Post();
  }
 }
Comm->CommandText="DELETE ����������.Del FROM ���������� WHERE (((����������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ���������� ( [����� ����������], [������������ ����������], ����� ) SELECT TempTerr.[����� ����������], TempTerr.[������������ ����������], TempTerr.����� FROM TempTerr;";
Comm->Execute();

Comm->CommandText="DELETE TempTerr.* FROM TempTerr;";
Comm->Execute();

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ���������� (��������)","");

ReadDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadDeyat1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadDeyat2");
Zast->MClient->ReadTable("Reference", "Select ����_6.[����� ����], ����_6.[��������], ����_6.[��������] From ����_6 Order by ��������, [����� ����];", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");


}
//---------------------------------------------------------------------------
void __fastcall TZast::ReadDeyat2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeDeyat1");
Zast->MClient->ReadTable("Reference", "Select �����_6.[����� �����], �����_6.[����� ��������], �����_6.[��������], �����_6.[�����] From �����_6 Order by [����� ��������], [����� �����];", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeDeyat1Execute(TObject *Sender)
{
Documents->MergeNode("����_6");
Documents->MergeBranch("�����_6");
Documents->LoadTab4();



 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeDeyat2");
Zast->MClient->ReadTable("�������", "Select ������������.[����� ������������], ������������.[������������ ������������], ������������.[�����] From ������������ order by [����� ������������];", "Select TempDeyat.[����� ������������], TempDeyat.[������������ ������������], TempDeyat.[�����] From TempDeyat order by [����� ������������];");


}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeDeyat2Execute(TObject *Sender)
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="UPDATE ������������ SET ������������.Del = False;";
Comm->Execute();

MP<TADODataSet>TempDeyat(this);
TempDeyat->Connection=Zast->ADOAspect;
TempDeyat->CommandText="Select TempDeyat.[����� ������������], TempDeyat.[������������ ������������], TempDeyat.[�����] From TempDeyat order by [����� ������������]";
TempDeyat->Active=true;

MP<TADODataSet>Deyat(this);
Deyat->Connection=Zast->ADOAspect;
Deyat->CommandText="Select * From ������������";
Deyat->Active=true;

 for(Deyat->First();!Deyat->Eof;Deyat->Next())
 {
  int N=Deyat->FieldByName("����� ������������")->Value;

  if(TempDeyat->Locate("����� ������������", N, SO))
  {
   Deyat->Edit();
   Deyat->FieldByName("������������ ������������")->Value=TempDeyat->FieldByName("������������ ������������")->Value;
   Deyat->FieldByName("�����")->Value=TempDeyat->FieldByName("�����")->Value;
   Deyat->Post();

   TempDeyat->Delete();
  }
  else
  {
   Deyat->Edit();
   Deyat->FieldByName("Del")->Value=true;
   Deyat->Post();
  }
 }
Comm->CommandText="DELETE ������������.Del FROM ������������ WHERE (((������������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ������������ ( [����� ������������], [������������ ������������], ����� ) SELECT TempDeyat.[����� ������������], TempDeyat.[������������ ������������], TempDeyat.����� FROM TempDeyat;";
Comm->Execute();

Comm->CommandText="DELETE TempDeyat.* FROM TempDeyat;";
Comm->Execute();

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ������������� (��������)","");

ReadDoc->Execute();

}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadAspect1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadAspect2");
Zast->MClient->ReadTable("Reference", "Select ����_7.[����� ����], ����_7.[��������], ����_7.[��������] From ����_7 Order by ��������, [����� ����];", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");


}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadAspect2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeAspect1");
Zast->MClient->ReadTable("Reference", "Select �����_7.[����� �����], �����_7.[����� ��������], �����_7.[��������], �����_7.[�����] From �����_7 Order by [����� ��������], [����� �����];", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");


}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeAspect1Execute(TObject *Sender)
{
Documents->MergeNode("����_7");
Documents->MergeBranch("�����_7");
Documents->LoadTab5();



 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeAspect2");
Zast->MClient->ReadTable("�������", "Select ������.[����� �������], ������.[������������ �������], ������.[�����] From ������ order by [����� �������];", "Select TempAspect.[����� �������], TempAspect.[������������ �������], TempAspect.[�����] From TempAspect order by [����� �������];");


}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeAspect2Execute(TObject *Sender)
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="UPDATE ������ SET ������.Del = False;";
Comm->Execute();

MP<TADODataSet>TempAspect(this);
TempAspect->Connection=Zast->ADOAspect;
TempAspect->CommandText="Select TempAspect.[����� �������], TempAspect.[������������ �������], TempAspect.[�����] From TempAspect order by [����� �������]";
TempAspect->Active=true;

MP<TADODataSet>Aspect(this);
Aspect->Connection=Zast->ADOAspect;
Aspect->CommandText="Select * From ������";
Aspect->Active=true;

 for(Aspect->First();!Aspect->Eof;Aspect->Next())
 {
  int N=Aspect->FieldByName("����� �������")->Value;

  if(TempAspect->Locate("����� �������", N, SO))
  {
   Aspect->Edit();
   Aspect->FieldByName("������������ �������")->Value=TempAspect->FieldByName("������������ �������")->Value;
   Aspect->FieldByName("�����")->Value=TempAspect->FieldByName("�����")->Value;
   Aspect->Post();

   TempAspect->Delete();
  }
  else
  {
   Aspect->Edit();
   Aspect->FieldByName("Del")->Value=true;
   Aspect->Post();
  }
 }
Comm->CommandText="DELETE ������.Del FROM ������ WHERE (((������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ������ ( [����� �������], [������������ �������], ����� ) SELECT TempAspect.[����� �������], TempAspect.[������������ �������], TempAspect.����� FROM TempAspect;";
Comm->Execute();

Comm->CommandText="DELETE TempDeyat.* FROM TempDeyat;";
Comm->Execute();

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ������������� �������� (��������)","");

ReadDoc->Execute();
}
//---------------------------------------------------------------------------
