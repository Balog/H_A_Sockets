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
#include "Rep1.h"
#include "Svod.h"
#include "FMoveAsp.h"
#include "InputFiltr.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TZast *Zast;

  HINSTANCE hDll;
  //���� �������� ������� �������
  DWORD __stdcall (*BlockInput)(bool Status);
  DWORD Result;
  TCursor Save_Cursor;
//---------------------------------------------------------------------------
__fastcall TZast::TZast(TComponent* Owner)
        : TForm(Owner)
{
Result=false;
Save_Cursor = Screen->Cursor;

MClient=new Client(ClientSocket, ActionManager1, this);
String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Hazards.ini");

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
ADOConn->PackDB();
ADOConn->Backup("Archive");

MClient->ADOConn=ADOConn;

ADOAspect=new MDBConnector(ExtractFilePath(MainDatabase2), ExtractFileName(MainDatabase2), this);
ADOAspect->SetPatchBackUp("Archive");
ADOAspect->OnArchTimer=ArchTimer;
ADOAspect->PackDB();
ADOAspect->Backup("Archive");

MClient->ADOAspect=ADOAspect;

ADOUsrAspect=new MDBConnector(ExtractFilePath(MainDatabase3), ExtractFileName(MainDatabase3), this);
ADOUsrAspect->SetPatchBackUp("Archive");
ADOUsrAspect->OnArchTimer=ArchTimer;
ADOUsrAspect->PackDB();
ADOUsrAspect->Backup("Archive");

MClient->ADOUsrAspect=ADOUsrAspect;
}
//---------------------------------------------------------------------------
void __fastcall TZast::Timer1Timer(TObject *Sender)
{

this->Hide();

this->BorderStyle=bsSingle;


Timer2->Enabled=false;
if(Start)
{
Timer1->Enabled=false;
//Pass->Show();

MClient->BlockServer("PreLoadLogins");

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
MP<TIniFile>Ini(Path+"Hazards.ini");
Server=Ini->ReadString("Server","Server","localhost");
Port=Ini->ReadInteger("Server","Port",2000);

MClient->VDB.clear();

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

MClient->VDB.push_back(DBI);
}


}
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
try
{
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
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------
void __fastcall TZast::PrepareConnectBaseExecute(TObject *Sender)
{

MClient->Act.NextCommand=4;
String Par="ConnectBase";
MClient->Act.ParamComm.clear();
String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Hazards.ini");
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

String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Hazards.ini");
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

MClient->UnBlockServer("");
Pass->ShowModal();

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
switch (Role)
{
 case 2:
 {
//���������� �������� ������, �������������, ������������� �������
//� ����� ����� � ��������� �����
//� ������ ����� ���������� ����� ����������

//������ �� �������� �������
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadPodr");
 Zast->MClient->ReadTable("���������", "Select Logins.Num, Logins.Login, Logins.Role From Logins Order by Num;", "���������", "Select TempLogins.Num, TempLogins.Login, TempLogins.Role From TempLogins Order by Num;");

 break;
 }
 case 3:
 {
  //Form1->Login=Login;
   Form1->Caption="������������: "+Form1->Login;
  Form1->Demo=false;
  Form1->Role=3;

 //��� ���� ���� ������� ������� ���������� ������
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadPodrUSR");
 Zast->MClient->ReadTable("���������", "Select Logins.Num, Logins.Login, Logins.Role From Logins Order by Num;", "���������_�", "Select TempLogins.Num, TempLogins.Login, TempLogins.Role From TempLogins Order by Num;");




 break;
 }
 case 4:
 {
 //��� ���� ���� ������� ������� ���������� ������

 //Form1->Login=Login;
 Form1->Role=4;
 Form1->Caption="���������������� ������������: "+Form1->Login+"                ***������ �� ������ ���������***";
 Form1->Demo=true;
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadPodrUSR");
 Zast->MClient->ReadTable("���������", "Select Logins.Num, Logins.Login, Logins.Role From Logins Order by Num;", "���������_�", "Select TempLogins.Num, TempLogins.Login, TempLogins.Role From TempLogins Order by Num;");


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
//Documents->Podr->Delete();
 Zast->MClient->Act.WaitCommand=16;
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("DeletePodr2");
 String Database="���������";
String NumPodr=IntToStr(NumP);
ClientSocket->Socket->SendText("Command:16;2|"+IntToStr(Database.Length())+"#"+Database+"|"+IntToStr(NumPodr.Length())+"#"+NumPodr+"|");
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
try
{
Documents->Metod->Active=false;
Documents->Metod->Active=true;

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� �������� (��������)","");
//Sleep(1000);
Zast->MClient->UnBlockServer("ReadWriteDoc");
}
catch(...)
{
Zast->BlockMK(false);
}
//ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadMetodikaExecute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeMetodika");
Zast->MClient->ReadTable("HReference", "Select [�����], [��������] From ��������", "HReference", "Select [�����], [��������] From ��������");
}
catch(...)
{
Zast->BlockMK(false);
}

}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadWriteDocExecute(TObject *Sender)
{
Str_RW S;
if(Documents->ReadWrite.size()!=0)
{
try
{


Documents->RW=Documents->ReadWrite.begin();
S=Documents->ReadWrite[0];
Documents->ReadWrite.erase(Documents->RW);
//Prog->Visible=true;
Prog->Label1->Caption=S.Text;
Prog->PB->Position=S.Num;
Prog->Repaint();
Sleep(1000);
Zast->MClient->BlockServer(S.NameAction);
}
catch(...)
{
Zast->BlockMK(false);
}
//MClient->StartAction(S.NameAction);
}
else
{
Prog->Close();
Zast->BlockMK(false);
if(Prog->SignComplete)
{
 ShowMessage("���������");
}
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadPodrazdExecute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergePodrazd");
 String DBTo;
 if(Role==2)
 {
  DBTo="���������";
 }
 else
 {
  DBTo="���������_�";
 }
Zast->MClient->ReadTable("���������", "Select �������������.[����� �������������], �������������.[�������� �������������] From ������������� order by [����� �������������]", DBTo, "Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������] From Temp������������� order by [����� �������������]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergePodrazdExecute(TObject *Sender)
{
try
{
Filter->SetDefFiltr();
MDBConnector* DB;
 if(Role==2)
 {
  DB=ADOAspect;
 }
 else
 {
  DB=ADOUsrAspect;
 }
MP<TADODataSet>Podr(this);
Podr->Connection=DB;
Podr->CommandText="Select * From �������������";
Podr->Active=true;

MP<TADODataSet>TempPodr(this);
TempPodr->Connection=DB;
TempPodr->CommandText="Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������] From Temp������������� order by [����� �������������]";
TempPodr->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
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

if(Role==2)
{
Documents->Podr->Active=false;
Documents->Podr->Active=true;

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ������������� (��������)","");
}
else
{
Form1->Initialize();

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ������������� (������������)","");
}
Zast->MClient->UnBlockServer("ReadWriteDoc");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadCritExecute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeCrit");
Zast->MClient->ReadTable("HReference", "Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[��������1], ����������.[��������], ����������.[��� �������], ����������.[���� �������], ����������.[����������� ����] From ���������� Order by [����� ����������];", "HReference", "Select TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.[��������1], TempZn.[��������], TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] From TempZn Order by [����� ����������];");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeCritExecute(TObject *Sender)
{
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;

MP<TADODataSet>TempZn(this);
TempZn->Connection=Zast->ADOConn;
TempZn->CommandText="Select TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.[��������1], TempZn.[��������], TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] From TempZn Order by [����� ����������]";
TempZn->Active=true;

Comm->CommandText="Delete * From ����������";
Comm->Execute();

Comm->CommandText="INSERT INTO ���������� ( [����� ����������], [������������ ����������], ��������1, ��������, [��� �������], [���� �������], [����������� ����] ) SELECT TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.��������1, TempZn.��������, TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] FROM TempZn;";
Comm->Execute();
if(Role==2)
{
Documents->ADODataSet1->Active=false;
Documents->ADODataSet1->Active=true;

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ��������� (��������)","");
}
else
{
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From ����������";
Comm->Execute();

MP<TADODataSet>From(this);
From->Connection=Zast->ADOConn;
From->CommandText="Select * From ���������� Order by [����� ����������]";
From->Active=true;

MP<TADODataSet>To(this);
To->Connection=Zast->ADOUsrAspect;
To->CommandText="Select * From ����������";
To->Active=true;

for(From->First();!From->Eof;From->Next())
{
 To->Append();
 To->FieldByName("����� ����������")->Value=From->FieldByName("����� ����������")->Value;
 To->FieldByName("������������ ����������")->Value=From->FieldByName("������������ ����������")->Value;
 To->FieldByName("��������1")->Value=From->FieldByName("��������1")->Value;
 To->FieldByName("��������")->Value=From->FieldByName("��������")->Value;
 To->FieldByName("��� �������")->Value=From->FieldByName("��� �������")->Value;
 To->FieldByName("���� �������")->Value=From->FieldByName("���� �������")->Value;
 To->FieldByName("����������� ����")->Value=From->FieldByName("����������� ����")->Value;
 To->Post();
}

Form1->Initialize();

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ��������� (������������)","");
}



Zast->MClient->UnBlockServer("ReadWriteDoc");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadSitExecute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeSit1");
Zast->MClient->ReadTable("HReference", "Select ��������.[����� ��������], ��������.[�������� ��������] From �������� order by [����� ��������];", "HReference", "Select TempSit.[����� ��������], TempSit.[�������� ��������] From TempSit order by [����� ��������];");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeSit1Execute(TObject *Sender)
{
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
Comm->CommandText="Delete * From ��������";
Comm->Execute();

Comm->CommandText="INSERT INTO �������� ( [����� ��������], [�������� ��������] ) SELECT TempSit.[����� ��������], TempSit.[�������� ��������] FROM TempSit;";
Comm->Execute();

//----
MDBConnector* DB;
String DBName;
 if(Role==2)
 {
  DB=ADOAspect;
  DBName="���������";
 }
 else
 {
  DB=ADOUsrAspect;
  DBName="���������_�";
 }

Comm->Connection=DB;
Comm->CommandText="Delete * From TempSit";
Comm->Execute();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeSit2");
Zast->MClient->ReadTable("���������", "Select ��������.[����� ��������], ��������.[�������� ��������] From �������� Where �����=True order by [����� ��������];", DBName, "Select TempSit.[����� ��������], TempSit.[�������� ��������] From TempSit order by [����� ��������];");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeSit2Execute(TObject *Sender)
{
try
{
MDBConnector* DB;
 if(Role==2)
 {
  DB=ADOAspect;
 }
 else
 {
  DB=ADOUsrAspect;
 }
MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="UPDATE �������� SET ��������.Del = True Where �����=true;";
Comm->Execute();

MP<TADODataSet>Temp(this);
Temp->Connection=DB;
Temp->CommandText="Select * from TempSit";
Temp->Active=true;

MP<TADODataSet>Sit(this);
Sit->Connection=DB;
Sit->CommandText="Select * From �������� Where �����=true";
Sit->Active=true;

for(Sit->First();!Sit->Eof;Sit->Next())
{
int Num=Sit->FieldByName("����� ��������")->Value;
 if(Temp->Locate("����� ��������", Num, SO))
 {
  Sit->Edit();
  Sit->FieldByName("�������� ��������")->Value=Temp->FieldByName("�������� ��������")->Value;
  Sit->FieldByName("Del")->Value=false;
  Sit->Post();

  Temp->Delete();
 }
 else
 {
  Sit->Edit();
  Sit->FieldByName("Del")->Value=true;
  Sit->Post();

Comm->CommandText="UPDATE ������� SET �������.�������� = 0 WHERE (((�������.��������)="+IntToStr(Num)+"));";
Comm->Execute();
 }
}
Comm->CommandText="DELETE * From �������� Where Del=true AND �����=true";
Comm->Execute();

Comm->CommandText="INSERT INTO �������� ( [����� ��������], [�������� ��������], �����, Del ) SELECT TempSit.[����� ��������], TempSit.[�������� ��������], True AS ���������1, False AS ���������2 FROM TempSit;";
Comm->Execute();

Comm->CommandText="DELETE * From TempSit";
Comm->Execute();

 if(Role==2)
 {

Documents->Sit->Active=false;
Documents->Sit->Active=true;

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� �������� (��������)","");
 }
 else
 {


Form1->Initialize();

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� �������� (������������)","");
 }
Zast->MClient->UnBlockServer("ReadWriteDoc");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadVozd1Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadVozd2");
Zast->MClient->ReadTable("HReference", "Select ����_3.[����� ����], ����_3.[��������], ����_3.[��������] From ����_3 Order by ��������, [����� ����];", "HReference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadVozd2Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeVozd");
Zast->MClient->ReadTable("HReference", "Select �����_3.[����� �����], �����_3.[����� ��������], �����_3.[��������], �����_3.[�����] From �����_3 Order by [����� ��������], [����� �����];", "HReference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeVozdExecute(TObject *Sender)
{
try
{
Documents->MergeNode("����_3");
Documents->MergeBranch("�����_3");

Documents->LoadTab1();

String DBName;
if(Role==2)
{
 DBName="���������";
}
else
{
 DBName="���������_�";
}

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeVozd2");
Zast->MClient->ReadTable("���������", "Select �����������.[����� �����������], �����������.[������������ �����������], �����������.[�����] From ����������� order by [����� �����������];", DBName, "Select TempVozd.[����� �����������], TempVozd.[������������ �����������], TempVozd.[�����] From TempVozd order by [����� �����������];");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeVozd2Execute(TObject *Sender)
{
try
{
MDBConnector* DB;
if(Role==2)
{
 DB=ADOAspect;
}
else
{
 DB=ADOUsrAspect;
}

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="UPDATE ����������� SET �����������.Del = False;";
Comm->Execute();

MP<TADODataSet>TempVozd(this);
TempVozd->Connection=DB;
TempVozd->CommandText="Select TempVozd.[����� �����������], TempVozd.[������������ �����������], TempVozd.[�����] From TempVozd order by [����� �����������]";
TempVozd->Active=true;

MP<TADODataSet>Vozd(this);
Vozd->Connection=DB;
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
  Comm->CommandText="UPDATE ������� SET �������.����������� = 0 WHERE (((�������.�����������)="+IntToStr(N)+"));";
  Comm->Execute();

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

if(Role==2)
{
Documents->Sit->Active=false;
Documents->Sit->Active=true;
Documents->LoadTab1();
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ����������� (��������)","");
}
else
{
Form1->Initialize();
Documents->LoadTab1();
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ����������� (������������)","");
}

Zast->MClient->UnBlockServer("ReadWriteDoc");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------


void __fastcall TZast::ReadMeropr1Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadMeropr2");
Zast->MClient->ReadTable("HReference", "Select ����_4.[����� ����], ����_4.[��������], ����_4.[��������] From ����_4 Order by ��������, [����� ����];", "HReference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------
void __fastcall TZast::ReadMeropr2Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeMeropr");
Zast->MClient->ReadTable("HReference", "Select �����_4.[����� �����], �����_4.[����� ��������], �����_4.[��������], �����_4.[�����] From �����_4 Order by [����� ��������], [����� �����];", "HReference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeMeroprExecute(TObject *Sender)
{
try
{
//�������� ������� ������ ������ � ��������������� ������� �������� � ���������������� ��������
Documents->MergeNode("����_4");
Documents->MergeBranch("�����_4");
if(Role==2)
{
Documents->LoadTab2();

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ����������� (��������)","");
}
else
{
Form1->Initialize();
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ����������� (������������)","");
}

Zast->MClient->UnBlockServer("ReadWriteDoc");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadTerr1Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadTerr2");
Zast->MClient->ReadTable("HReference", "Select ����_5.[����� ����], ����_5.[��������], ����_5.[��������] From ����_5 Order by ��������, [����� ����];", "HReference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");
}
catch(...)
{
 Zast->BlockMK(false);
}

}
//---------------------------------------------------------------------------
void __fastcall TZast::ReadTerr2Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeTerr1");
Zast->MClient->ReadTable("HReference", "Select �����_5.[����� �����], �����_5.[����� ��������], �����_5.[��������], �����_5.[�����] From �����_5 Order by [����� ��������], [����� �����];", "HReference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");
}
catch(...)
{
 Zast->BlockMK(false);
}

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeTerr1Execute(TObject *Sender)
{
try
{
MDBConnector* DB;
String DBName;
if(Role==2)
{
 DB=ADOAspect;
 DBName="���������";
}
else
{
 DB=ADOUsrAspect;
 DBName="���������_�";
}
MP<TADODataSet>Ref(this);
Ref->Connection=Zast->ADOConn;
Ref->CommandText="Select * From �����_5";
Ref->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="UPDATE ���������� SET ����������.Del = False;";
Comm->Execute();

Comm->CommandText="Delete * From TempTerr";
Comm->Execute();

MP<TADODataSet>TempTerr(this);
TempTerr->Connection=DB;
TempTerr->CommandText="Select TempTerr.[����� ����������], TempTerr.[������������ ����������], TempTerr.[�����] From TempTerr order by [����� ����������]";
TempTerr->Active=true;

for(Ref->First();!Ref->Eof;Ref->Next())
{
 TempTerr->Append();
 TempTerr->FieldByName("����� ����������")->Value=Ref->FieldByName("����� �����")->Value;
 TempTerr->FieldByName("������������ ����������")->Value=Ref->FieldByName("��������")->Value;
 TempTerr->FieldByName("�����")->Value=Ref->FieldByName("�����")->Value;
 TempTerr->Post();
}

MP<TADODataSet>Terr(this);
Terr->Connection=DB;
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
  Comm->CommandText="UPDATE ������� SET �������.[��� ����������] = 0 WHERE (((�������.[��� ����������])="+IntToStr(N)+"));";
  Comm->Execute();

   Terr->Edit();
   Terr->FieldByName("Del")->Value=true;
   Terr->Post();
  }
 }
Comm->CommandText="UPDATE ���������� INNER JOIN ������� ON ����������.[����� ����������] = �������.[��� ����������] SET �������.[��� ����������] = 0 WHERE (((����������.Del)=True));";
Comm->Execute();

Comm->CommandText="DELETE ����������.Del FROM ���������� WHERE (((����������.Del)=True) AND �����=true);";
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
Zast->MClient->ReadTable("���������", "Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[�����] From ���������� order by [����� ����������];", DBName, "Select TempTerr.[����� ����������], TempTerr.[������������ ����������], TempTerr.[�����] From TempTerr order by [����� ����������];");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeTerr2Execute(TObject *Sender)
{
try
{
MDBConnector* DB;
String DBName;
if(Role==2)
{
 DB=ADOAspect;
 DBName="���������";
}
else
{
 DB=ADOUsrAspect;
 DBName="���������_�";
}

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="UPDATE ���������� SET ����������.Del = False;";
Comm->Execute();

MP<TADODataSet>TempTerr(this);
TempTerr->Connection=DB;
TempTerr->CommandText="Select TempTerr.[����� ����������], TempTerr.[������������ ����������], TempTerr.[�����] From TempTerr order by [����� ����������]";
TempTerr->Active=true;

MP<TADODataSet>Terr(this);
Terr->Connection=DB;
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

if(Role==2)
{
//Form1->Initialize();
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ���������� (��������)","");
}
else
{
//
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ���������� (������������)","");
}

Zast->MClient->UnBlockServer("ReadWriteDoc");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadDeyat1Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadDeyat2");
Zast->MClient->ReadTable("HReference", "Select ����_6.[����� ����], ����_6.[��������], ����_6.[��������] From ����_6 Order by ��������, [����� ����];", "HReference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");
}
catch(...)
{
 Zast->BlockMK(false);
}

}
//---------------------------------------------------------------------------
void __fastcall TZast::ReadDeyat2Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeDeyat1");
Zast->MClient->ReadTable("HReference", "Select �����_6.[����� �����], �����_6.[����� ��������], �����_6.[��������], �����_6.[�����] From �����_6 Order by [����� ��������], [����� �����];", "HReference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeDeyat1Execute(TObject *Sender)
{
try
{
String DBName;
if(Role==2)
{
 DBName="���������";
}
else
{
 DBName="���������_�";
}

Documents->MergeNode("����_6");
Documents->MergeBranch("�����_6");

if(Role==2)
{
Documents->LoadTab4();
}


 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeDeyat2");
Zast->MClient->ReadTable("���������", "Select ������������.[����� ������������], ������������.[������������ ������������], ������������.[�����] From ������������ order by [����� ������������];", DBName, "Select TempDeyat.[����� ������������], TempDeyat.[������������ ������������], TempDeyat.[�����] From TempDeyat order by [����� ������������];");

}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeDeyat2Execute(TObject *Sender)
{

try
{
MDBConnector* DB;
String DBName;
if(Role==2)
{
 DB=ADOAspect;
 DBName="���������";
}
else
{
 DB=ADOUsrAspect;
 DBName="���������_�";
}
MP<TADODataSet>Ref(this);
Ref->Connection=Zast->ADOConn;
Ref->CommandText="Select * From �����_6";
Ref->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="UPDATE ������������ SET ������������.Del = False;";
Comm->Execute();

//Comm->CommandText="Delete * From TempDeyat";
//Comm->Execute();

MP<TADODataSet>TempDeyat(this);
TempDeyat->Connection=DB;
TempDeyat->CommandText="Select TempDeyat.[����� ������������], TempDeyat.[������������ ������������], TempDeyat.[�����] From TempDeyat order by [����� ������������]";
TempDeyat->Active=true;
/*
for(Ref->First();!Ref->Eof;Ref->Next())
{
 TempDeyat->Append();
 TempDeyat->FieldByName("����� ������������")->Value=Ref->FieldByName("����� �����")->Value;
 TempDeyat->FieldByName("������������ ������������")->Value=Ref->FieldByName("��������")->Value;
 TempDeyat->FieldByName("�����")->Value=Ref->FieldByName("�����")->Value;
 TempDeyat->Post();
}
*/
MP<TADODataSet>Deyat(this);
Deyat->Connection=DB;
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
  Comm->CommandText="UPDATE ������� SET �������.������������ = 0 WHERE (((�������.������������)="+IntToStr(N)+"));";
  Comm->Execute();

   Deyat->Edit();
   Deyat->FieldByName("Del")->Value=true;
   Deyat->Post();
  }
 }
Comm->CommandText="UPDATE ������������ INNER JOIN ������� ON ������������.[����� ������������] = �������.[������������] SET �������.[������������] = 0 WHERE (((������������.Del)=True));";
Comm->Execute();

Comm->CommandText="DELETE ������������.Del FROM ������������ WHERE (((������������.Del)=True AND �����=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ������������ ( [����� ������������], [������������ ������������], ����� ) SELECT TempDeyat.[����� ������������], TempDeyat.[������������ ������������], TempDeyat.����� FROM TempDeyat;";
Comm->Execute();

Comm->CommandText="DELETE TempDeyat.* FROM TempDeyat;";
Comm->Execute();

if(Role==2)
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ������������� (��������)","");
}
else
{
Form1->Initialize();
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ������������� (������������)","");
}
Zast->BlockMK(false);
Zast->MClient->UnBlockServer("ReadWriteDoc");

}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadAspect1Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadAspect2");
Zast->MClient->ReadTable("HReference", "Select ����_7.[����� ����], ����_7.[��������], ����_7.[��������] From ����_7 Order by ��������, [����� ����];", "HReference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");
}
catch(...)
{
 Zast->BlockMK(false);
}

}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadAspect2Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeAspect1");
Zast->MClient->ReadTable("HReference", "Select �����_7.[����� �����], �����_7.[����� ��������], �����_7.[��������], �����_7.[�����] From �����_7 Order by [����� ��������], [����� �����];", "HReference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");
}
catch(...)
{
 Zast->BlockMK(false);
}

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeAspect1Execute(TObject *Sender)
{
try
{
Documents->MergeNode("����_7");
Documents->MergeBranch("�����_7");
Documents->LoadTab5();

String DBName;
if(Role==2)
{
 DBName="���������";
}
else
{
 DBName="���������_�";
}

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeAspect2");
Zast->MClient->ReadTable("���������", "Select ������.[����� �������], ������.[������������ �������], ������.[�����] From ������ order by [����� �������];", DBName, "Select TempAspect.[����� �������], TempAspect.[������������ �������], TempAspect.[�����] From TempAspect order by [����� �������];");

}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeAspect2Execute(TObject *Sender)
{
try
{
MDBConnector* DB;
if(Role==2)
{
 DB=ADOAspect;
}
else
{
 DB=ADOUsrAspect;
}

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="UPDATE ������ SET ������.Del = False;";
Comm->Execute();

MP<TADODataSet>TempAspect(this);
TempAspect->Connection=DB;
TempAspect->CommandText="Select TempAspect.[����� �������], TempAspect.[������������ �������], TempAspect.[�����] From TempAspect order by [����� �������]";
TempAspect->Active=true;

MP<TADODataSet>Aspect(this);
Aspect->Connection=DB;
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
  Comm->CommandText="UPDATE ������� SET �������.������ = 0 WHERE (((�������.������)="+IntToStr(N)+"));";
  Comm->Execute();

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

Zast->MClient->UnBlockServer("ReadWriteDoc");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteDocExecute(TObject *Sender)
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

void __fastcall TZast::WriteMetodikaExecute(TObject *Sender)
{
try
{
MClient->Act.ParamComm.clear();
MClient->Act.ParamComm.push_back("UnblockAndRWD");
MClient->WriteTable("HReference","Select �����, �������� From �������� Order by �����", "HReference","Select �����, �������� From �������� Order by �����");
Sleep(1000);
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ �������� (��������)","");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::WritePodrExecute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerPodr");
MClient->WriteTable("���������","Select �������������.[����� �������������], �������������.[�������� �������������], �������������.[ServerNum] from ������������� Order by [����� �������������]", "���������", "Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������], Temp�������������.[ServerNum] from Temp������������� Order by [����� �������������]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerPodrExecute(TObject *Sender)
{
try
{
 MClient->Act.WaitCommand=11;
 String DB="���������";
 String Mess="Command:11;1|"+IntToStr(DB.Length())+"#"+DB+"|";
 ClientSocket->Socket->SendText(Mess);
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------
void __fastcall TZast::WritePodr2Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerPodr2");
MClient->ReadTable("���������","Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������], Temp�������������.[ServerNum] from Temp������������� Order by [����� �������������]", "���������", "Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������], Temp�������������.[ServerNum] from Temp������������� Order by [����� �������������]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------
void __fastcall TZast::MergeServerPodr2Execute(TObject *Sender)
{
try
{
MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOAspect;
Temp->CommandText="Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������], Temp�������������.[ServerNum] from Temp������������� Order by [����� �������������]";
Temp->Active=true;

MP<TADODataSet>TempDoc(this);
TempDoc->Connection=Zast->ADOAspect;
TempDoc->CommandText="Select �������������.[����� �������������], �������������.[�������� �������������], �������������.[ServerNum] from �������������  Where ServerNum=0 Order by [����� �������������]";
TempDoc->Active=true;

for(TempDoc->First();!TempDoc->Eof;TempDoc->Next())
{
int N=TempDoc->FieldByName("����� �������������")->Value;
if(Temp->Locate("����� �������������", N, SO))
{
TempDoc->Edit();
TempDoc->FieldByName("ServerNum")->Value=Temp->FieldByName("ServerNum")->Value;
TempDoc->Post();
}
else
{
ShowMessage("������ ������ ������������� �����="+IntToStr(N));
}
}

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * from Temp�������������";
Comm->Execute();

 Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ ������������� (��������)","");

 //Zast->ReadWriteDoc->Execute();
 Zast->MClient->UnBlockServer("ReadWriteDoc");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------
void __fastcall TZast::WriteCritExecute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerCrit");
MClient->WriteTable("HReference","Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[��������1], ����������.[��������], ����������.[��� �������], ����������.[���� �������], ����������.[����������� ����] From ���������� Order by [����� ����������]", "HReference", "Select TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.[��������1], TempZn.[��������], TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] From TempZn Order by [����� ����������]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------
void __fastcall TZast::MergeServerCritExecute(TObject *Sender)
{
try
{
 MClient->Act.WaitCommand=12;
 String DB="HReference";
 String DB2="���������";
 String Mess="Command:12;2|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(DB2.Length())+"#"+DB2+"|";
 ClientSocket->Socket->SendText(Mess);
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteSitExecute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteSit2");
MClient->WriteTable("HReference","Select ��������.[����� ��������], ��������.[�������� ��������] from �������� Order by [����� ��������]", "HReference", "Select TempSit.[����� ��������], TempSit.[�������� ��������] from TempSit Order by [����� ��������]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteSit2Execute(TObject *Sender)
{
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=ADOAspect;
Comm->CommandText="Delete * From TempSit";
Comm->Execute();

MP<TADODataSet>RefSit(this);
RefSit->Connection=ADOConn;
RefSit->CommandText="Select ��������.[����� ��������], ��������.[�������� ��������] from �������� Order by [����� ��������]";
RefSit->Active=true;

MP<TADODataSet>TempSit(this);
TempSit->Connection=ADOAspect;
TempSit->CommandText="Select * from TempSit";
TempSit->Active=true;

for(RefSit->First();!RefSit->Eof;RefSit->Next())
{
 TempSit->Append();
 TempSit->FieldByName("����� ��������")->Value=RefSit->FieldByName("����� ��������")->Value;
 TempSit->FieldByName("�������� ��������")->Value=RefSit->FieldByName("�������� ��������")->Value;
 TempSit->Post();
}

Comm->CommandText="UPDATE �������� SET ��������.Del = False;";
Comm->Execute();



MP<TADODataSet>AspSit(this);
AspSit->Connection=ADOAspect;
AspSit->CommandText="Select * from �������� Order by [����� ��������]";
AspSit->Active=true;

for(AspSit->First();!AspSit->Eof;AspSit->Next())
{
 int Num=AspSit->FieldByName("����� ��������")->Value;
 if(TempSit->Locate("����� ��������", Num, SO))
 {
  AspSit->Edit();
  AspSit->FieldByName("�������� ��������")->Value=TempSit->FieldByName("�������� ��������")->Value;
  AspSit->FieldByName("�����")->Value=true;
  AspSit->Post();

  TempSit->Delete();
 }
 else
 {
  AspSit->Edit();
  AspSit->FieldByName("Del")->Value=true;
  AspSit->Post();
 }
}

Comm->CommandText="Delete * from �������� Where Del=true AND [�����]=true";
Comm->Execute();

Comm->CommandText="INSERT INTO �������� ( [����� ��������], [�������� ��������], ����� ) SELECT TempSit.[����� ��������], TempSit.[�������� ��������], True AS ���������1 FROM TempSit;";
Comm->Execute();

Comm->CommandText="UPDATE �������� SET ��������.Del = False;";
Comm->Execute();

Comm->CommandText="Delete * From TempSit";
Comm->Execute();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerSit");
MClient->WriteTable("���������","Select ��������.[����� ��������], ��������.[�������� ��������] from �������� Order by [����� ��������]", "���������", "Select TempSit.[����� ��������], TempSit.[�������� ��������] from TempSit Order by [����� ��������]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------
void __fastcall TZast::MergeServerSitExecute(TObject *Sender)
{
try
{
 MClient->Act.WaitCommand=13;
 String DB="HReference";
 String DB2="���������";
 String Mess="Command:13;2|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(DB2.Length())+"#"+DB2+"|";
 ClientSocket->Socket->SendText(Mess);
 }
 catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteVozd1Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteVozd2");
MClient->WriteTable("HReference","Select ����_3.[����� ����], ����_3.[��������], ����_3.[��������] From ����_3 Order by [��������], [����� ����]", "HReference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by [��������], [����� ����]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteVozd2Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerVozd");
MClient->WriteTable("HReference","Select �����_3.[����� �����], �����_3.[����� ��������], �����_3.[��������], �����_3.[�����] From �����_3 Order by [����� ��������], [����� �����]", "HReference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerVozdExecute(TObject *Sender)
{
try
{
 MClient->Act.WaitCommand=14;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerVozd");

 String DB="HReference";
 String Node="����_3";
 String Branch="�����_3";
 String DB2="���������";
 String Table="�����������";
 String AspField="�����������";
 String KeyTarget="����� �����������";
 String NameTarget="������������ �����������";

 String Mess="Command:14;8|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|"+IntToStr(DB2.Length())+"#"+DB2+"|"+IntToStr(Table.Length())+"#"+Table+"|"+IntToStr(AspField.Length())+"#"+AspField+"|"+IntToStr(KeyTarget.Length())+"#"+KeyTarget+"|"+IntToStr(NameTarget.Length())+"#"+NameTarget+"|";
 ClientSocket->Socket->SendText(Mess);
 }
 catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerVozdExecute(TObject *Sender)
{
try
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ ����������� (��������)","");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteMeropr1Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteMeropr2");
MClient->WriteTable("HReference","Select ����_4.[����� ����], ����_4.[��������], ����_4.[��������] From ����_4 Order by [��������], [����� ����]", "HReference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by [��������], [����� ����]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteMeropr2Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerMeropr");
MClient->WriteTable("HReference","Select �����_4.[����� �����], �����_4.[����� ��������], �����_4.[��������], �����_4.[�����] From �����_4 Order by [����� ��������], [����� �����]", "HReference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerMeroprExecute(TObject *Sender)
{
try
{
 MClient->Act.WaitCommand=15;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerMeropr");

 //"Reference", "����_4","�����_4"
 String DB="HReference";
 String Node="����_4";
 String Branch="�����_4";


 String Mess="Command:15;3|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|";
 ClientSocket->Socket->SendText(Mess);
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerMeroprExecute(TObject *Sender)
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ ����������� (��������)","");
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteTerr1Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteTerr2");
MClient->WriteTable("HReference","Select ����_5.[����� ����], ����_5.[��������], ����_5.[��������] From ����_5 Order by [��������], [����� ����]", "HReference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by [��������], [����� ����]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteTerr2Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerTerr");
MClient->WriteTable("HReference","Select �����_5.[����� �����], �����_5.[����� ��������], �����_5.[��������], �����_5.[�����] From �����_5 Order by [����� ��������], [����� �����]", "HReference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerTerrExecute(TObject *Sender)
{
try
{
 MClient->Act.WaitCommand=14;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerTerr");

 String DB="HReference";
 String Node="����_5";
 String Branch="�����_5";
 String DB2="���������";
 String Table="����������";
 String AspField="��� ����������";
 String KeyTarget="����� ����������";
 String NameTarget="������������ ����������";

 String Mess="Command:14;8|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|"+IntToStr(DB2.Length())+"#"+DB2+"|"+IntToStr(Table.Length())+"#"+Table+"|"+IntToStr(AspField.Length())+"#"+AspField+"|"+IntToStr(KeyTarget.Length())+"#"+KeyTarget+"|"+IntToStr(NameTarget.Length())+"#"+NameTarget+"|";
 ClientSocket->Socket->SendText(Mess);
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerTerrExecute(TObject *Sender)
{
try
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ ���������� (��������)","");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteDeyat1Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteDeyat2");
MClient->WriteTable("HReference","Select ����_6.[����� ����], ����_6.[��������], ����_6.[��������] From ����_6 Order by [��������], [����� ����]", "HReference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by [��������], [����� ����]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteDeyat2Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerDeyat");
MClient->WriteTable("HReference","Select �����_6.[����� �����], �����_6.[����� ��������], �����_6.[��������], �����_6.[�����] From �����_6 Order by [����� ��������], [����� �����]", "HReference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerDeyatExecute(TObject *Sender)
{
try
{
 MClient->Act.WaitCommand=14;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerDeyat");

 String DB="HReference";
 String Node="����_6";
 String Branch="�����_6";
 String DB2="���������";
 String Table="������������";
 String AspField="������������";
 String KeyTarget="����� ������������";
 String NameTarget="������������ ������������";

 String Mess="Command:14;8|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|"+IntToStr(DB2.Length())+"#"+DB2+"|"+IntToStr(Table.Length())+"#"+Table+"|"+IntToStr(AspField.Length())+"#"+AspField+"|"+IntToStr(KeyTarget.Length())+"#"+KeyTarget+"|"+IntToStr(NameTarget.Length())+"#"+NameTarget+"|";
 ClientSocket->Socket->SendText(Mess);
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerDeyatExecute(TObject *Sender)
{
try
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ ����� ������������ (��������)","");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteAspect1Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteAspect2");
MClient->WriteTable("HReference","Select ����_7.[����� ����], ����_7.[��������], ����_7.[��������] From ����_7 Order by [��������], [����� ����]", "HReference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by [��������], [����� ����]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteAspect2Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerAspect");
MClient->WriteTable("HReference","Select �����_7.[����� �����], �����_7.[����� ��������], �����_7.[��������], �����_7.[�����] From �����_7 Order by [����� ��������], [����� �����]", "HReference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����]");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerAspectExecute(TObject *Sender)
{
try
{
 MClient->Act.WaitCommand=14;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerAspect");

 String DB="HReference";
 String Node="����_7";
 String Branch="�����_7";
 String DB2="���������";
 String Table="������";
 String AspField="������";
 String KeyTarget="����� �������";
 String NameTarget="������������ �������";

 String Mess="Command:14;8|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|"+IntToStr(DB2.Length())+"#"+DB2+"|"+IntToStr(Table.Length())+"#"+Table+"|"+IntToStr(AspField.Length())+"#"+AspField+"|"+IntToStr(KeyTarget.Length())+"#"+KeyTarget+"|"+IntToStr(NameTarget.Length())+"#"+NameTarget+"|";
 ClientSocket->Socket->SendText(Mess);
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerAspectExecute(TObject *Sender)
{
try
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ ������ ������������� �������� (��������)","");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------


void __fastcall TZast::ContSvodReportExecute(TObject *Sender)
{
 FSvod->ContSvodReport();
}
//---------------------------------------------------------------------------


void __fastcall TZast::StartLoadPodrExecute(TObject *Sender)
{
//��������� �������� �������������
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadObslOtd");
Zast->MClient->ReadTable("���������", "Select �������������.[����� �������������], �������������.[�������� �������������] From ������������� Order by [����� �������������];", "���������", "Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������] From Temp������������� Order by [����� �������������];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::StartLoadObslOtdExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 //Zast->MClient->Act.ParamComm.push_back("StartMergeLoginsPodr");
 Zast->MClient->Act.ParamComm.push_back("PrepStartMergeLoginPodr");
Zast->MClient->ReadTable("���������", "Select ObslOtdel.Login, ObslOtdel.NumObslOtdel From ObslOtdel Order by Login, NumObslOtdel;", "���������", "Select TempObslOtdel.Login, TempObslOtdel.NumObslOtdel From TempObslOtdel Order by Login, NumObslOtdel;");

}
//---------------------------------------------------------------------------

void __fastcall TZast::StartMergeLoginsPodrExecute(TObject *Sender)
{
MDBConnector *DB;
if(Role==2)
{
DB=ADOAspect;
}
else
{
DB=ADOUsrAspect;
}
//����������� ���������� � ������� �������, ������������� � ������������� �������������
//��� ������ ������ ��������� ��� ����������
//����� �� ����������� ������������� �� ������ ���� �������

//���������� ������� �������������
MergeOtdels();
//���������� ������� �������, ������������ ����������� ������� TempObslOtdel
MergeLogins();

//���������� ������� ������������� �������������
MP<TADODataSet>ToTable(this);
ToTable->Connection=DB;
ToTable->CommandText="Select TempObslOtdel.Login, TempObslOtdel.NumObslOtdel From TempObslOtdel Order by Login, NumObslOtdel";
ToTable->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="Delete * From ObslOtdel";
Comm->Execute();

MP<TADODataSet>Logins(this);
Logins->Connection=DB;
Logins->CommandText="Select * from Logins";
Logins->Active=true;

MP<TADODataSet>Podr(this);
Podr->Connection=DB;
Podr->CommandText="Select * from �������������";
Podr->Active=true;

int Log;
int Pod;
for(ToTable->First();!ToTable->Eof;ToTable->Next())
{
 int N=ToTable->FieldByName("Login")->Value;

 if(Logins->Locate("ServerNum", N, SO))
 {
  Log=Logins->FieldByName("Num")->Value;

 int N1=ToTable->FieldByName("NumObslOtdel")->Value;

 if(Podr->Locate("ServerNum", N1, SO))
 {
  Pod=Podr->FieldByName("����� �������������")->Value;

  ToTable->Edit();
  ToTable->FieldByName("Login")->Value=Log;
  ToTable->FieldByName("NumObslOtdel")->Value=Pod;
  ToTable->Post();
 }
 else
 {
  ShowMessage("������ ������������� ������������� N="+IntToStr(N1));
 }
 }
 else
 {
  ShowMessage("������ ������������� ������� N="+IntToStr(N));
 }
}
Comm->CommandText="INSERT INTO ObslOtdel ( Login, NumObslOtdel) SELECT TempObslOtdel.Login, TempObslOtdel.NumObslOtdel FROM TempObslOtdel;";
Comm->Execute();

Comm->CommandText="Delete * From TempObslOtdel";
Comm->Execute();

if(Role==2)
{

Documents->Show();
}
else
{
 MP<TADODataSet>Log(this);
 Log->Connection=ADOUsrAspect;
 Log->CommandText="Select * From Logins Where Login='"+Form1->Login+"'";
 Log->Active=true;

 Form1->NumLogin=Log->FieldByName("ServerNum")->AsInteger;

Prog->SignComplete=false;
Prog->Show();
Prog->PB->Min=0;
Prog->PB->Position=0;
Prog->PB->Max=9;

Documents->ReadWrite.clear();
Str_RW S;
S.NameAction="ReadMetodika";
S.Text="������ ��������...";
S.Num=1;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadPodrazd";
S.Text="������ �������������...";
S.Num=2;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadCrit";
S.Text="������ ���������...";
S.Num=3;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadSit";
S.Text="������ ��������...";
S.Num=4;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadVozd1";
S.Text="������ ������ �����������...";
S.Num=5;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadMeropr1";
S.Text="������ ������ �����������...";
S.Num=6;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadTerr1";
S.Text="������ ������ ����������...";
S.Num=7;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadDeyat1";
S.Text="������ ������ ����� ������������...";
S.Num=8;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadAspect1";
S.Text="������ ������ ������������� ��������...";
S.Num=9;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadTempAsp";
S.Num=9;
Documents->ReadWrite.push_back(S);

S.NameAction="ShowForm1";
S.Num=9;
Documents->ReadWrite.push_back(S);

Zast->ReadWriteDoc->Execute();
}
}
//---------------------------------------------------------------------------
void TZast::MergeLogins()
{
MDBConnector *DB;
if(Role==2)
{
DB=ADOAspect;
}
else
{
DB=ADOUsrAspect;
}


MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="UPDATE Logins SET Logins.Del = False;";
Comm->Execute();

MP<TADODataSet>TempLogins(this);
TempLogins->Connection=DB;
TempLogins->CommandText="Select * From TempLogins";
TempLogins->Active=true;

MP<TADODataSet>Logins(this);
Logins->Connection=DB;
Logins->CommandText="Select * From Logins";
Logins->Active=true;

for(Logins->First();!Logins->Eof;Logins->Next())
{
 int N=Logins->FieldByName("ServerNum")->Value;
 if(TempLogins->Locate("Num", N, SO))
 {
  Logins->Edit();
  Logins->FieldByName("Login")->Value=TempLogins->FieldByName("Login")->Value;
  Logins->FieldByName("Role")->Value=TempLogins->FieldByName("Role")->Value;
  Logins->Post();

  TempLogins->Delete();
 }
 else
 {
  Logins->Edit();
  Logins->FieldByName("Del")->Value=true;
  Logins->Post();
 }
}

Comm->CommandText="Delete * From Logins Where Logins.Del = true;";
Comm->Execute();

Comm->CommandText="INSERT INTO Logins ( ServerNum, Login, Role ) SELECT TempLogins.Num, TempLogins.Login, TempLogins.Role FROM TempLogins;";
Comm->Execute();

}
//---------------------------------------------------------------
void TZast::MergeOtdels()
{
TLocateOptions SO;

MDBConnector *DB;
if(Role==2)
{
DB=ADOAspect;
}
else
{
DB=ADOUsrAspect;
}

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="UPDATE ������������� SET �������������.Del = False;";
Comm->Execute();

MP<TADODataSet>Otdels(this);
Otdels->Connection=DB;
Otdels->CommandText="Select * From ������������� Order by [����� �������������]";
Otdels->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=DB;
Temp->CommandText="Select * From Temp������������� Order by [����� �������������]";
Temp->Active=true;

for(Otdels->First();!Otdels->Eof;Otdels->Next())
{
int SNum=Otdels->FieldByName("ServerNum")->AsInteger;
if(Temp->Locate("����� �������������", SNum, SO))
{
//������
//�������� �������� �������������
//������� ������ �� Temp�������������
Otdels->Edit();
Otdels->FieldByName("�������� �������������")->Value=Temp->FieldByName("�������� �������������")->Value;
Otdels->Post();

Temp->Delete();
}
else
{
//��������
//��� ������������� �� ������� �������
//�������� � ��������

}
}

//������� ������ �������������

//���� � Temp������������� �������� ������ �� ��� ����� ������������� �� �������
//��������� �� � �������������

for(Temp->First();!Temp->Eof;Temp->Next())
{
Otdels->Insert();
Otdels->FieldByName("�������� �������������")->Value=Temp-> FieldByName("�������� �������������")->Value;
Otdels->FieldByName("ServerNum")->Value=Temp-> FieldByName("����� �������������")->Value;
Otdels->FieldByName("Del")->Value=false;
Otdels->Post();
}
//�������� Temp�������������
Comm->CommandText="Delete * from Temp�������������";
Comm->Execute();


}
//----------------------------------------------------
void __fastcall TZast::CompareMSpecAspectsExecute(TObject *Sender)
{
//��������� ������� �������� ��� ��������� �� ����� ������� � �������
//��� �������� ������� � ������������� ������ ������� ��������
MP<TADODataSet>Tab(this);
Tab->Connection=ADOAspect;

Tab->CommandText="SELECT �������.[����� �������], �������.��������, �������.[��� ����������], �������.������������, �������.������, �������.�����������,  �������.Z, �������.���������� FROM ������� ORDER BY �������.[����� �������];";
Tab->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=ADOAspect;
Temp->CommandText="SELECT TempAspects.[����� �������], TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.������, TempAspects.�����������, TempAspects.Z, TempAspects.���������� FROM TempAspects ORDER BY TempAspects.[����� �������]";
Temp->Active=true;

bool Res=false;
String Mess;
if(Tab->RecordCount==Temp->RecordCount)
{
 Temp->First();
 for(Tab->First();!Tab->Eof;Tab->Next())
 {
  for(int i=0;i<Tab->Fields->Count;i++)
  {

   if(Tab->FieldList->Fields[i]->AsString==Temp->FieldList->Fields[i]->AsString)
   {
    Res=true;
   }
   else
   {
    Res=false;
    Mess="���������� �������� �� ������� �� ��������� � ����������� �������� � ��������� ���� ������\r�������� ������ ��������?";
    break;
   }
  }
  if(!Res)
  {
   break;
  }
  Temp->Next();
 }
}
else
{
 Mess="���������� �������� �� ������� �� ��������� � ����������� �������� � ��������� ���� ������\r�������� ������ ��������?";
}

if(!Res)
{
 if(Application->MessageBoxA(Mess.c_str(),"���������� ��������",MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON1)==IDYES)
 {
  //���������� ������ �������� ��� ���������
MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOAspect;
Temp->CommandText="Select * From TempAspects";
Temp->Active=true;

MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOAspect;
Podr->CommandText="Select * From �������������";
Podr->Active=true;

for(Temp->First();!Temp->Eof;Temp->Next())
{
 int N=Temp->FieldByName("�������������")->Value;

 if(Podr->Locate("ServerNum", N, SO))
 {
  int Num=Podr->FieldByName("����� �������������")->Value;

  Temp->Edit();
  Temp->FieldByName("�������������")->Value=Num;
  Temp->Post();
 }
 else
 {
  ShowMessage("������ ����������� ��������");
 }
}

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;



Comm->CommandText="Delete * From �������";
Comm->Execute();


String CT="INSERT INTO ������� ( [����� �������], �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, G, O, R, S, T, L, N, Z, ����������, [���������� �����������], [������� �����������], ��������������, [������������� �����������], [������������ �����������], [���������� � ��������], [������������ ���������� � ��������], �����������, [���� ��������], [������ ��������], [����� ��������] ) ";
CT=CT+" SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������], TempAspects.[������������ �����������], TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������], TempAspects.�����������, TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] ";
CT=CT+" FROM TempAspects;";
Comm->CommandText=CT;
Comm->Execute();

Comm->CommandText="Delete * From TempAspects";
Comm->Execute();
 }
}

MAsp->MoveAspects->Active=false;
String CT="SELECT �������.[����� �������], �������������.[����� �������������], �������������.[�������� �������������], ����������.[����� ����������], ����������.[������������ ����������], ������������.[����� ������������], ������������.[������������ ������������], ������.[����� �������], ������.[������������ �������], �����������.[����� �����������], �����������.[������������ �����������], �������.����������, �������.Z, ��������.[����� ��������], ��������.[�������� ��������], �������.������������� FROM �������� INNER JOIN (����������� INNER JOIN (������ INNER JOIN (������������ INNER JOIN (���������� INNER JOIN (������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) ON ����������.[����� ����������] = �������.[��� ����������]) ON ������������.[����� ������������] = �������.������������) ON ������.[����� �������] = �������.������) ON �����������.[����� �����������] = �������.�����������) ON ��������.[����� ��������] = �������.��������";
CT=CT+" Order by �������.[����� �������]; ";

MAsp->MoveAspects->CommandText=CT;
MAsp->MoveAspects->Connection=Zast->ADOAspect;
MAsp->MoveAspects->Active=true;

MClient->UnBlockServer("");
MAsp->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TZast::SaveAspectsMSpecExecute(TObject *Sender)
{
try
{
Zast->MClient->Act.WaitCommand=17;
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("EndsaveAspectsMSpec");

ClientSocket->Socket->SendText("Command:17;0|");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::EndsaveAspectsMSpecExecute(TObject *Sender)
{
try
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ �������� (��������)","");
Sleep(1000);
 Zast->MClient->UnBlockServer("ReadWriteDoc");
}
catch(...)
{
 Zast->BlockMK(false);
}
//ShowMessage("������ ������ ���������");
}
//---------------------------------------------------------------------------

void __fastcall TZast::CompareMSpecAspects2Execute(TObject *Sender)
{
//��������� ������� �������� ��� ��������� �� ����� ������� � �������
//��� �������� ������� � ������������� ������ ������� ��������


MP<TADODataSet>Tab(this);
Tab->Connection=ADOAspect;
Tab->CommandText="SELECT �������.[����� �������], �������������.ServerNum FROM ������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� ORDER BY �������.[����� �������];";
Tab->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=ADOAspect;
Temp->CommandText="SELECT TempAspects.[����� �������], TempAspects.������������� From TempAspects ORDER BY [����� �������]";
Temp->Active=true;

bool Res=false;
String Mess;
if(Tab->RecordCount==Temp->RecordCount)
{
 Temp->First();
 for(Tab->First();!Tab->Eof;Tab->Next())
 {
  for(int i=0;i<Tab->Fields->Count;i++)
  {

   if(Tab->FieldList->Fields[i]->AsString==Temp->FieldList->Fields[i]->AsString)
   {
    Res=true;
   }
   else
   {
    Res=false;
    Mess="���������� �������� �� ������� �� ��������� � ����������� �������� � ��������� ���� ������\r�������� ������ �������� �� ������?";
    break;
   }
  }
  if(!Res)
  {
   break;
  }
  Temp->Next();
 }
}
else
{
 Mess="���������� �������� �� ������� �� ��������� � ����������� �������� � ��������� ���� ������\r�������� ������ �������� �� ������?";
}

if(!Res)
{

 if(Application->MessageBoxA(Mess.c_str(),"������ ��������",MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON1)==IDYES)
 {
  //������ ������ �������� ����������
  /*
Documents->ReadWrite.clear();
Str_RW S;
Documents->ReadWrite.push_back(S);
 */
Zast->BlockMK(true);
try
{
Zast->MClient->Act.ParamComm.clear();
Zast->MClient->Act.ParamComm.push_back("SaveAspectsMSpec2");
Zast->MClient->WriteTable("���������","SELECT �������.[����� �������], �������������.ServerNum FROM ������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� Order by �������.[����� �������]; ", "���������", "SELECT TempAspects.[����� �������], TempAspects.������������� From TempAspects order by [����� �������];");
}
catch(...)
{
 Zast->BlockMK(false);
}
//Zast->ReadWriteDoc->Execute();
 }
 else
 {

 Zast->MClient->WriteDiaryEvent("NetAspects","����� �� ���������� �������� �������� (��������)","");
 }
}

}
//---------------------------------------------------------------------------

void __fastcall TZast::SaveAspectsMSpec2Execute(TObject *Sender)
{
try
{
Zast->MClient->Act.WaitCommand=17;
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("EndsaveAspectsMSpec2");

ClientSocket->Socket->SendText("Command:17;0|");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::EndsaveAspectsMSpec2Execute(TObject *Sender)
{
try
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ �������� (��������)","");
}
catch(...)
{
 Zast->BlockMK(false);
}

Zast->BlockMK(false);
}
//---------------------------------------------------------------------------

void __fastcall TZast::LoadMSpecAspectsExecute(TObject *Sender)
{
try
{
MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOAspect;
Temp->CommandText="Select * From TempAspects";
Temp->Active=true;

MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOAspect;
Podr->CommandText="Select * From �������������";
Podr->Active=true;

for(Temp->First();!Temp->Eof;Temp->Next())
{
 int N=Temp->FieldByName("�������������")->Value;

 if(Podr->Locate("ServerNum", N, SO))
 {
  int Num=Podr->FieldByName("����� �������������")->Value;

  Temp->Edit();
  Temp->FieldByName("�������������")->Value=Num;
  Temp->Post();
 }
 else
 {
  ShowMessage("������ ����������� ��������");
 }
}

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;



Comm->CommandText="Delete * From �������";
Comm->Execute();


String CT="INSERT INTO ������� ( [����� �������], �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, G, O, R, S, T, L, N, Z, ����������, [���������� �����������], [������� �����������], ��������������, [������������� �����������], [������������ �����������], [���������� � ��������], [������������ ���������� � ��������], �����������, [���� ��������], [������ ��������], [����� ��������] ) ";
CT=CT+" SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������], TempAspects.[������������ �����������], TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������], TempAspects.�����������, TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] ";
CT=CT+" FROM TempAspects;";
Comm->CommandText=CT;
Comm->Execute();

Comm->CommandText="Delete * From TempAspects";
Comm->Execute();

MAsp->MoveAspects->Active=false;
CT="SELECT �������.[����� �������], �������������.[����� �������������], �������������.[�������� �������������], ����������.[����� ����������], ����������.[������������ ����������], ������������.[����� ������������], ������������.[������������ ������������], ������.[����� �������], ������.[������������ �������], �����������.[����� �����������], �����������.[������������ �����������], �������.����������, �������.Z, ��������.[����� ��������], ��������.[�������� ��������], �������.������������� FROM �������� INNER JOIN (����������� INNER JOIN (������ INNER JOIN (������������ INNER JOIN (���������� INNER JOIN (������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) ON ����������.[����� ����������] = �������.[��� ����������]) ON ������������.[����� ������������] = �������.������������) ON ������.[����� �������] = �������.������) ON �����������.[����� �����������] = �������.�����������) ON ��������.[����� ��������] = �������.��������";
CT=CT+" Order by �������.[����� �������]; ";

MAsp->MoveAspects->CommandText=CT;
MAsp->MoveAspects->Connection=Zast->ADOAspect;
MAsp->MoveAspects->Active=true;

MAsp->ChangeCPodr();
 Zast->MClient->UnBlockServer("EndReadAspectsMSpec");
 Zast->BlockMK(false);
ShowMessage("������ ���������");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::SaveAspectsMSpec0Execute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("SaveAspectsMSpec");
Zast->MClient->WriteTable("���������","SELECT �������.[����� �������], �������������.ServerNum FROM ������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� Order by �������.[����� �������]; ", "���������", "SELECT TempAspects.[����� �������], TempAspects.������������� From TempAspects order by [����� �������];");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::StopProgramExecute(TObject *Sender)
{
MClient->UnBlockServer("StopProgram2");

}
//---------------------------------------------------------------------------

void __fastcall TZast::StartLoadPodrUSRExecute(TObject *Sender)
{
//��������� �������� �������������
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadObslOtdUSR");
Zast->MClient->ReadTable("���������", "Select �������������.[����� �������������], �������������.[�������� �������������] From ������������� Order by [����� �������������];", "���������_�", "Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������] From Temp������������� Order by [����� �������������];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::StartLoadObslOtdUSRExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 //Zast->MClient->Act.ParamComm.push_back("StartMergeLoginsPodr");
 Zast->MClient->Act.ParamComm.push_back("PrepStartMergeLoginPodr");
Zast->MClient->ReadTable("���������", "Select ObslOtdel.Login, ObslOtdel.NumObslOtdel From ObslOtdel Order by Login, NumObslOtdel;", "���������_�", "Select TempObslOtdel.Login, TempObslOtdel.NumObslOtdel From TempObslOtdel Order by Login, NumObslOtdel;");

}
//---------------------------------------------------------------------------

void __fastcall TZast::ShowForm1Execute(TObject *Sender)
{
Form1->Show();
MClient->UnBlockServer("");
ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeAspectsUserExecute(TObject *Sender)
{
try
{

MClient->UnBlockServer("MergeAspectsUser1");



}
catch(...)
{
 Zast->BlockMK(false);
}

Zast->BlockMK(false);
}
//---------------------------------------------------------------------------

void  TZast::MergeAspects(int NumLogin, bool Quit)
{
try
{
MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOUsrAspect;
Podr->CommandText="Select * From �������������";
Podr->Active=true;

MP<TADODataSet>TempAsp(this);
TempAsp->Connection=Zast->ADOUsrAspect;
TempAsp->CommandText="Select * From TempAspects";
TempAsp->Active=true;

for(TempAsp->First();!TempAsp->Eof;TempAsp->Next())
{
 int N=TempAsp->FieldByName("�������������")->Value;

 if(Podr->Locate("ServerNum",N,SO))
 {
  int Num=Podr->FieldByName("����� �������������")->Value;

  TempAsp->Edit();
  TempAsp->FieldByName("�������������")->Value=Num;
  TempAsp->Post();
 }
 else
 {
 Zast->MClient->WriteDiaryEvent("NetAspects ������","���� ���������� ��������","");
  ShowMessage("������ ����������� ��������");
 }
}
//������������ ���������.
//������� ������ ��� ��� �������, �� ���� ����������

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
/*
Comm->CommandText="Delete * from �������";
Comm->Execute();
*/
String ST="INSERT INTO ������� ( �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, G, O, R, S, T, L, N, Z, ����������, [���������� �����������], [������� �����������], ��������������, [������������� �����������], [������������ �����������], [���������� � ��������], [������������ ���������� � ��������], �����������, [���� ��������], [������ ��������], [����� ��������], ServerNum ) ";
ST=ST+" SELECT TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������], TempAspects.[������������ �����������], TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������], TempAspects.�����������, TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������], TempAspects.[����� �������] ";
ST=ST+" FROM TempAspects; ";
Comm->CommandText=ST;
Comm->Execute();

String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Hazards.ini");

int Num=Ini->ReadInteger(IntToStr(NumLogin),"CurrentRecord",1);
Filter->CText=Ini->ReadString(IntToStr(NumLogin),"Filter","");
if(Filter->CText=="")
{
 Filter->SetDefFiltr();
}

Form1->LFiltr->Caption=Ini->ReadString(IntToStr(NumLogin),"NameFilter","��������");
Filter->RadioGroup1->ItemIndex=Ini->ReadInteger(IntToStr(NumLogin),"NumFilter", 0);
switch (Filter->RadioGroup1->ItemIndex)
{
 case 1:
 {
Filter->ComboBox1->Text=Ini->ReadString(IntToStr(NumLogin),"TextFilter","");
 break;
 }
 case 2:
 {
Filter->ComboBox4->Text=Ini->ReadString(IntToStr(NumLogin),"TextFilter","");
 break;
 }
 case 3:
 {
Filter->ComboBox5->Text=Ini->ReadString(IntToStr(NumLogin),"TextFilter","");
 break;
 }
 case 4:
 {
Filter->ComboBox6->Text=Ini->ReadString(IntToStr(NumLogin),"TextFilter","");
 break;
 }
 case 5:
 {
Filter->ComboBox7->Text=Ini->ReadString(IntToStr(NumLogin),"TextFilter","");
 break;
 }
 case 6:
 {
Filter->ComboBox2->Text=Ini->ReadString(IntToStr(NumLogin),"TextFilter","");
 break;
 }
 case 7:
 {
Filter->ComboBox3->Text=Ini->ReadString(IntToStr(NumLogin),"TextFilter","");
 break;
 }

}
Form1->Initialize(Num);

if(!Quit)
{
ShowMessage("���������");
}
else
{
this->Close();
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ���������� ��������"," ������ "+IntToStr(GetLastError()));

}
}
//----------------------------------------------------------------------
void __fastcall TZast::WriteAspectsUsrExecute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.WaitCommand=18;
Zast->ClientSocket->Socket->SendText("Command:18;1|"+IntToStr(IntToStr(Form1->NumLogin).Length())+"#"+IntToStr(Form1->NumLogin));
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeAspectsUserQExecute(TObject *Sender)
{
MClient->UnBlockServer("PrepMergeAspectsUserQ");
}
//---------------------------------------------------------------------------

void __fastcall TZast::PreLoadLoginsExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ViewLogins");

 Zast->MClient->ReadTable("���������","Select Login, Code1, Code2, Role from Logins Where Role<>1", "���������", "Select Login, Code1, Code2, Role From TempLogins");

}
//---------------------------------------------------------------------------

void __fastcall TZast::BlockServerTimer(TObject *Sender)
{
/*
if(Prog->Position<Prog->PB->Max)
{
Prog->PB->Position++;
}
else
{

}
*/
try
{
ClientSocket->Socket->SendText("Command:19;1|1#1");
BlockServer->Enabled=false;
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------
void TZast::WaitBlockServer(bool Flag)
{
String Mess="������ �����, �������...";
if(Flag)
{
 //�������� ��������
 Prog->Label1->Caption=Mess;
 /*
 Prog->PB->Min=0;
 Prog->PB->Position=0;
 Prog->PB->Max=9;
 */

 Prog->Show();
 BlockServer->Enabled=true;
}
else
{
 //�������� ���������
  BlockServer->Enabled=false;
  
/*
 if(Prog->Label1->Caption==Mess)
 {
 Prog->Close();
 }
*/
}
}
//---------------------------------------------------------------------------
void __fastcall TZast::UnBlockServerTimer(TObject *Sender)
{
try
{
 Zast->MClient->Act.WaitCommand=20;
 ClientSocket->Socket->SendText("Command:20;1|1#0");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::UnblockAndRWDExecute(TObject *Sender)
{
try
{
Zast->MClient->UnBlockServer("ReadWriteDoc");
}
catch(...)
{
Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadAspectsMSpecExecute(TObject *Sender)
{
try
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("LoadMSpecAspects");
 String ServerSQL="SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������,  �������.[������������� �����������],  �������.[������������ �����������],  �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.[���� ��������], �������.[������ ��������], �������.[����� ��������] FROM �������;";
 String ClientSQL="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������],  TempAspects.[������������ �����������],  TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������],  TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;";
Zast->MClient->ReadTable("���������",ServerSQL, "���������", ClientSQL);
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::PrepareCompareMSpecAspectsExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("CompareMSpecAspects");
 String ServerSQL="SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������,  �������.[������������� �����������],  �������.[������������ �����������],  �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.[���� ��������], �������.[������ ��������], �������.[����� ��������] FROM �������;";
 String ClientSQL="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������],  TempAspects.[������������ �����������],  TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������],  TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;";
Zast->MClient->ReadTable("���������",ServerSQL, "���������", ClientSQL);
}
//---------------------------------------------------------------------------

void __fastcall TZast::EndReadAspectsMSpecExecute(TObject *Sender)
{
try
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ �������� (��������)","");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::StopProgram2Execute(TObject *Sender)
{
Zast->Close();        
}
//---------------------------------------------------------------------------

void __fastcall TZast::PrepStartMergeLoginPodrExecute(TObject *Sender)
{
MClient->UnBlockServer("StartMergeLoginsPodr");
}
//---------------------------------------------------------------------------

void __fastcall TZast::PrepareReadAspectsUsrExecute(TObject *Sender)
{
try
{

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="DELETE Logins.AdmNum, �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+"));";
Comm->Execute();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeAspectsUser");
 String ServerSQL="SELECT �������.[����� �������],     �������.�������������,     �������.��������,     �������.[��� ����������],     �������.������������,     �������.�������������,     �������.������,     �������.�����������,     �������.G,     �������.O,     �������.R,     �������.S,     �������.T,     �������.L,     �������.N,     �������.Z,     �������.����������,     �������.[���������� �����������],     �������.[������� �����������],     �������.��������������,     �������.[������������� �����������],     �������.[������������ �����������],     �������.[���������� � ��������],     �������.[������������ ���������� � ��������],     �������.�����������,      �������.[���� ��������],     �������.[������ ��������],     �������.[����� ��������] FROM (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� WHERE (((ObslOtdel.Login)="+IntToStr(Form1->NumLogin)+"));";
 String ClientSQL="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������], TempAspects.[������������ �����������], TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������], TempAspects.�����������,  TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;";
Zast->MClient->ReadTable("���������", ServerSQL, "���������_�", ClientSQL);
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeAspectsUser1Execute(TObject *Sender)
{
try
{
MergeAspects(Form1->NumLogin, false);

ReadTempAsp->Execute();
Form1->CountInvalid();
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ �������� (������������)","");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::PrepWriteAspUsrExecute(TObject *Sender)
{




Zast->MClient->Act.ParamComm.clear();
Zast->MClient->Act.ParamComm.push_back("WriteAspectsUsr");
String ClientSQL="SELECT �������.[ServerNum],      �������������.ServerNum,     �������.��������,     �������.[��� ����������],     �������.������������,     �������.�������������,     �������.������,     �������.�����������,     �������.G,     �������.O,     �������.R,     �������.S,     �������.T,     �������.L,     �������.N,     �������.Z,     �������.����������,     �������.[���������� �����������],     �������.[������� �����������],     �������.��������������,     �������.[������������� �����������],     �������.[������������ �����������],     �������.[���������� � ��������],     �������.[������������ ���������� � ��������],     �������.�����������,      �������.[���� ��������],     �������.[������ ��������],     �������.[����� ��������] FROM ������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� ORDER BY �������.[����� �������];";
String ServerSQL="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������], TempAspects.[������������ �����������], TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������], TempAspects.�����������,  TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;";
Zast->MClient->WriteTable("���������_�", ClientSQL, "���������", ServerSQL);

}
//---------------------------------------------------------------------------

void __fastcall TZast::PrepMergeAspectsUserQExecute(TObject *Sender)
{
MergeAspects(Form1->NumLogin, true);
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ �������� (������������)","");
}
//---------------------------------------------------------------------------

void __fastcall TZast::ContStartReportsExecute(TObject *Sender)
{
 String ServerSQL;





//������� TempAspects
MP<TADOCommand>Comm(this);
Comm->Connection=Report1->RepBase;
Comm->CommandText="Delete * From TempAspects";
Comm->Execute();


if(Role==4)
{
 //����������������
 //����� ��������� ������

 //�������� �� ����� ������ �������������!!!

/*
String S="INSERT INTO TempAspects ( [����� �������], �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, G, O, R, S, T, L, N, Z, ����������, [���������� �����������], [������� �����������], ��������������, [������������� �����������], [������������ �����������], [���������� � ��������], [������������ ���������� � ��������], �����������, [���� ��������], [������ ��������], [����� ��������], ServerNum ) ";
S=S+" SELECT TOP 2 �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������, �������.[������������� �����������], �������.[������������ �����������], �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.�����������, �������.[���� ��������], �������.[������ ��������], �������.[����� ��������], �������.ServerNum ";
S=S+" FROM �������;";
*/
String S="INSERT INTO TempAspects ( [����� �������], �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, G, O, R, S, T, L, N, Z, ����������, [���������� �����������], [������� �����������], ��������������, [������������� �����������], [������������ �����������], [���������� � ��������], [������������ ���������� � ��������], �����������, [���� ��������], [������ ��������], [����� ��������], ServerNum ) ";
S=S+" SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������, �������.[������������� �����������], �������.[������������ �����������], �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.�����������, �������.[���� ��������], �������.[������ ��������], �������.[����� ��������], �������.ServerNum ";
S=S+" FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login ";
S=S+" WHERE (((Logins.Role)=4));";
Comm->CommandText=S;
Comm->Execute();

ContStartReports2->Execute();
 /*
Report1->CreateRep();
*/
}
else
{
 //�������� ��� ������������
 //����������� ������ � �������
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ContStartReports2");
// String ServerSQL="SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������,  �������.[������������� �����������],  �������.[������������ �����������],  �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.[���� ��������], �������.[������ ��������], �������.[����� ��������] FROM ������� Where �������������="+IntToStr(NumPodr)+" ;";
 String ClientSQL="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������],  TempAspects.[������������ �����������],  TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������],  TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;";
if(Role==2)
{
/*
if(Report1->CPodrazdel->ItemIndex==0)
{
ServerSQL="SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������,  �������.[������������� �����������],  �������.[������������ �����������],  �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.[���� ��������], �������.[������ ��������], �������.[����� ��������] FROM ������� ;";

}
else
{
Report1->Podr->First();
Report1->Podr->MoveBy(Report1->CPodrazdel->ItemIndex-1);
int NumPodr=Report1->Podr->FieldByName("ServerNum")->Value;
ServerSQL="SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������,  �������.[������������� �����������],  �������.[������������ �����������],  �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.[���� ��������], �������.[������ ��������], �������.[����� ��������] FROM ������� Where �������������="+IntToStr(NumPodr)+" ;";

}
*/
//ServerSQL="SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������,  �������.[������������� �����������],  �������.[������������ �����������],  �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.[���� ��������], �������.[������ ��������], �������.[����� ��������] FROM ������� ;";
//ServerSQL="SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������, �������.[������������� �����������], �������.[������������ �����������], �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.[���� ��������], �������.[������ ��������], �������.[����� ��������] FROM (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� WHERE (((ObslOtdel.Login)="+IntToStr(Report1->NumLogin)+"));";
ServerSQL="SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������, �������.[������������� �����������], �������.[������������ �����������], �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.[���� ��������], �������.[������ ��������], �������.[����� ��������] FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.Role)<>4));";


Zast->MClient->ReadTable("���������",ServerSQL, "���������", ClientSQL);
}
else
{

/*
if(Report1->CPodrazdel->ItemIndex==0)
{
ServerSQL="SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������, �������.[������������� �����������], �������.[������������ �����������], �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.[���� ��������], �������.[������ ��������], �������.[����� ��������] FROM (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� WHERE (((ObslOtdel.Login)="+IntToStr(Report1->NumLogin)+"));";
}
else
{

Report1->Podr->Active=false;
Report1->Podr->Connection=Report1->RepBase;
Report1->Podr->CommandText="Select * from Temp�������������";
Report1->Podr->Active=true;

Report1->Podr->First();
Report1->Podr->MoveBy(Report1->CPodrazdel->ItemIndex-1);
int NumPodr=Report1->Podr->FieldByName("ServerNum")->Value;
ServerSQL="SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������,  �������.[������������� �����������],  �������.[������������ �����������],  �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.[���� ��������], �������.[������ ��������], �������.[����� ��������] FROM ������� Where �������������="+IntToStr(NumPodr)+" ;";

}
*/
ServerSQL="SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������, �������.[������������� �����������], �������.[������������ �����������], �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.[���� ��������], �������.[������ ��������], �������.[����� ��������] FROM (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� WHERE (((ObslOtdel.Login)="+IntToStr(Report1->NumLogin)+"));";

Zast->MClient->ReadTable("���������",ServerSQL, "���������_�", ClientSQL);
}
}



}
//---------------------------------------------------------------------------

void __fastcall TZast::CloseMAspExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("CompareMSpecAspects2");
 String ServerSQL="SELECT �������.[����� �������], �������.������������� FROM ������� order by �������.[����� �������];";
 String ClientSQL="SELECT TempAspects.[����� �������], TempAspects.������������� FROM TempAspects;";
Zast->MClient->ReadTable("���������",ServerSQL, "���������", ClientSQL);
}
//---------------------------------------------------------------------------
void TZast::BlockMK(bool B)
{
if(B)
{
Save_Cursor = Screen->Cursor;

Screen->Cursor = crNone;
}
else
{
Screen->Cursor=Save_Cursor;
}

if(Result | B)
{
if(B)
{
//Documents->Memo1->Lines->Add("�������������");
  hDll = LoadLibrary("User32.dll");
  BlockInput = (DWORD __stdcall (*)(bool Status))GetProcAddress(hDll, "BlockInput");

  if(!BlockInput)
  {
    FreeLibrary(hDll);
    return;
  }
}
  Result = BlockInput(B);

if(B)
{
//Documents->Memo1->Lines->Add("B=true");
}
else
{
//Documents->Memo1->Lines->Add("B=false");
}

if(Result)
{
//Documents->Memo1->Lines->Add("Result=true");
}
else
{
//Documents->Memo1->Lines->Add("Result=false");
}

  if(!B | !Result)
  {
  FreeLibrary(hDll);
  Result=false;

//Documents->Memo1->Lines->Add("������������");
  }
}
//Documents->Memo1->Lines->Add("�����");
}
//---------------------------------------------------------------------------
void __fastcall TZast::ContStartReports2Execute(TObject *Sender)
{
 Report1->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TZast::PrepWriteAspUsr_ADMExecute(TObject *Sender)
{
//������� ��������� ������ �� ���������� ������ ������������ � ����������� �� ��������� �������(� TempObslOtdel).

/*
Zast->MClient->Act.ParamComm.clear();
Zast->MClient->Act.ParamComm.push_back("WriteAspectsUsr");
String ClientSQL="SELECT �������.[ServerNum],      �������������.ServerNum,     �������.��������,     �������.[��� ����������],     �������.������������,     �������.�������������,     �������.������,     �������.�����������,     �������.G,     �������.O,     �������.R,     �������.S,     �������.T,     �������.L,     �������.N,     �������.Z,     �������.����������,     �������.[���������� �����������],     �������.[������� �����������],     �������.��������������,     �������.[������������� �����������],     �������.[������������ �����������],     �������.[���������� � ��������],     �������.[������������ ���������� � ��������],     �������.�����������,      �������.[���� ��������],     �������.[������ ��������],     �������.[����� ��������] FROM ������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� ORDER BY �������.[����� �������];";
String ServerSQL="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������], TempAspects.[������������ �����������], TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������], TempAspects.�����������,  TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;";
Zast->MClient->WriteTable("�������_�", ClientSQL, "�������", ServerSQL);
*/
MP<TADOCommand>Comm(this);
Comm->Connection=ADOUsrAspect;
Comm->CommandText="Delete * From TempObslOtdel";
Comm->Execute();

Zast->MClient->Act.ParamComm.clear();
Zast->MClient->Act.ParamComm.push_back("PrepWriteAspUsr_ADM_1");
String ClientSQL="SELECT Login, NumObslOtdel FROM TempObslOtdel;";
String ServerSQL="SELECT Login, NumObslOtdel FROM ObslOtdel Where Login="+IntToStr(Form1->NumLogin)+" Order by NumObslOtdel;";
Zast->MClient->ReadTable( "���������", ServerSQL, "���������_�", ClientSQL);


}
//---------------------------------------------------------------------------

void __fastcall TZast::PrepWriteAspUsr_ADM_1Execute(TObject *Sender)
{
bool Deleting=false;
Form1->DataSetRefresh2->Execute();

MP<TADOCommand>Comm(this);
Comm->Connection=ADOUsrAspect;
Comm->CommandText="UPDATE ������� SET �������.Del = False;";
Comm->Execute();

MP<TADODataSet>LAsp(this);
LAsp->Connection=ADOUsrAspect;
LAsp->CommandText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)=174)); ";
//"SELECT ObslOtdel.NumObslOtdel, ObslOtdel.Login, �������������.ServerNum, Logins.ServerNum FROM ������������� INNER JOIN (Logins INNER JOIN ObslOtdel ON Logins.Num = ObslOtdel.Login) ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel WHERE Logins.ServerNum="+IntToStr(Form1->NumLogin)+" Order by �������������.ServerNum; ";
//"SELECT Login, NumObslOtdel FROM ObslOtdel Where Login="+IntToStr(Form1->NumLogin)+" Order by NumObslOtdel;";
LAsp->Active=true;

MP<TADODataSet>ServOtd(this);
ServOtd->Connection=ADOUsrAspect;
ServOtd->CommandText="SELECT TempObslOtdel.Login, �������������.[����� �������������] FROM TempObslOtdel INNER JOIN ������������� ON TempObslOtdel.NumObslOtdel = �������������.ServerNum WHERE (((TempObslOtdel.Login)="+IntToStr(Form1->NumLogin)+"));";
//"SELECT Login, NumObslOtdel FROM TempObslOtdel Order by NumObslOtdel";
ServOtd->Active=true;

for(LAsp->First();!LAsp->Eof;LAsp->Next())
{
 int NumPodr=LAsp->FieldByName("�������������")->Value;
 if(!ServOtd->Locate("����� �������������",NumPodr, SO))
 {
  //���������
  LAsp->Edit();
  LAsp->FieldByName("Del")->Value=true;
  LAsp->Post();
  Deleting=true;
 }
}
if(Deleting)
{
 Zast->BlockMK(false);
ShowMessage("�� ����� ������ �� ������� ���� �������� ������������� ������������� �� �������������\n ����� �������� �������� ����� �������, ��� ��� �� ����.");
 Zast->BlockMK(true);
}
Comm->CommandText="DELETE �������.* FROM ������� WHERE (((�������.Del)=True));";
Comm->Execute();

Zast->MClient->BlockServer("PrepWriteAspUsr");
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
//������ �� �������� �������� ����������� �� ������ ������ �������������
//���� �� ����������� - �������
//����� ���������� ������� �� ������, �������� ��� ��� ���������� �� ������ ������������ ������� ���� � ��������
//������ �� ������������� ������ � ��� ���� �� �������������


/*
//���������� ���������� ������� � ��������
if(Otd->RecordCount==ServOtd->RecordCount)
{
//���������� ������������� ���������
//����� ��������� ������ �����������
ServOtd->First();
for(Otd->First(); !Otd->Eof; Otd->Next())
{
if(Otd->FieldByName("�������������.ServerNum")->AsInteger==ServOtd->FieldByName("NumObslOtdel")->AsInteger)
{
//���������, ���� ������
ServOtd->Next();
//������� ����, ���� ��������� �������� ������� �����������
//Zast->MClient->StartAction("PrepWriteAspUsr");

//Zast->MClient->StartAction("PrepWriteAspUsr_MSpec_1");


}
else
{
//�� ���������
CorrectPodrazd();
}
}
}
else
{
//���������� �� ���������
CorrectPodrazd();
}
//������ ������ �� ������, ������ ��������� ���������
Zast->MClient->BlockServer("PrepWriteAspUsr");



//����� ������������� ������� ������ �� �������� ������
//����� �������� �������� ������� ������ �� �������� ���������.

//PrepWriteAspUsr_MSpec();
*/
}
//---------------------------------------------------------------------------
void TZast::CorrectPodrazd()
{
//������� ��������� � ������������ ������ ������������� �� ������� � ����������
//������� ������� ��� ��� �� ����������� ����� ������������
 Zast->BlockMK(false);
ShowMessage("�� ����� ������ �� ������� ���� �������� ������������� ������������� �� �������������\n ����� �������� ���� �������, ��� ��� ����� ��� �� ����");
 Zast->BlockMK(true);





/*
if(B)
{
 Zast->BlockMK(false);
ShowMessage("�� ����� ������ �� ������� ���� �������� ������������� ������������� �� �������������\n ���������� ���������� ���������� ������������");
 Zast->BlockMK(true);
 }
 */
/*
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadPodrUSR");
 Zast->MClient->ReadTable("�������", "Select Logins.Num, Logins.Login, Logins.Role From Logins Order by Num;", "�������_�", "Select TempLogins.Num, TempLogins.Login, TempLogins.Role From TempLogins Order by Num;");
*/

/*
Zast->MClient->Act.ParamComm.clear();
//Form1->ReadSprav();

Prog->SignComplete=true;
Prog->Show();
Prog->PB->Min=0;
Prog->PB->Position=0;
Prog->PB->Max=10;

Documents->ReadWrite.clear();
Str_RW S;
S.NameAction="ReadMetodika";
S.Text="������ ��������...";
S.Num=1;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadPodrazd";
S.Text="������ �������������...";
S.Num=2;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadCrit";
S.Text="������ ���������...";
S.Num=3;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadSit";
S.Text="������ ��������...";
S.Num=4;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadVozd1";
S.Text="������ ������ �����������...";
S.Num=5;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadMeropr1";
S.Text="������ ������ �����������...";
S.Num=6;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadTerr1";
S.Text="������ ������ ����������...";
S.Num=7;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadDeyat1";
S.Text="������ ������ ����� ������������...";
S.Num=8;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadAspect1";
S.Text="������ ������ ������������� ��������...";
S.Num=9;
Documents->ReadWrite.push_back(S);

//S.NameAction="PrepWriteAspUsr_MSpec";

/*
S.NameAction="PrepWriteAspUsr";
S.Text="����������� ������ ��������...";
S.Num=10;
Documents->ReadWrite.push_back(S);
*/

/*
Zast->ReadWriteDoc->Execute();
*/

}

//-----------------------------------------------------------------------------
void __fastcall TZast::PrepWriteAspUsr_MSpec_1Execute(TObject *Sender)
{
//���������� ������� ������ ������������� �������� �� ������� � ������ ������������� �������� ���������
//�� ��������� ������ ������������
MP<TADOCommand>Comm(this);
Comm->Connection=ADOUsrAspect;
Comm->CommandText="UPDATE ������� SET �������.Del = False;";
Comm->Execute();

MP<TADODataSet>Local(this);
Local->Connection=ADOUsrAspect;
Local->CommandText="SELECT �������.*, Logins.ServerNum FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+"));";
//"SELECT  * FROM �������";
//"SELECT �������.[ServerNum],      �������������.[ServerNum],     �������.��������,     �������.[��� ����������],     �������.������������,     �������.�������������,     �������.������,     �������.�����������     FROM ������� Where ServerNum<>0 order by ServerNum";
Local->Active=true;

MP<TADODataSet>Server(this);
Server->Connection=ADOUsrAspect;
Server->CommandText="SELECT * FROM TempAspects";
Server->Active=true;

//���������� �������� ���������� � ���������� ������� �������
//���� � ��������� ������ �� ����� ������ - ��������� ������ �������� �� �������� � ����� �������

//���� ��������� ������� ����� ������ - ����� ��������� ������ ������������.
//���� ��������� ����� �� ����� ���� � ��������� ����� �� ��� ������� ���� ��� ������������ ����� ����������� ��� ������

int Deleting=0;
bool SpravError=false;
for(Local->First();!Local->Eof;Local->Next())
{
 int SNum=Local->FieldByName("�������.ServerNum")->AsInteger;
 bool Res=Server->Locate("����� �������", SNum, SO);
 if(Res)
 {
  //�������
  if(Server->FieldByName("��������")->AsInteger==0)
  {
   if(Local->FieldByName("��������")->AsInteger!=0)
   {
   if(Local->FieldByName("�������.ServerNum")->AsInteger!=0)
   {
    //�������� ���� �������

    SpravError=true;

    Local->Edit();
    Local->FieldByName("��������")->Value=0;
    Local->Post();
    }
   }
  }

  if(Server->FieldByName("��� ����������")->AsInteger==0)
  {
   if(Local->FieldByName("��� ����������")->AsInteger!=0)
   {
      if(Local->FieldByName("�������.ServerNum")->AsInteger!=0)
   {
    //��� ���������� ��� ������

    SpravError=true;

    Local->Edit();
    Local->FieldByName("��� ����������")->Value=0;
    Local->Post();
    }
   }
  }

  if(Server->FieldByName("������������")->AsInteger==0)
  {
   if(Local->FieldByName("������������")->AsInteger!=0)
   {
      if(Local->FieldByName("�������.ServerNum")->AsInteger!=0)
   {
    //������������ ���� �������

    SpravError=true;

    Local->Edit();
    Local->FieldByName("������������")->Value=0;
    Local->Post();
    }
   }
  }

  if(Server->FieldByName("������")->AsInteger==0)
  {
   if(Local->FieldByName("������")->AsInteger!=0)
   {
      if(Local->FieldByName("�������.ServerNum")->AsInteger!=0)
   {
    //������ ��� ������

    SpravError=true;

    Local->Edit();
    Local->FieldByName("������")->Value=0;
    Local->Post();
    }
   }
  }

  if(Server->FieldByName("�����������")->AsInteger==0)
  {
   if(Local->FieldByName("�����������")->AsInteger!=0)
   {
      if(Local->FieldByName("�������.ServerNum")->AsInteger!=0)
   {
    //����������� ���� �������

    SpravError=true;


    Local->Edit();
    Local->FieldByName("�����������")->Value=0;
    Local->Post();
    }   
   }
  }




 }
 else
 {
 //���������
 if(Local->FieldByName("�������.ServerNum")->AsInteger!=0)
 {
 Local->Edit();
 Local->FieldByName("Del")->Value=true;
 Local->Post();
 Deleting++;
 }
 }
}

if(Deleting!=0)
{
Comm->CommandText="Delete * from ������� where Del=true";
Comm->Execute();
  Zast->BlockMK(false);
ShowMessage("��-�� ��������� �� ������� ������ �������� � ���� ������� "+IntToStr(Deleting)+" ��������");
  Zast->BlockMK(true);
}
Prog->Hide();
if(SpravError)
{
 //���� ���� ������ �� ������������ ����� �������� ��� ��������� ������ ����������� ������� �� �������
 //���������� �������� �����������. ��������� ���������� ������������ � ��������� ������
  Zast->BlockMK(false);
 ShowMessage("�� ����� ������ �� ������� \n���� ������� ��������� ������ ������������, �������������� � ��������\n���������� �������� �����������, ��������� ������� \n� ����� ������ ������ ��������");
  Zast->BlockMK(true);

//CorrectPodrazd(false);
}
else
{
 //������ �� ����������, ���������� ������
// Zast->MClient->BlockServer("PrepWriteAspUsr");
Form1->Aspects->Active=false;
Form1->Aspects->Active=true;
Form1-> Initialize(1);
MClient->StartAction("PrepWriteAspUsr");
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::PrepWriteAspUsr_MSpecExecute(TObject *Sender)
{
//��������� � ������� TempAspects ��� ����������� ������� ������� ������������
MP<TADOCommand>Comm(this);
Comm->Connection=ADOUsrAspect;
Comm->CommandText="Delete * From TempAspects";
Comm->Execute();

Zast->MClient->Act.ParamComm.clear();
Zast->MClient->Act.ParamComm.push_back("PrepWriteAspUsr_MSpec_1");
String ServerSQL="SELECT  �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.����������� FROM (������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel WHERE (((ObslOtdel.Login)="+IntToStr(Form1->NumLogin)+")) ORDER BY �������.[����� �������];";
//"SELECT �������.[����� �������],      �������������.[����� �������������],     �������.��������,     �������.[��� ����������],     �������.������������,     �������.�������������,     �������.������,     �������.�����������     FROM (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� WHERE (((ObslOtdel.Login)="+IntToStr(Form1->NumLogin)+")) ORDER BY �������.[����� �������];";
String ClientSQL="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.����������� FROM TempAspects;";
Zast->MClient->ReadTable("���������", ServerSQL, "���������_�", ClientSQL );

}
//---------------------------------------------------------------------------

void __fastcall TZast::PostWriteAspectsUsr1Execute(TObject *Sender)
{
Zast->MClient->PostWriteAspectsUsr();
}
//---------------------------------------------------------------------------

void __fastcall TZast::AspQ1Execute(TObject *Sender)
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="DELETE Logins.AdmNum, �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+"));";
Comm->Execute();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeAspectsUserQ");
 String ServerSQL="SELECT �������.[����� �������],     �������.�������������,     �������.��������,     �������.[��� ����������],     �������.������������,     �������.�������������,     �������.������,     �������.�����������,     �������.G,     �������.O,     �������.R,     �������.S,     �������.T,     �������.L,     �������.N,     �������.Z,     �������.����������,     �������.[���������� �����������],     �������.[������� �����������],     �������.��������������,     �������.[������������� �����������],     �������.[������������ �����������],     �������.[���������� � ��������],     �������.[������������ ���������� � ��������],     �������.�����������,      �������.[���� ��������],     �������.[������ ��������],     �������.[����� ��������] FROM (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� WHERE (((ObslOtdel.Login)="+IntToStr(Form1->NumLogin)+"));";
 String ClientSQL="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������], TempAspects.[������������ �����������], TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������], TempAspects.�����������,  TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;";
Zast->MClient->ReadTable("���������", ServerSQL, "���������_�", ClientSQL);
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadTempAspExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("VerIsNew");
 String ServerSQL="SELECT �������.[����� �������],     �������.�������������,     �������.��������,     �������.[��� ����������],     �������.������������,     �������.�������������,     �������.������,     �������.�����������,     �������.G,     �������.O,     �������.R,     �������.S,     �������.T,     �������.L,     �������.N,     �������.Z,     �������.����������,     �������.[���������� �����������],     �������.[������� �����������],     �������.��������������,     �������.[������������� �����������],     �������.[������������ �����������],     �������.[���������� � ��������],     �������.[������������ ���������� � ��������],     �������.�����������,      �������.[���� ��������],     �������.[������ ��������],     �������.[����� ��������] FROM (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� WHERE (((ObslOtdel.Login)="+IntToStr(Form1->NumLogin)+"));";
 String ClientSQL="SELECT CompareAspects.[����� �������], CompareAspects.�������������, CompareAspects.��������, CompareAspects.[��� ����������], CompareAspects.������������, CompareAspects.�������������, CompareAspects.������, CompareAspects.�����������, CompareAspects.G, CompareAspects.O, CompareAspects.R, CompareAspects.S, CompareAspects.T, CompareAspects.L, CompareAspects.N, CompareAspects.Z, CompareAspects.����������, CompareAspects.[���������� �����������], CompareAspects.[������� �����������], CompareAspects.��������������, CompareAspects.[������������� �����������], CompareAspects.[������������ �����������], CompareAspects.[���������� � ��������], CompareAspects.[������������ ���������� � ��������], CompareAspects.�����������,  CompareAspects.[���� ��������], CompareAspects.[������ ��������], CompareAspects.[����� ��������] FROM CompareAspects;";
Zast->MClient->ReadTable("���������", ServerSQL, "���������_�", ClientSQL);


}
//---------------------------------------------------------------------------

void __fastcall TZast::VerIsNewExecute(TObject *Sender)
{
Form1->Label1->Visible=Form1->IsNew();
ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------
