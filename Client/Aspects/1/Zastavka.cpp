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
//

this->BorderStyle=bsSingle;


Timer2->Enabled=false;
if(Start)
{
Timer1->Enabled=false;
//Pass->Show();
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ViewLogins");

 Zast->MClient->ReadTable("�������","Select Login, Code1, Code2, Role from Logins Where Role<>1", "�������", "Select Login, Code1, Code2, Role From TempLogins");


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
 Zast->MClient->ReadTable("�������", "Select Logins.Num, Logins.Login, Logins.Role From Logins Order by Num;", "�������", "Select TempLogins.Num, TempLogins.Login, TempLogins.Role From TempLogins Order by Num;");

//  Documents->Show();
 break;
 }
 case 3:
 {
  Form1->Login=Login;
   Form1->Caption="������������: "+Login;


 //��� ���� ���� ������� ������� ���������� ������
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadPodrUSR");
 //Zast->MClient->Act.ParamComm.push_back("�������_�");
 Zast->MClient->ReadTable("�������", "Select Logins.Num, Logins.Login, Logins.Role From Logins Order by Num;", "�������_�", "Select TempLogins.Num, TempLogins.Login, TempLogins.Role From TempLogins Order by Num;");




 break;
 }
 case 4:
 {
 //��� ���� ���� ������� ������� ���������� ������

 Form1->Login=Login;
 Form1->Caption="���������������� ������������: "+Login+"                ***������ �� ������ ���������***";
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadPodrUSR");
 //Zast->MClient->Act.ParamComm.push_back("�������_�");
 Zast->MClient->ReadTable("�������", "Select Logins.Num, Logins.Login, Logins.Role From Logins Order by Num;", "�������_�", "Select TempLogins.Num, TempLogins.Login, TempLogins.Role From TempLogins Order by Num;");


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
 String Database="�������";
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
Documents->Metod->Active=false;
Documents->Metod->Active=true;

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� �������� (��������)","");

ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadMetodikaExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeMetodika");
Zast->MClient->ReadTable("Reference", "Select [�����], [��������] From ��������", "Reference", "Select [�����], [��������] From ��������");

}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadWriteDocExecute(TObject *Sender)
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
if(Prog->SignComplete)
{
 ShowMessage("���������");
}
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadPodrazdExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergePodrazd");
 String DBTo;
 if(Role==2)
 {
  DBTo="�������";
 }
 else
 {
  DBTo="�������_�";
 }
Zast->MClient->ReadTable("�������", "Select �������������.[����� �������������], �������������.[�������� �������������] From ������������� order by [����� �������������]", DBTo, "Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������] From Temp������������� order by [����� �������������]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergePodrazdExecute(TObject *Sender)
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





ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadCritExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeCrit");
Zast->MClient->ReadTable("Reference", "Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[��������1], ����������.[��������], ����������.[��� �������], ����������.[���� �������], ����������.[����������� ����] From ���������� Order by [����� ����������];", "Reference", "Select TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.[��������1], TempZn.[��������], TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] From TempZn Order by [����� ����������];");

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



ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadSitExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeSit1");
Zast->MClient->ReadTable("Reference", "Select ��������.[����� ��������], ��������.[�������� ��������] From �������� order by [����� ��������];", "Reference", "Select TempSit.[����� ��������], TempSit.[�������� ��������] From TempSit order by [����� ��������];");

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
MDBConnector* DB;
String DBName;
 if(Role==2)
 {
  DB=ADOAspect;
  DBName="�������";
 }
 else
 {
  DB=ADOUsrAspect;
  DBName="�������_�";
 }

Comm->Connection=DB;
Comm->CommandText="Delete * From TempSit";
Comm->Execute();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeSit2");
Zast->MClient->ReadTable("�������", "Select ��������.[����� ��������], ��������.[�������� ��������] From �������� Where �����=True order by [����� ��������];", DBName, "Select TempSit.[����� ��������], TempSit.[�������� ��������] From TempSit order by [����� ��������];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeSit2Execute(TObject *Sender)
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




ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadVozd1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadVozd2");
Zast->MClient->ReadTable("Reference", "Select ����_3.[����� ����], ����_3.[��������], ����_3.[��������] From ����_3 Order by ��������, [����� ����];", "Reference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadVozd2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeVozd");
Zast->MClient->ReadTable("Reference", "Select �����_3.[����� �����], �����_3.[����� ��������], �����_3.[��������], �����_3.[�����] From �����_3 Order by [����� ��������], [����� �����];", "Reference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeVozdExecute(TObject *Sender)
{
Documents->MergeNode("����_3");
Documents->MergeBranch("�����_3");

Documents->LoadTab1();

String DBName;
if(Role==2)
{
 DBName="�������";
}
else
{
 DBName="�������_�";
}

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeVozd2");
Zast->MClient->ReadTable("�������", "Select �����������.[����� �����������], �����������.[������������ �����������], �����������.[�����] From ����������� order by [����� �����������];", DBName, "Select TempVozd.[����� �����������], TempVozd.[������������ �����������], TempVozd.[�����] From TempVozd order by [����� �����������];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeVozd2Execute(TObject *Sender)
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

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ����������� (��������)","");
}
else
{
Form1->Initialize();
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ����������� (������������)","");
}

ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------


void __fastcall TZast::ReadMeropr1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadMeropr2");
Zast->MClient->ReadTable("Reference", "Select ����_4.[����� ����], ����_4.[��������], ����_4.[��������] From ����_4 Order by ��������, [����� ����];", "Reference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");

}
//---------------------------------------------------------------------------
void __fastcall TZast::ReadMeropr2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeMeropr");
Zast->MClient->ReadTable("Reference", "Select �����_4.[����� �����], �����_4.[����� ��������], �����_4.[��������], �����_4.[�����] From �����_4 Order by [����� ��������], [����� �����];", "Reference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeMeroprExecute(TObject *Sender)
{
//�������� ������� ������ ������ � ��������������� ������� �������� � ���������������� ��������
Documents->MergeNode("����_4");
Documents->MergeBranch("�����_4");
/*
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
Comm->Connection=Zast->ADOConn;
Comm->CommandText="UPDATE �����_4 SET �����_4.Del = True WHERE (((�����_4.�����)=True));";
Comm->Execute();

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOConn;
Temp->CommandText="Select * From �����4";
Temp->Active=true;

MP<TADODataSet>Tab(this);
Tab->Connection=DB;
Tab->CommandText="Select * from ";
 */
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

ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadTerr1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadTerr2");
Zast->MClient->ReadTable("Reference", "Select ����_5.[����� ����], ����_5.[��������], ����_5.[��������] From ����_5 Order by ��������, [����� ����];", "Reference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");


}
//---------------------------------------------------------------------------
void __fastcall TZast::ReadTerr2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeTerr1");
Zast->MClient->ReadTable("Reference", "Select �����_5.[����� �����], �����_5.[����� ��������], �����_5.[��������], �����_5.[�����] From �����_5 Order by [����� ��������], [����� �����];", "Reference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");


}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeTerr1Execute(TObject *Sender)
{
MDBConnector* DB;
String DBName;
if(Role==2)
{
 DB=ADOAspect;
 DBName="�������";
}
else
{
 DB=ADOUsrAspect;
 DBName="�������_�";
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
Zast->MClient->ReadTable("�������", "Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[�����] From ���������� order by [����� ����������];", DBName, "Select TempTerr.[����� ����������], TempTerr.[������������ ����������], TempTerr.[�����] From TempTerr order by [����� ����������];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeTerr2Execute(TObject *Sender)
{
MDBConnector* DB;
String DBName;
if(Role==2)
{
 DB=ADOAspect;
 DBName="�������";
}
else
{
 DB=ADOUsrAspect;
 DBName="�������_�";
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
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ���������� (��������)","");
}
else
{
Form1->Initialize();
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ���������� (������������)","");
}

ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadDeyat1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadDeyat2");
Zast->MClient->ReadTable("Reference", "Select ����_6.[����� ����], ����_6.[��������], ����_6.[��������] From ����_6 Order by ��������, [����� ����];", "Reference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");


}
//---------------------------------------------------------------------------
void __fastcall TZast::ReadDeyat2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeDeyat1");
Zast->MClient->ReadTable("Reference", "Select �����_6.[����� �����], �����_6.[����� ��������], �����_6.[��������], �����_6.[�����] From �����_6 Order by [����� ��������], [����� �����];", "Reference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeDeyat1Execute(TObject *Sender)
{

String DBName;
if(Role==2)
{
 DBName="�������";
}
else
{
 DBName="�������_�";
}

Documents->MergeNode("����_6");
Documents->MergeBranch("�����_6");

if(Role==2)
{
Documents->LoadTab4();
}


 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeDeyat2");
Zast->MClient->ReadTable("�������", "Select ������������.[����� ������������], ������������.[������������ ������������], ������������.[�����] From ������������ order by [����� ������������];", DBName, "Select TempDeyat.[����� ������������], TempDeyat.[������������ ������������], TempDeyat.[�����] From TempDeyat order by [����� ������������];");


}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeDeyat2Execute(TObject *Sender)
{
/*
MDBConnector* DB;
String DBName;
if(Role==2)
{
 DB=ADOAspect;
 DBName="�������";
}
else
{
 DB=ADOUsrAspect;
 DBName="�������_�";
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
*/

String DBName;
if(Role==2)
{

 DBName="�������";
}
else
{

 DBName="�������_�";
}
MP<TADODataSet>Ref(this);
Ref->Connection=Zast->ADOConn;
Ref->CommandText="Select * From �����_6";
Ref->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="UPDATE ������������ SET ������������.Del = False;";
Comm->Execute();

Comm->CommandText="Delete * From TempDeyat";
Comm->Execute();

MP<TADODataSet>TempDeyat(this);
TempDeyat->Connection=Zast->ADOAspect;
TempDeyat->CommandText="Select TempDeyat.[����� ������������], TempDeyat.[������������ ������������], TempDeyat.[�����] From TempDeyat order by [����� ������������]";
TempDeyat->Active=true;

for(Ref->First();!Ref->Eof;Ref->Next())
{
 TempDeyat->Append();
 TempDeyat->FieldByName("����� ������������")->Value=Ref->FieldByName("����� �����")->Value;
 TempDeyat->FieldByName("������������ ������������")->Value=Ref->FieldByName("��������")->Value;
 TempDeyat->FieldByName("�����")->Value=Ref->FieldByName("�����")->Value;
 TempDeyat->Post();
}

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

ReadWriteDoc->Execute();

}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadAspect1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadAspect2");
Zast->MClient->ReadTable("Reference", "Select ����_7.[����� ����], ����_7.[��������], ����_7.[��������] From ����_7 Order by ��������, [����� ����];", "Reference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����];");


}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadAspect2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeAspect1");
Zast->MClient->ReadTable("Reference", "Select �����_7.[����� �����], �����_7.[����� ��������], �����_7.[��������], �����_7.[�����] From �����_7 Order by [����� ��������], [����� �����];", "Reference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����];");


}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeAspect1Execute(TObject *Sender)
{
Documents->MergeNode("����_7");
Documents->MergeBranch("�����_7");
Documents->LoadTab5();



 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeAspect2");
Zast->MClient->ReadTable("�������", "Select ������.[����� �������], ������.[������������ �������], ������.[�����] From ������ order by [����� �������];", "�������", "Select TempAspect.[����� �������], TempAspect.[������������ �������], TempAspect.[�����] From TempAspect order by [����� �������];");


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

ReadWriteDoc->Execute();
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

MClient->Act.ParamComm.clear();
MClient->Act.ParamComm.push_back("ReadWriteDoc");
MClient->WriteTable("Reference","Select �����, �������� From �������� Order by �����", "Reference","Select �����, �������� From �������� Order by �����");

Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ �������� (��������)","");

//ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::WritePodrExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerPodr");
MClient->WriteTable("�������","Select �������������.[����� �������������], �������������.[�������� �������������], �������������.[ServerNum] from ������������� Order by [����� �������������]", "�������", "Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������], Temp�������������.[ServerNum] from Temp������������� Order by [����� �������������]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerPodrExecute(TObject *Sender)
{

 MClient->Act.WaitCommand=11;
 String DB="�������";
 String Mess="Command:11;1|"+IntToStr(DB.Length())+"#"+DB+"|";
 ClientSocket->Socket->SendText(Mess);
}
//---------------------------------------------------------------------------
void __fastcall TZast::WritePodr2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerPodr2");
MClient->ReadTable("�������","Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������], Temp�������������.[ServerNum] from Temp������������� Order by [����� �������������]", "�������", "Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������], Temp�������������.[ServerNum] from Temp������������� Order by [����� �������������]");

}
//---------------------------------------------------------------------------
void __fastcall TZast::MergeServerPodr2Execute(TObject *Sender)
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
/*
Zast->MClient->SetCommandText("�������","Delete * From Temp�������������");
Zast->MClient->CommandExec("�������");

Zast->MClient->WriteDiaryEvent("NetAspects","����� ���������� ������������� (��������)","");
*/
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * from Temp�������������";
Comm->Execute();

 Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ ������������� (��������)","");

 Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------
void __fastcall TZast::WriteCritExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerCrit");
MClient->WriteTable("Reference","Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[��������1], ����������.[��������], ����������.[��� �������], ����������.[���� �������], ����������.[����������� ����] From ���������� Order by [����� ����������]", "Reference", "Select TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.[��������1], TempZn.[��������], TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] From TempZn Order by [����� ����������]");

}
//---------------------------------------------------------------------------
void __fastcall TZast::MergeServerCritExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=12;
 String DB="Reference";
 String DB2="�������";
 String Mess="Command:12;2|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(DB2.Length())+"#"+DB2+"|";
 ClientSocket->Socket->SendText(Mess);
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteSitExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteSit2");
MClient->WriteTable("Reference","Select ��������.[����� ��������], ��������.[�������� ��������] from �������� Order by [����� ��������]", "Reference", "Select TempSit.[����� ��������], TempSit.[�������� ��������] from TempSit Order by [����� ��������]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteSit2Execute(TObject *Sender)
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
MClient->WriteTable("�������","Select ��������.[����� ��������], ��������.[�������� ��������] from �������� Order by [����� ��������]", "�������", "Select TempSit.[����� ��������], TempSit.[�������� ��������] from TempSit Order by [����� ��������]");

}
//---------------------------------------------------------------------------
void __fastcall TZast::MergeServerSitExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=13;
 String DB="Reference";
 String DB2="�������";
 String Mess="Command:13;2|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(DB2.Length())+"#"+DB2+"|";
 ClientSocket->Socket->SendText(Mess);
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteVozd1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteVozd2");
MClient->WriteTable("Reference","Select ����_3.[����� ����], ����_3.[��������], ����_3.[��������] From ����_3 Order by [��������], [����� ����]", "Reference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by [��������], [����� ����]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteVozd2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerVozd");
MClient->WriteTable("Reference","Select �����_3.[����� �����], �����_3.[����� ��������], �����_3.[��������], �����_3.[�����] From �����_3 Order by [����� ��������], [����� �����]", "Reference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerVozdExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=14;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerVozd");

 //"Reference", "����_3","�����_3","�������","�����������","�����������","����� �����������","������������ �����������"
 String DB="Reference";
 String Node="����_3";
 String Branch="�����_3";
 String DB2="�������";
 String Table="�����������";
 String AspField="�����������";
 String KeyTarget="����� �����������";
 String NameTarget="������������ �����������";

 String Mess="Command:14;8|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|"+IntToStr(DB2.Length())+"#"+DB2+"|"+IntToStr(Table.Length())+"#"+Table+"|"+IntToStr(AspField.Length())+"#"+AspField+"|"+IntToStr(KeyTarget.Length())+"#"+KeyTarget+"|"+IntToStr(NameTarget.Length())+"#"+NameTarget+"|";
 ClientSocket->Socket->SendText(Mess);
}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerVozdExecute(TObject *Sender)
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ ����������� (��������)","");
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteMeropr1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteMeropr2");
MClient->WriteTable("Reference","Select ����_4.[����� ����], ����_4.[��������], ����_4.[��������] From ����_4 Order by [��������], [����� ����]", "Reference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by [��������], [����� ����]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteMeropr2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerMeropr");
MClient->WriteTable("Reference","Select �����_4.[����� �����], �����_4.[����� ��������], �����_4.[��������], �����_4.[�����] From �����_4 Order by [����� ��������], [����� �����]", "Reference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerMeroprExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=15;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerMeropr");

 //"Reference", "����_4","�����_4"
 String DB="Reference";
 String Node="����_4";
 String Branch="�����_4";


 String Mess="Command:15;3|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|";
 ClientSocket->Socket->SendText(Mess);

}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerMeroprExecute(TObject *Sender)
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ ����������� (��������)","");
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteTerr1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteTerr2");
MClient->WriteTable("Reference","Select ����_5.[����� ����], ����_5.[��������], ����_5.[��������] From ����_5 Order by [��������], [����� ����]", "Reference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by [��������], [����� ����]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteTerr2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerTerr");
MClient->WriteTable("Reference","Select �����_5.[����� �����], �����_5.[����� ��������], �����_5.[��������], �����_5.[�����] From �����_5 Order by [����� ��������], [����� �����]", "Reference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerTerrExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=14;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerVozd");

 //"Reference", "����_5","�����_5","�������","����������","��� ����������","����� ����������","������������ ����������"
 String DB="Reference";
 String Node="����_5";
 String Branch="�����_5";
 String DB2="�������";
 String Table="����������";
 String AspField="��� ����������";
 String KeyTarget="����� ����������";
 String NameTarget="������������ ����������";

 String Mess="Command:14;8|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|"+IntToStr(DB2.Length())+"#"+DB2+"|"+IntToStr(Table.Length())+"#"+Table+"|"+IntToStr(AspField.Length())+"#"+AspField+"|"+IntToStr(KeyTarget.Length())+"#"+KeyTarget+"|"+IntToStr(NameTarget.Length())+"#"+NameTarget+"|";
 ClientSocket->Socket->SendText(Mess);

}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerTerrExecute(TObject *Sender)
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ ���������� (��������)","");
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteDeyat1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteDeyat2");
MClient->WriteTable("Reference","Select ����_6.[����� ����], ����_6.[��������], ����_6.[��������] From ����_6 Order by [��������], [����� ����]", "Reference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by [��������], [����� ����]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteDeyat2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerDeyat");
MClient->WriteTable("Reference","Select �����_6.[����� �����], �����_6.[����� ��������], �����_6.[��������], �����_6.[�����] From �����_6 Order by [����� ��������], [����� �����]", "Reference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerDeyatExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=14;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerDeyat");

 //"Reference", "����_6","�����_6","�������","������������","������������","����� ������������","������������ ������������"
 String DB="Reference";
 String Node="����_6";
 String Branch="�����_6";
 String DB2="�������";
 String Table="������������";
 String AspField="������������";
 String KeyTarget="����� ������������";
 String NameTarget="������������ ������������";

 String Mess="Command:14;8|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|"+IntToStr(DB2.Length())+"#"+DB2+"|"+IntToStr(Table.Length())+"#"+Table+"|"+IntToStr(AspField.Length())+"#"+AspField+"|"+IntToStr(KeyTarget.Length())+"#"+KeyTarget+"|"+IntToStr(NameTarget.Length())+"#"+NameTarget+"|";
 ClientSocket->Socket->SendText(Mess);

}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerDeyatExecute(TObject *Sender)
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ ����� ������������ (��������)","");
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteAspect1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteAspect2");
MClient->WriteTable("Reference","Select ����_7.[����� ����], ����_7.[��������], ����_7.[��������] From ����_7 Order by [��������], [����� ����]", "Reference", "Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by [��������], [����� ����]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteAspect2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerAspect");
MClient->WriteTable("Reference","Select �����_7.[����� �����], �����_7.[����� ��������], �����_7.[��������], �����_7.[�����] From �����_7 Order by [����� ��������], [����� �����]", "Reference", "Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerAspectExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=14;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerAspect");

 //"Reference", "����_7","�����_7","�������","������","������","����� �������","������������ �������"
 String DB="Reference";
 String Node="����_7";
 String Branch="�����_7";
 String DB2="�������";
 String Table="������";
 String AspField="������";
 String KeyTarget="����� �������";
 String NameTarget="������������ �������";

 String Mess="Command:14;8|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|"+IntToStr(DB2.Length())+"#"+DB2+"|"+IntToStr(Table.Length())+"#"+Table+"|"+IntToStr(AspField.Length())+"#"+AspField+"|"+IntToStr(KeyTarget.Length())+"#"+KeyTarget+"|"+IntToStr(NameTarget.Length())+"#"+NameTarget+"|";
 ClientSocket->Socket->SendText(Mess);

}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerAspectExecute(TObject *Sender)
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ ������ ������������� �������� (��������)","");
}
//---------------------------------------------------------------------------

void __fastcall TZast::ContReportExecute(TObject *Sender)
{
 Report1->CreateRep();
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
Zast->MClient->ReadTable("�������", "Select �������������.[����� �������������], �������������.[�������� �������������] From ������������� Order by [����� �������������];", "�������", "Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������] From Temp������������� Order by [����� �������������];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::StartLoadObslOtdExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartMergeLoginsPodr");
Zast->MClient->ReadTable("�������", "Select ObslOtdel.Login, ObslOtdel.NumObslOtdel From ObslOtdel Order by Login, NumObslOtdel;", "�������", "Select TempObslOtdel.Login, TempObslOtdel.NumObslOtdel From TempObslOtdel Order by Login, NumObslOtdel;");

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


// Form1->Show();
Prog->SignComplete=false;
Prog->Show();
Prog->PB->Min=0;
Prog->PB->Position=0;
Prog->PB->Max=8;

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

S.NameAction="ShowForm1";
S.Num=8;
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
/*
Otdels->Edit();
Otdels->FieldByName("Del")->Value=true;
Otdels->Post();
*/
}
}

//������� ������ �������������
/*
Comm->CommandText="Delete * from ������������� Where Del=true";
Comm->Execute();
*/

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
/*
" SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������, �������.[������������� �����������], �������.[������������ �����������], �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.�����������, �������.[���� ��������], �������.[������ ��������], �������.[����� ��������], �������.ServerNum FROM �������;"
*/
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
  //Tab->FieldList->Fields[i]

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
MAsp->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TZast::SaveAspectsMSpecExecute(TObject *Sender)
{
Zast->MClient->Act.WaitCommand=17;
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("EndsaveAspectsMSpec");

ClientSocket->Socket->SendText("Command:17;0|");

}
//---------------------------------------------------------------------------

void __fastcall TZast::EndsaveAspectsMSpecExecute(TObject *Sender)
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ �������� (��������)","");

Zast->ReadWriteDoc->Execute();
//ShowMessage("������ ������ ���������");
}
//---------------------------------------------------------------------------

void __fastcall TZast::CompareMSpecAspects2Execute(TObject *Sender)
{
//��������� ������� �������� ��� ��������� �� ����� ������� � �������
//��� �������� ������� � ������������� ������ ������� ��������
MP<TADODataSet>Tab(this);
Tab->Connection=ADOAspect;
/*
" SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������, �������.[������������� �����������], �������.[������������ �����������], �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.�����������, �������.[���� ��������], �������.[������ ��������], �������.[����� ��������], �������.ServerNum FROM �������;"
*/
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
  //Tab->FieldList->Fields[i]

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
Documents->ReadWrite.clear();
Str_RW S;
Documents->ReadWrite.push_back(S);


Zast->MClient->Act.ParamComm.clear();
Zast->MClient->Act.ParamComm.push_back("SaveAspectsMSpec2");
Zast->MClient->WriteTable("�������","SELECT �������.[����� �������], �������������.ServerNum FROM ������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� Order by �������.[����� �������]; ", "�������", "SELECT TempAspects.[����� �������], TempAspects.������������� From TempAspects order by [����� �������];");

Zast->ReadWriteDoc->Execute();
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
Zast->MClient->Act.WaitCommand=17;
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("EndsaveAspectsMSpec2");

ClientSocket->Socket->SendText("Command:17;0|");
}
//---------------------------------------------------------------------------

void __fastcall TZast::EndsaveAspectsMSpec2Execute(TObject *Sender)
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ �������� (��������)","");
}
//---------------------------------------------------------------------------

void __fastcall TZast::LoadMSpecAspectsExecute(TObject *Sender)
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

ShowMessage("������ ���������");
}
//---------------------------------------------------------------------------

void __fastcall TZast::SaveAspectsMSpec0Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("SaveAspectsMSpec");
Zast->MClient->WriteTable("�������","SELECT �������.[����� �������], �������������.ServerNum FROM ������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� Order by �������.[����� �������]; ", "�������", "SELECT TempAspects.[����� �������], TempAspects.������������� From TempAspects order by [����� �������];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::StopProgramExecute(TObject *Sender)
{
Zast->Close();
}
//---------------------------------------------------------------------------

void __fastcall TZast::StartLoadPodrUSRExecute(TObject *Sender)
{
//��������� �������� �������������
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadObslOtdUSR");
Zast->MClient->ReadTable("�������", "Select �������������.[����� �������������], �������������.[�������� �������������] From ������������� Order by [����� �������������];", "�������_�", "Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������] From Temp������������� Order by [����� �������������];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::StartLoadObslOtdUSRExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartMergeLoginsPodr");
Zast->MClient->ReadTable("�������", "Select ObslOtdel.Login, ObslOtdel.NumObslOtdel From ObslOtdel Order by Login, NumObslOtdel;", "�������_�", "Select TempObslOtdel.Login, TempObslOtdel.NumObslOtdel From TempObslOtdel Order by Login, NumObslOtdel;");

}
//---------------------------------------------------------------------------

void __fastcall TZast::ShowForm1Execute(TObject *Sender)
{
Form1->Show();
ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeAspectsUserExecute(TObject *Sender)
{
/*
MP<TADODataSet>LPodr(this);
LPodr->Connection=Zast->ADOUsrAspect;
LPodr->CommandText="SELECT �������������.[����� �������������], �������������.[�������� �������������], �������������.ServerNum FROM Logins INNER JOIN (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+"));";
LPodr->Active=true;
*/
MergeAspects(Form1->NumLogin);
}
//---------------------------------------------------------------------------
void  TZast::MergeAspects(int NumLogin)
{
//Zast->MClient->WriteDiaryEvent("NetAspects","���������� ��������","");
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

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * from �������";
Comm->Execute();

String ST="INSERT INTO ������� ( �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, G, O, R, S, T, L, N, Z, ����������, [���������� �����������], [������� �����������], ��������������, [������������� �����������], [������������ �����������], [���������� � ��������], [������������ ���������� � ��������], �����������, [���� ��������], [������ ��������], [����� ��������], ServerNum ) ";
ST=ST+" SELECT TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������], TempAspects.[������������ �����������], TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������], TempAspects.�����������, TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������], TempAspects.[����� �������] ";
ST=ST+" FROM TempAspects; ";
Comm->CommandText=ST;
Comm->Execute();
Form1->Initialize();
ShowMessage("���������");
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ���������� ��������"," ������ "+IntToStr(GetLastError()));

}
}
//----------------------------------------------------------------------
void __fastcall TZast::WriteAspectsUsrExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.WaitCommand=18;
Zast->ClientSocket->Socket->SendText("Command:18;1|"+IntToStr(IntToStr(Form1->NumLogin).Length())+"#"+IntToStr(Form1->NumLogin));
}
//---------------------------------------------------------------------------

