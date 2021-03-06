//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Zastavka.h"
#include <comobj.hpp>
//#include "CodeText.h"
#include <fstream.h>
//#include "Docs.h"
#include "MasterPointer.h"
//#include "Password.h"
#include "Main.h"
#include <FileCtrl.hpp>
#include "PassForm.h"
#include "Diary.h"
#include "Progress.h"
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

MClient->BlockServer("PreViewLogins1");

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
MP<TIniFile>Ini(Path+"Admin.ini");
Server=Ini->ReadString("Server","Server","localhost");
Port=Ini->ReadInteger("Server","Port",2000);

MClient->VDB.clear();
Form1->CBDatabase->Clear();
ServerName=Ini->ReadString("Server","ServerName","");
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
Form1->CBDatabase->Items->Add(Name);
MClient->VDB.push_back(DBI);
}


}
Form1->CBDatabase->ItemIndex=0;

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

void TZast::MergeOtdels()
{
try
{
TLocateOptions SO;

MP<TADOCommand>Comm(this);
Comm->Connection=MClient->Database;
Comm->CommandText="UPDATE ������������� SET �������������.Del = False;";
Comm->Execute();

MP<TADODataSet>Otdels(this);
Otdels->Connection=MClient->Database;
Otdels->CommandText="Select * From ������������� where NumDatabase="+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].NumDatabase)+" Order by [����� �������������]";
Otdels->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=MClient->Database;
Temp->CommandText="Select * From Temp������������� Order by [����� �������������]";
Temp->Active=true;

for(Otdels->First();!Otdels->Eof;Otdels->Next())
{
int SNum=Otdels->FieldByName("ServerNum")->Value;
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
Otdels->Edit();
Otdels->FieldByName("Del")->Value=true;
Otdels->Post();
}
}

//������� ������ �������������
Comm->CommandText="Delete * from ������������� Where Del=true";
Comm->Execute();

//���� � Temp������������� �������� ������ �� ��� ����� ������������� �� �������
//��������� �� � �������������

Comm->CommandText="INSERT INTO ������������� ( ServerNum, [�������� �������������], NumDatabase ) SELECT Temp�������������.[����� �������������], Temp�������������.[�������� �������������], "+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].NumDatabase)+" AS [Database] FROM Temp�������������;";
Comm->Execute();
}
catch(...)
{
  Zast->MClient->Act.WaitCommand=0;
MClient->WriteDiaryEvent("AdminARM ������","������ ����������� �������������"," ������ "+IntToStr(GetLastError()));
}

}
//--------------------------------------------------------

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

void __fastcall TZast::RegForm_Form1Execute(TObject *Sender)
{
MClient->Act.WaitCommand=3;
MClient->Act.NextCommand=3;
String Par="PostRegForm_Form1";
MClient->Act.ParamComm.clear();
MClient->Act.ParamComm.push_back(Par);

MClient->RegForm(this->Name);        
}
//---------------------------------------------------------------------------

void __fastcall TZast::RegForm_Form2Execute(TObject *Sender)
{
MClient->Act.NextCommand=4;
String Par="PostRegForm_Form2";
MClient->Act.ParamComm.clear();
MClient->Act.ParamComm.push_back(Par);

MClient->RegForm(Pass->Name);
}
//---------------------------------------------------------------------------

void __fastcall TZast::PostRegForm_Form1Execute(TObject *Sender)
{
int max=MClient->VForms.size()-1;
MClient->VForms[max]->IDF=StrToInt(MClient->Act.ParamComm[0]);

MClient->StartAction("RegForm_Form2");        
}
//---------------------------------------------------------------------------

void __fastcall TZast::PostRegForm_Form2Execute(TObject *Sender)
{
int max=MClient->VForms.size()-1;
MClient->VForms[max]->IDF=StrToInt(MClient->Act.ParamComm[0]);

MClient->StartAction("PrepareConnectBase");
}
//---------------------------------------------------------------------------

void __fastcall TZast::PrepareConnectBaseExecute(TObject *Sender)
{

MClient->Act.NextCommand=4;
String Par="ConnectBase";
MClient->Act.ParamComm.clear();
String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Admin.ini");
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
MP<TIniFile>Ini(Path+"Admin.ini");
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
//ViewLogins2
MClient->UnBlockServer("ViewLogins2");
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


void __fastcall TZast::UpdateOtdelsExecute(TObject *Sender)
{
MergeOtdels();

  MClient->VTrigger[0].Var++;
  MClient->ActTrigger(0);
}
//---------------------------------------------------------------------------



void __fastcall TZast::UpdateLoginsExecute(TObject *Sender)
{
Form1->UpdateTempLogin();
Form1->Users->ItemIndex=0;
}
//---------------------------------------------------------------------------


void __fastcall TZast::UpdateOtdelsManExecute(TObject *Sender)
{
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=MClient->Database;



Comm->CommandText="Delete * From TempLogins";
Comm->Execute();
//������ ������� �������
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("UpdateLoginsMan");
 Zast->MClient->Act.NextCommand=5;

 Prog->PB->Position++;
 Zast->MClient->ReadTable(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].Name,"Select [Num], [Login], [Code1], [Code2], Role from Logins", "Select [Num], [Login], [Code1], [Code2], Role from TempLogins");
}
catch(...)
{
Zast->BlockMK(false);
}

}
//---------------------------------------------------------------------------

void __fastcall TZast::UpdateLoginsManExecute(TObject *Sender)
{
try
{
//������ ������� �������������
MP<TADOCommand>Comm(this);
Comm->Connection=MClient->Database;
Comm->CommandText="Delete * From TempObslOtdel";
Comm->Execute();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("UpdateObslOtdelMan");
 Zast->MClient->Act.NextCommand=5;

 Prog->PB->Position++;
 Zast->MClient->ReadTable(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].Name,"Select [Login], [NumObslOtdel] from ObslOtdel", "Select [Login], [NumObslOtdel] from TempObslOtdel");
}
catch(...)
{
Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------
void __fastcall TZast::UpdateObslOtdelManExecute(TObject *Sender)
{
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=MClient->Database;

MergeLogins();
MergeOtdels();


MP<TADODataSet>Otdel(this);
Otdel->Connection=MClient->Database;
Otdel->CommandText="Select * from ������������� Where Numdatabase="+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].NumDatabase);
Otdel->Active=true;

MP<TADODataSet>TempObslOtd(this);
TempObslOtd->Connection=MClient->Database;
TempObslOtd->CommandText="Select * From TempObslOtdel";
TempObslOtd->Active=true;

MP<TADODataSet>Login(this);
Login->Connection=MClient->Database;
Login->CommandText="Select * From Logins  Where Numdatabase="+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].NumDatabase);
Login->Active=true;

for(TempObslOtd->First();!TempObslOtd->Eof;TempObslOtd->Next())
{
int Log2;
int Otd2;
int Log1=TempObslOtd->FieldByName("Login")->Value;
if(Login->Locate("ServerNum", Log1, SO))
{
Log2=Login->FieldByName("Num")->Value;
}


int Otd1=TempObslOtd->FieldByName("NumObslOtdel")->Value;
if(Otdel->Locate("ServerNum", Otd1, SO))
{
Otd2=Otdel->FieldByName("����� �������������")->Value;
}


TempObslOtd->Edit();
TempObslOtd->FieldByName("Login")->Value=Log2;
TempObslOtd->FieldByName("NumObslOtdel")->Value=Otd2;
TempObslOtd->Post();
}

Comm->CommandText="Delete * From ObslOtdel Where Numdatabase="+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].NumDatabase);
Comm->Execute();

Comm->CommandText="INSERT INTO ObslOtdel ( Login, NumObslOtdel, Numdatabase ) SELECT TempObslOtdel.Login, TempObslOtdel.NumObslOtdel, "+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].NumDatabase)+" AS [Database] FROM TempObslOtdel;";
Comm->Execute();

Comm->CommandText="Delete * From TempObslOtdel";
Comm->Execute();


 Prog->PB->Position++;
  MClient->VTrigger[0].Var++;
  MClient->ActTrigger(0);
}
catch(...)
{
Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------
void TZast::MergeLogins()
{

try
{
MP<TADOCommand>Comm(this);
Comm->Connection=MClient->Database;
Comm->CommandText="UPDATE Logins SET Logins.Del = False;";
Comm->Execute();


MP<TADODataSet>TempLogins(this);
TempLogins->Connection=MClient->Database;
TempLogins->CommandText="Select * From TempLogins";
TempLogins->Active=true;

MP<TADODataSet>Logins(this);
Logins->Connection=MClient->Database;
Logins->CommandText="Select * From Logins where Numdatabase="+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].NumDatabase)+" Order by Num";
Logins->Active=true;

MP<TADODataSet>TempObslOtdel(this);
TempObslOtdel->Connection=MClient->Database;

//������ �� Logins, ���������� ������ ����������� �������,������ ������ �� TempLogins
//�������� � �������� ��, � ������� ��� ���������� � TempLogins � ������� �� � �����

for(Logins->First();!Logins->Eof;Logins->Next())
{
int N= Logins->FieldByName("ServerNum")->Value;

bool B=TempLogins->Locate("Num",N,SO);
if(B)
{
//���� ����������

Logins->Edit();
Logins->FieldByName("Login")->Value=TempLogins->FieldByName("Login")->Value;
Logins->FieldByName("Code1")->Value=TempLogins->FieldByName("Code1")->Value;
Logins->FieldByName("Code2")->Value=TempLogins->FieldByName("Code2")->Value;
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
Comm->CommandText="Delete * from Logins where Del=true AND Numdatabase="+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].NumDatabase);
Comm->Execute();

Comm->CommandText="INSERT INTO Logins ( ServerNum, Login, Code1, Code2, Role, NumDatabase ) SELECT TempLogins.Num, TempLogins.Login, TempLogins.Code1, TempLogins.Code2, TempLogins.Role, "+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].NumDatabase)+" AS [Database] FROM TempLogins;";
Comm->Execute();
//������ �� TempLogins � ��������� � Logins ����������


//�������� TempLogins
Comm->CommandText="Delete * from TempLogins";
Comm->Execute();

}
catch(...)
{


}
}
//---------------------------------------------------------------

void __fastcall TZast::SaveLoginsExecute(TObject *Sender)
{
try
{
Prog->PB->Position++;
Zast->MClient->Act.ParamComm.clear();
Zast->MClient->Act.ParamComm.push_back("SaveObslOtd");
Zast->MClient->Act.WaitCommand=8;

Zast->MClient->WriteTable(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].Name,"Select [Num], [Login], [Code1], [Code2], [Role], [ServerNum] From Logins Where NumDatabase="+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].NumDatabase), "Select [Num], [Login], [Code1], [Code2], [Role], [ServerNum] From TempLogins");
}
catch(...)
{
Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::SaveObslOtdExecute(TObject *Sender)
{
try
{
Prog->PB->Position++;
Zast->MClient->Act.ParamComm[0]="SaveTempPodr";
Zast->MClient->Act.WaitCommand=8;

Zast->MClient->WriteTable(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].Name,"Select [Login], [NumObslOtdel] From ObslOtdel Where NumDatabase="+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].NumDatabase), "Select [Login], [NumObslOtdel] From TempObslOtdel");
}
catch(...)
{
Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::SendMergeSaveExecute(TObject *Sender)
{
try
{
Prog->PB->Position++;
//�������� ������� �� ������ ����������� ������ �� �������
Zast->MClient->Act.WaitCommand=9;

ClientSocket->Socket->SendText("Command:9;1|"+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].Name.Length())+"#"+Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].Name+"|");
}
catch(...)
{
Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::PrepareSaveLoginsExecute(TObject *Sender)
{
try
{
Prog->PB->Position++;
Zast->MClient->Act.ParamComm.clear();

Zast->MClient->VTrigger.clear();

Trig T;
T.Var=0;
T.Max=MClient->VDB.size();
T.TrueAction="SaveLogins";
T.FalseAction="PostSaveLogins";
Zast->MClient->VTrigger.push_back(T);
MClient->ActTrigger(0);
}
catch(...)
{
Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::PostSaveLoginsExecute(TObject *Sender)
{
MClient->UnBlockServer("PostSaveLogins1");
Zast->BlockMK(false);
}
//---------------------------------------------------------------------------

void __fastcall TZast::SaveTempPodrExecute(TObject *Sender)
{
try
{
Prog->PB->Position++;
Zast->MClient->Act.ParamComm[0]="SendMergeSave";
Zast->MClient->Act.WaitCommand=8;

Zast->MClient->WriteTable(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].Name,"Select [����� �������������], [�������� �������������], [ServerNum] From ������������� Where NumDatabase="+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].NumDatabase), "Select [����� �������������], [�������� �������������], [ServerNum] From Temp�������������");
}
catch(...)
{
Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::LoadNewLoginsExecute(TObject *Sender)
{
// ������ ��������� ������� ������� � ������� ��� ��������� ��������� ������� ����� �������
//������� ���������� ��������� ������� �������
Prog->PB->Position++;
MP<TADOCommand>Comm(this);
Comm->Connection=MClient->Database;
Comm->CommandText="Delete * From TempLogins";
Comm->Execute();

String NameDB=Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].Name;
Zast->MClient->Act.ParamComm[0]="CorrectNewLogins";
MClient->ReadTable(NameDB, "SELECT TempLogins.Num, TempLogins.ServerNum FROM TempLogins", "SELECT TempLogins.Num, TempLogins.ServerNum FROM TempLogins");




}
//---------------------------------------------------------------------------

void __fastcall TZast::CorrectNewLoginsExecute(TObject *Sender)
{
//������ ����� ������� ������� � ����
Prog->PB->Position++;
TLocateOptions SO;

MP<TADODataSet>CLogins(this);
CLogins->Connection=MClient->Database;
CLogins->CommandText="Select  * From Logins where NumDatabase="+IntToStr(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].NumDatabase)+" Order by Num";
CLogins->Active=true;

MP<TADODataSet>TempLogin(this);
TempLogin->Connection=MClient->Database;
TempLogin->CommandText="Select  * From TempLogins Order by Num";
TempLogin->Active=true;

for(TempLogin->First();!TempLogin->Eof;TempLogin->Next())
{
 int N=TempLogin->FieldByName("Num")->Value;
 if(CLogins->Locate("Num",N,SO))
 {
  CLogins->Edit();
  CLogins->FieldByName("ServerNum")->Value=TempLogin->FieldByName("ServerNum")->Value;
  CLogins->Post();
 }
 else
 {
  ShowMessage("������ ��������� ServerNum ��� ����� ������� N="+IntToStr(N));
 }
}
  MClient->VTrigger[0].Var++;
  MClient->ActTrigger(0);
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadOtdelsExecute(TObject *Sender)
{
//������ �������������
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->MClient->Database;
Comm->CommandText="Delete * From Temp�������������";
Comm->Execute();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("UpdateOtdelsMan");
 Zast->MClient->Act.NextCommand=5;

 Prog->PB->Position++;

 Zast->MClient->ReadTable(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].Name,"Select [����� �������������], [�������� �������������] from �������������", "Select [����� �������������], [�������� �������������] from Temp�������������");
}
catch(...)
{
Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::PrepareReadExecute(TObject *Sender)
{
try
{
Zast->MClient->Act.ParamComm.clear();

Zast->MClient->VTrigger.clear();

Trig T;
T.Var=0;
T.Max=MClient->VDB.size();
T.TrueAction="ReadOtdels";
T.FalseAction="PostRead";
Zast->MClient->VTrigger.push_back(T);
Prog->PB->Position++;

MClient->ActTrigger(0);
}
catch(...)
{
Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::PostReadExecute(TObject *Sender)
{
MClient->UnBlockServer("PostRead1");
Zast->BlockMK(false);
}
//---------------------------------------------------------------------------

void __fastcall TZast::PrepareUpdateOtdExecute(TObject *Sender)
{
Zast->MClient->VTrigger.clear();
Trig T;
T.Var=0;
T.Max=MClient->VDB.size();
T.TrueAction="ReadOtdelsAuto";
T.FalseAction="PostUpdateOtd";
Zast->MClient->VTrigger.push_back(T);
MClient->ActTrigger(0);
}
//---------------------------------------------------------------------------

void __fastcall TZast::PostUpdateOtdExecute(TObject *Sender)
{
//PostUpdateOtd1
MClient->UnBlockServer("PostUpdateOtd1");
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadOtdelsAutoExecute(TObject *Sender)
{
//������ �������������
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->MClient->Database;
Comm->CommandText="Delete * From Temp�������������";
Comm->Execute();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("UpdateOtdels");
 Zast->MClient->Act.NextCommand=5;

 Zast->MClient->ReadTable(Zast->MClient->VDB[StrToInt(MClient->VTrigger[0].Var)].Name,"Select [����� �������������], [�������� �������������] from �������������", "Select [����� �������������], [�������� �������������] from Temp�������������");

}

//---------------------------------------------------------------------------

void __fastcall TZast::LoadTypeOpExecute(TObject *Sender)
{
Prog->PB->Position++;
FDiary->PB->Position++;
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("LoadOperation");
 Zast->MClient->Act.NextCommand=5;

 Zast->MClient->ReadTable("Diary","Select Num, NameType from TypeOp order by Num;", "Select Num, NameType from TempTypeOp order by Num;");

}
//---------------------------------------------------------------------------
void __fastcall TZast::LoadOperationExecute(TObject *Sender)
{
Prog->PB->Position++;
FDiary->PB->Position++;
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeType_Op");
 Zast->MClient->Act.NextCommand=5;

 Zast->MClient->ReadTable("Diary","Select Num, Type, NameOperation from Operations order by Num", "Select Num, Type, NameOperation from TempOperations order by Num");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeType_OpExecute(TObject *Sender)
{
Prog->PB->Position++;
FDiary->PB->Position++;
FDiary->MergeTypeOp();

FDiary->MergeOperations();

LoadEvents->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::LoadEventsExecute(TObject *Sender)
{
Prog->PB->Position++;
FDiary->PB->Position++;
Word Y;
Word M;
Word D;

int Year;
int Month;
int Day;
String Dat, Dat1;

DecodeDate(FDiary->NDate->Date, Y, M, D);
Year=(int)Y;
Month=(int)M;
Day=(int)D;
Dat=IntToStr(Month)+"/"+IntToStr(Day)+"/"+IntToStr(Year);

DecodeDate(FDiary->KDate->Date+1, Y, M, D);
Year=(int)Y;
Month=(int)M;
Day=(int)D;
Dat1=IntToStr(Month)+"/"+IntToStr(Day)+"/"+IntToStr(Year);

String ClientSQL;
String ServerSQL;

if(FDiary->EnNDate->Checked & FDiary->EnKDate->Checked)
{
//�������� ���


ClientSQL="SELECT TempEvents.Num,  TempEvents.Date_Time, TempEvents.Comp, TempEvents.Login, TempEvents.Operation, TempEvents.Prim FROM TempEvents;";
ServerSQL="SELECT Events.Num,  Events.Date_Time, Events.Comp, Events.Login, Events.Operation, Events.Prim FROM Events WHERE (((Events.Date_Time)>=#"+Dat+"#) AND ((Events.Date_Time)<=#"+Dat1+"#)) order by Events.Num";

}
else
{
if(FDiary->EnNDate->Checked)
{
//������� ������ ���������
ClientSQL="SELECT TempEvents.Num,  TempEvents.Date_Time, TempEvents.Comp, TempEvents.Login, TempEvents.Operation, TempEvents.Prim FROM TempEvents;";
ServerSQL="SELECT Events.Num,  Events.Date_Time, Events.Comp, Events.Login, Events.Operation, Events.Prim FROM Events WHERE (((Events.Date_Time)>=#"+Dat+"#)) order by Events.Num;";

}
else
{
if(FDiary->EnKDate->Checked)
{
//������� ������ ��������
ClientSQL="SELECT TempEvents.Num,  TempEvents.Date_Time, TempEvents.Comp, TempEvents.Login, TempEvents.Operation, TempEvents.Prim FROM TempEvents;";
ServerSQL="SELECT Events.Num,  Events.Date_Time, Events.Comp, Events.Login, Events.Operation, Events.Prim FROM Events WHERE (((Events.Date_Time)<=#"+Dat1+"#)) order by Events.Num";

}
else
{
//ShowMessage("��������� ���");
//�� ������� ������
ClientSQL="SELECT TempEvents.Num,  TempEvents.Date_Time, TempEvents.Comp, TempEvents.Login, TempEvents.Operation, TempEvents.Prim FROM TempEvents;";

ServerSQL="SELECT Events.Num,  Events.Date_Time, Events.Comp, Events.Login, Events.Operation, Events.Prim FROM Events order by Events.Num";

}
}
}

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeEvents");
 Zast->MClient->Act.NextCommand=5;

 Zast->MClient->ReadTable("Diary", ServerSQL, ClientSQL);

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeEventsExecute(TObject *Sender)
{
Prog->PB->Position++;
FDiary->PB->Position++;
MP<TADOCommand>Comm(this);
Comm->Connection=MClient->Diary;
Comm->CommandText="Delete * From Events";
Comm->Execute();

Comm->CommandText="INSERT INTO Events ( Num, Date_Time, [Comp], Login, Operation, Prim ) SELECT TempEvents.Num, TempEvents.Date_Time, TempEvents.Comp, TempEvents.Login, TempEvents.Operation, TempEvents.Prim FROM TempEvents;";
Comm->Execute();

if(!FDiary->Visible)
{
Prog->Close();
FDiary->ShowModal();
}
else
{
FDiary->Refresh();
}
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

}
else
{

}

if(Result)
{

}
else
{

}

  if(!B | !Result)
  {
  FreeLibrary(hDll);
  Result=false;


  }
}

}
//---------------------------------------------------------------------------
void __fastcall TZast::BlockServerTimer(TObject *Sender)
{
ClientSocket->Socket->SendText("Command:19;1|1#1");
BlockServer->Enabled=false;        
}
//---------------------------------------------------------------------------

void __fastcall TZast::UnBlockServerTimer(TObject *Sender)
{
 Zast->MClient->Act.WaitCommand=20;
 ClientSocket->Socket->SendText("Command:20;1|1#0");
}
//---------------------------------------------------------------------------

void __fastcall TZast::PostRead1Execute(TObject *Sender)
{
Form1->UpdateTempLogin();
Form1->Users->ItemIndex=0;
Form1->UpdateOtdel(0);

Zast->MClient->Act.ParamComm.clear();

Zast->MClient->VTrigger.clear();
 Zast->MClient->Act.ParamComm.clear();
  Zast->MClient->Act.WaitCommand=0;
  Prog->PB->Position++;
  Prog->Close();
ShowMessage("������ ���������");        
}
//---------------------------------------------------------------------------
void TZast::WaitBlockServer(bool Flag)
{
String Mess="������ �����, �������...";
if(Flag)
{
 //�������� ��������
 Prog->Label1->Caption=Mess;


 Prog->Show();
 BlockServer->Enabled=true;
}
else
{
 //�������� ���������
  BlockServer->Enabled=false;
  

}
}
//---------------------------------------------------------------------------

void __fastcall TZast::PostSaveLogins1Execute(TObject *Sender)
{
Zast->MClient->Act.ParamComm.clear();
Zast->MClient->Act.WaitCommand=0;
Prog->PB->Position++;
Zast->Saved=true;
if(!Stop)
{
Prog->Close();
ShowMessage("������ ���������");
}
else
{
Zast->MClient->WriteDiaryEvent("AdminARM","���������� ������ ������ ������","");
Sleep(2000);
 this->Close();
}        
}
//---------------------------------------------------------------------------

void __fastcall TZast::PreViewLogins1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ViewLogins");

 Zast->MClient->ReadTable("�������","Select Login, Code1, Code2 from Logins Where Role=1", "Select Login, Code1, Code2 From TempLogins");

}
//---------------------------------------------------------------------------

void __fastcall TZast::ViewLogins2Execute(TObject *Sender)
{
Pass->ViewLogins();
Pass->Show();
Pass->EdPass->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TZast::PostUpdateOtd1Execute(TObject *Sender)
{
 Form1->Show();

Form1->Users->ItemIndex=0;
Form1->UpdateOtdel(0);

Zast->MClient->WriteDiaryEvent("AdminARM","���������� �������","");
}
//---------------------------------------------------------------------------

