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
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TZast *Zast;

//---------------------------------------------------------------------------
__fastcall TZast::TZast(TComponent* Owner)
        : TForm(Owner)
{
/*
Reg=false;
Start=false;
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
*/
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

/*
Database=new MDBConnector(ExtractFilePath(MainDatabase), ExtractFileName(MainDatabase), this);
Database->SetPatchBackUp("Archive");
Database->OnArchTimer=ArchTimer;

Diary=new MDBConnector(ExtractFilePath(MainDatabase2), ExtractFileName(MainDatabase2), this);
Diary->SetPatchBackUp("Archive");
Diary->OnArchTimer=ArchTimer;
*/
/*
ADOUsrAspect=new MDBConnector(ExtractFilePath(MainDatabase3), ExtractFileName(MainDatabase3), this);
ADOUsrAspect->SetPatchBackUp("Archive");
ADOUsrAspect->OnArchTimer=ArchTimer;
*/

}
//---------------------------------------------------------------------------
void __fastcall TZast::Timer1Timer(TObject *Sender)
{

this->Hide();
//

this->BorderStyle=bsSingle;


Timer2->Enabled=false;
//ShowMessage("������");
if(Start)
{
Timer1->Enabled=false;
Pass->Show();

Timer1->Enabled=false;
}
else
{
Timer1->Interval=1000;
}
}
//---------------------------------------------------------------------------

void __fastcall TZast::Image1Click(TObject *Sender)
{
this->Hide();
//Documents->Show();
Timer1->Enabled=false;
this->BorderStyle=bsSingle;
Timer1->Enabled=false;
Timer1->Interval=100;
Timer1->Enabled=true;
/*
if(Start)
{
Pass->ShowModal();
}
else
{
Start=true;
}
*/
//Pass->Show();
}
//---------------------------------------------------------------------------
void __fastcall TZast::MyException(TObject *Sender,
                                    Exception *E)
{
/*
E->ClassName();
if(String(E->ClassName()) =="EOleException")
{
 ShowMessage("������ ����������� � ����");

}
else
{
// ShowMessage("������!");
}
*/
}




//---------------------------------------------------
void __fastcall TZast::FormCreate(TObject *Sender)
{
//Application->OnException = MyException;


}
//---------------------------------------------------------------------------

void __fastcall TZast::FormShow(TObject *Sender)
{
Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Admin.ini");
Server=Ini->ReadString("Main","Server","localhost");
Port=Ini->ReadInteger("Main","Port",2000);
//MClient=new Client(ClientSocket, ActionManager1, this);

MClient->VDB.clear();
Form1->CBDatabase->Clear();
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
Form1->CBDatabase->Items->Add(Name);
}

MClient->VDB.push_back(DBI);
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



/*
String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"NetAspects.ini");
ServerName=Ini->ReadString("Server","ServerName","");

MClient=new Client();

MClient->Connect(ServerName, "");

MClient->RegForm(this);

MClient->WriteDiaryEvent("NetAspects","������","");

int Num=Ini->ReadInteger("Server","NumServerBase",0);
for(int i=0;i<Num;i++)
{
String S="Base"+IntToStr(i+1);
String ServerBase=Ini->ReadString(S,"ServerDatabase","");
String Name=Ini->ReadString(S,"Name","");

MClient->AddDatabase(ServerBase, Name);

int LicCount=MClient->GetLicenseCount("�������");
Reg=LicCount!=0;

CDBItem DBI;
DBI.Name=Name;
DBI.ServerDB=ServerBase;
DBI.Num=i;

//CBDatabase->Items->Add(Name);

VDB.push_back(DBI);
}
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
/*
Label2->Caption="������������� ���� ������... "+IntToStr(N)+" ���";
this->Repaint();
*/
}
//----------------------------------




void __fastcall TZast::Timer2Timer(TObject *Sender)
{
Timer2->Enabled=false;
/*
String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile> Ini(Path+"NetAspects.ini");
int Days=Ini->ReadInteger("Main","StoreArchive",0);
ADOConn->ClearArchive(Days);
ADOAspect->ClearArchive(Days);
ADOUsrAspect->ClearArchive(Days);

ADOConn->PackDB();
ADOAspect->PackDB();
ADOUsrAspect->PackDB();
if(Days!=-1)
{

ADOConn->Backup("Archive");
ADOAspect->Backup("Archive");
ADOUsrAspect->Backup("Archive");
}

ADOConn->ConnectDB();
ADOAspect->ConnectDB();
ADOUsrAspect->ConnectDB();
*/
MClient->Connect(Server, Port);
Timer1->Enabled=true;
}
//---------------------------------------------------------------------------

bool TZast::LoadLogin(MDBConnector* DB)
{
/*
bool Mess=true;
MP<TADOCommand>Comm(this);
Comm->Connection=DB;

//������� ������ � ��������������
Comm->CommandText="Delete * From Temp�������������";
Comm->Execute();
MP<TADODataSet>Local(this);
Local->Connection=DB;
Local->CommandText="Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������] From Temp������������� Order by [����� �������������]";
Local->Active=true;


//CDBItem Item=VDB[CBDatabase->ItemIndex];
Table* Remote=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Remote->SetCommandText("Select �������������.[����� �������������], �������������.[�������� �������������] From ������������� Order by [����� �������������]");
Remote->Active(true);

MClient->LoadTable(Remote, Local);
//��������� �������� �������� �������������

if(MClient->VerifyTable(Remote, Local)==0)
{

MergeOtdels(DB);



MP<TADODataSet>LTable(this);
LTable->Connection=DB;
LTable->CommandText="Select TempLogins.Num, TempLogins.Login, TempLogins.Role From TempLogins Order by Num";
LTable->Active=true;

//CDBItem Item=VDB[CBDatabase->ItemIndex];
Table* RTable=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);
//
RTable->SetCommandText("Select Logins.Num, Logins.Login, Logins.Role From Logins Order by Num");
RTable->Active(true);


Comm->CommandText="Delete * From TempLogins";
Comm->Execute();

Comm->CommandText="Delete * From TempObslOtdel";
Comm->Execute();
//��������� ������� �������
MClient->LoadTable(RTable, LTable);
//��������� �������� �������� �������
LTable->Active=false;
LTable->Active=true;
if(MClient->VerifyTable(LTable, RTable)==0)
{
//��������� ������� ������������� �������������
MP<TADODataSet>ToTable(this);
ToTable->Connection=DB;
ToTable->CommandText="Select TempObslOtdel.Login, TempObslOtdel.NumObslOtdel From TempObslOtdel Order by Login, NumObslOtdel";
ToTable->Active=true;

//CDBItem Item=VDB[CBDatabase->ItemIndex];
Table* FromTable=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);
//
FromTable->SetCommandText("Select ObslOtdel.Login, ObslOtdel.NumObslOtdel From ObslOtdel Order by Login, NumObslOtdel");
FromTable->Active(true);
MClient->LoadTable(FromTable, ToTable);
//��������� �������� �������� ������� ������������� �������������
if(MClient->VerifyTable(FromTable, ToTable)==0)
{
//���������� ������� �������, ������������ ����������� ������� TempObslOtdel
MergeLogins(DB);

//���������� ������� ������������� �������������
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
}
else
{
Mess=false;
}
MClient->DeleteTable(this, FromTable);

}
else
{
Mess=false;
}
MClient->DeleteTable(this, RTable);

}
else
{
Mess=false;
}
MClient->DeleteTable(this, Remote);
return Mess;
*/
}
//--------------------------------------------------------------------------
void TZast::MergeOtdels()
{
//Main->MClient->WriteDiaryEvent("AdminARM","������ ����������� �������������","");
try
{
TLocateOptions SO;

MP<TADOCommand>Comm(this);
Comm->Connection=MClient->Database;
Comm->CommandText="UPDATE ������������� SET �������������.Del = False;";
Comm->Execute();

MP<TADODataSet>Otdels(this);
Otdels->Connection=MClient->Database;
Otdels->CommandText="Select * From ������������� where NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" Order by [����� �������������]";
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

Comm->CommandText="INSERT INTO ������������� ( ServerNum, [�������� �������������], NumDatabase ) SELECT Temp�������������.[����� �������������], Temp�������������.[�������� �������������], "+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" AS [Database] FROM Temp�������������;";
Comm->Execute();
//Main->MClient->WriteDiaryEvent("AdminARM","����� ����������� �������������","");
}
catch(...)
{
//Main->MClient->WriteDiaryEvent("AdminARM ������","������ ����������� �������������"," ������ "+IntToStr(GetLastError()));

}

}
//--------------------------------------------------------
void TZast::MergeLogins(MDBConnector* DB)
{
/*
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
*/
}
//---------------------------------------------------------------

void __fastcall TZast::FormDestroy(TObject *Sender)
{
delete MClient;
}
//---------------------------------------------------------------------------



void __fastcall TZast::ClientSocketRead(TObject *Sender,
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
/*
MClient->Act.NextCommand=4;
String Par="DiaryConnect";
MClient->Act.ParamComm.clear();
MClient->Act.ParamComm.push_back(Par);

MClient->ConnectDatabase("�������", true);
*/
MClient->Act.NextCommand=4;
String Par="ConnectBase";
MClient->Act.ParamComm.clear();
String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Admin.ini");
int MaxBase=Ini->ReadInteger("Server","NumServerBase",1);
MClient->Act.ParamComm.push_back(IntToStr(MaxBase));
MClient->Act.ParamComm.push_back("1");
MClient->Act.ParamComm.push_back("ConnectBase");
MClient->Act.ParamComm.push_back("LoadLogins");

MClient->StartAction("Trigger");

}
//---------------------------------------------------------------------------

void __fastcall TZast::ConnectBaseExecute(TObject *Sender)
{
/*
MClient->Act.NextCommand=0;
String Par="LoadLogins";
MClient->Act.ParamComm.clear();
MClient->Act.ParamComm.push_back(Par);

MClient->ConnectDatabase("Diary", true);
*/
int TekNumBase=StrToInt(MClient->Act.ParamComm[1]);
//MClient->Act.ParamComm.clear();
String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Admin.ini");
int NumDatabase=Ini->ReadInteger("Base"+IntToStr(TekNumBase),"NumDatabase",-1);
String NameDatabase=Ini->ReadString("Base"+IntToStr(TekNumBase),"Name","");

/////////////////////////////////////////////////
//�������� ����������� ����������� ���
//�� Ini ����� � ���������� ����� ��������
//����� ������� 4
/////////////////////////////////////////////////
MClient->ConnectDatabase(NameDatabase, NumDatabase, true);

MClient->Act.ParamComm[1]=IntToStr(StrToInt(MClient->Act.ParamComm[1])+1);
}
//---------------------------------------------------------------------------

void __fastcall TZast::LoadLoginsExecute(TObject *Sender)
{
Start=true;

}
//---------------------------------------------------------------------------

void __fastcall TZast::ViewLoginsExecute(TObject *Sender)
{
Pass->ViewLogins();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ClientSocketError(TObject *Sender,
      TCustomWinSocket *Socket, TErrorEvent ErrorEvent, int &ErrorCode)
{
if(ErrorCode==10061)
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
/*
ShowMessage(MClient->Act.ParamComm[1]);
ShowMessage(MClient->Act.ParamComm[2]);
*/
MergeOtdels();
 Form1->Show();

Form1->Users->ItemIndex=0;
Form1->UpdateOtdel(0);
}
//---------------------------------------------------------------------------

void __fastcall TZast::TriggerExecute(TObject *Sender)
{
if(StrToInt(MClient->Act.ParamComm[1]-1)<StrToInt(MClient->Act.ParamComm[0]))
{
 MClient->StartAction(MClient->Act.ParamComm[2]);
}
else
{
 MClient->StartAction(MClient->Act.ParamComm[3]);
}
}
//---------------------------------------------------------------------------


void __fastcall TZast::UpdateLoginsExecute(TObject *Sender)
{
Form1->UpdateTempLogin();
Form1->Users->ItemIndex=0;
}
//---------------------------------------------------------------------------
