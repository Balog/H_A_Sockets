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

 Zast->MClient->ReadTable("Аспекты","Select Login, Code1, Code2, Role from Logins Where Role<>1", "Аспекты", "Select Login, Code1, Code2, Role From TempLogins");


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
//НАПИСАТЬ ЦИКЛИЧЕСКОЕ ПОДКЛЮЧЕНИЕ БАЗ
//ИЗ Ini ФАЙЛА С ПОЛУЧЕНИЕМ ЧИСЛА ЛИЦЕНЗИЙ
//ЧЕРЕЗ КОМАНДУ 4
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
 ShowMessage("Не найден сервер!");
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
//Необходимо прочесть логины, подразделения, распределение логинов
//а также слить с локальной базой
//и только потом показывать форму документов

//Запрос на загрузку логинов
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadPodr");
 Zast->MClient->ReadTable("Аспекты", "Select Logins.Num, Logins.Login, Logins.Role From Logins Order by Num;", "Аспекты", "Select TempLogins.Num, TempLogins.Login, TempLogins.Role From TempLogins Order by Num;");

//  Documents->Show();
 break;
 }
 case 3:
 {
  Form1->Login=Login;
   Form1->Caption="Пользователь: "+Login;


 //тут тоже надо сначала считать обновления данных
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadPodrUSR");
 //Zast->MClient->Act.ParamComm.push_back("Аспекты_П");
 Zast->MClient->ReadTable("Аспекты", "Select Logins.Num, Logins.Login, Logins.Role From Logins Order by Num;", "Аспекты_П", "Select TempLogins.Num, TempLogins.Login, TempLogins.Role From TempLogins Order by Num;");




 break;
 }
 case 4:
 {
 //тут тоже надо сначала считать обновления данных

 Form1->Login=Login;
 Form1->Caption="Демонстрационный пользователь: "+Login+"                ***ЗАПИСЬ НА СЕРВЕР ОТКЛЮЧЕНА***";
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadPodrUSR");
 //Zast->MClient->Act.ParamComm.push_back("Аспекты_П");
 Zast->MClient->ReadTable("Аспекты", "Select Logins.Num, Logins.Login, Logins.Role From Logins Order by Num;", "Аспекты_П", "Select TempLogins.Num, TempLogins.Login, TempLogins.Role From TempLogins Order by Num;");


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
STab->CommandText="SELECT Аспекты.[Номер аспекта], Аспекты.Подразделение FROM Аспекты WHERE (((Аспекты.Подразделение)="+IntToStr(NumP)+"));";
STab->Active=true;

int N=STab->RecordCount;


if(N==0)
{
//Documents->Podr->Delete();
 Zast->MClient->Act.WaitCommand=16;
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("DeletePodr2");
 String Database="Аспекты";
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
 S1="На сервере зафиксирован "+IntToStr(N)+" аспект, принадлежащий этому подразделению.";
 break;
 }
 case 2: case 3: case 4:
 {
 S1="На сервере зафиксировано "+IntToStr(N)+" аспекта, принадлежащих этому подразделению.";
 break;
 }
 default:
 {
 S1="На сервере зафиксировано "+IntToStr(N)+" аспектов, принадлежащих этому подразделению.";
 }
}

String S=S1+"\rИспользуя \"Движение аспектов\" предварительно освободите удаляемое подразделение от аспектов";
Application->MessageBoxA(S.c_str(),"Удаление подразделения",MB_ICONEXCLAMATION);
}

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeMetodikaExecute(TObject *Sender)
{
Documents->Metod->Active=false;
Documents->Metod->Active=true;

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки методики (главспец)","");

ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadMetodikaExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeMetodika");
Zast->MClient->ReadTable("Reference", "Select [Номер], [Методика] From Методика", "Reference", "Select [Номер], [Методика] From Методика");

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
 ShowMessage("Завершено");
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
  DBTo="Аспекты";
 }
 else
 {
  DBTo="Аспекты_П";
 }
Zast->MClient->ReadTable("Аспекты", "Select Подразделения.[Номер подразделения], Подразделения.[Название подразделения] From Подразделения order by [Номер подразделения]", DBTo, "Select TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения] From TempПодразделения order by [Номер подразделения]");

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
Podr->CommandText="Select * From Подразделения";
Podr->Active=true;

MP<TADODataSet>TempPodr(this);
TempPodr->Connection=DB;
TempPodr->CommandText="Select TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения] From TempПодразделения order by [Номер подразделения]";
TempPodr->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="UPDATE Подразделения SET Подразделения.Del = False;";
Comm->Execute();

for(Podr->First();!Podr->Eof;Podr->Next())
{
 int N=Podr->FieldByName("ServerNum")->Value;
 if(TempPodr->Locate("Номер подразделения",N,SO))
 {
  Podr->Edit();
  Podr->FieldByName("Название подразделения")->Value=TempPodr->FieldByName("Название подразделения")->Value;
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
Comm->CommandText="DELETE * FROM Подразделения WHERE (((Подразделения.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO Подразделения ( ServerNum, [Название подразделения] ) SELECT TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения] FROM TempПодразделения;";
Comm->Execute();

Comm->CommandText="Delete * From TempПодразделения";
Comm->Execute();

if(Role==2)
{
Documents->Podr->Active=false;
Documents->Podr->Active=true;

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки подразделений (главспец)","");
}
else
{
Form1->Initialize();

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки подразделений (пользователь)","");
}





ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadCritExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeCrit");
Zast->MClient->ReadTable("Reference", "Select Значимость.[Номер значимости], Значимость.[Наименование значимости], Значимость.[Критерий1], Значимость.[Критерий], Значимость.[Мин граница], Значимость.[Макс граница], Значимость.[Необходимая мера] From Значимость Order by [Номер значимости];", "Reference", "Select TempZn.[Номер значимости], TempZn.[Наименование значимости], TempZn.[Критерий1], TempZn.[Критерий], TempZn.[Мин граница], TempZn.[Макс граница], TempZn.[Необходимая мера] From TempZn Order by [Номер значимости];");

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
TempZn->CommandText="Select TempZn.[Номер значимости], TempZn.[Наименование значимости], TempZn.[Критерий1], TempZn.[Критерий], TempZn.[Мин граница], TempZn.[Макс граница], TempZn.[Необходимая мера] From TempZn Order by [Номер значимости]";
TempZn->Active=true;

Comm->CommandText="Delete * From Значимость";
Comm->Execute();

Comm->CommandText="INSERT INTO Значимость ( [Номер значимости], [Наименование значимости], Критерий1, Критерий, [Мин граница], [Макс граница], [Необходимая мера] ) SELECT TempZn.[Номер значимости], TempZn.[Наименование значимости], TempZn.Критерий1, TempZn.Критерий, TempZn.[Мин граница], TempZn.[Макс граница], TempZn.[Необходимая мера] FROM TempZn;";
Comm->Execute();
if(Role==2)
{
Documents->ADODataSet1->Active=false;
Documents->ADODataSet1->Active=true;

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки критериев (главспец)","");
}
else
{
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From Значимость";
Comm->Execute();

MP<TADODataSet>From(this);
From->Connection=Zast->ADOConn;
From->CommandText="Select * From Значимость Order by [Номер значимости]";
From->Active=true;

MP<TADODataSet>To(this);
To->Connection=Zast->ADOUsrAspect;
To->CommandText="Select * From Значимость";
To->Active=true;

for(From->First();!From->Eof;From->Next())
{
 To->Append();
 To->FieldByName("Номер значимости")->Value=From->FieldByName("Номер значимости")->Value;
 To->FieldByName("Наименование значимости")->Value=From->FieldByName("Наименование значимости")->Value;
 To->FieldByName("Критерий1")->Value=From->FieldByName("Критерий1")->Value;
 To->FieldByName("Критерий")->Value=From->FieldByName("Критерий")->Value;
 To->FieldByName("Мин граница")->Value=From->FieldByName("Мин граница")->Value;
 To->FieldByName("Макс граница")->Value=From->FieldByName("Макс граница")->Value;
 To->FieldByName("Необходимая мера")->Value=From->FieldByName("Необходимая мера")->Value;
 To->Post();
}

Form1->Initialize();

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки критериев (пользователь)","");
}



ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadSitExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeSit1");
Zast->MClient->ReadTable("Reference", "Select Ситуации.[Номер ситуации], Ситуации.[Название ситуации] From Ситуации order by [Номер ситуации];", "Reference", "Select TempSit.[Номер ситуации], TempSit.[Название ситуации] From TempSit order by [Номер ситуации];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeSit1Execute(TObject *Sender)
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
Comm->CommandText="Delete * From Ситуации";
Comm->Execute();

Comm->CommandText="INSERT INTO Ситуации ( [Номер ситуации], [Название ситуации] ) SELECT TempSit.[Номер ситуации], TempSit.[Название ситуации] FROM TempSit;";
Comm->Execute();

//----
MDBConnector* DB;
String DBName;
 if(Role==2)
 {
  DB=ADOAspect;
  DBName="Аспекты";
 }
 else
 {
  DB=ADOUsrAspect;
  DBName="Аспекты_П";
 }

Comm->Connection=DB;
Comm->CommandText="Delete * From TempSit";
Comm->Execute();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeSit2");
Zast->MClient->ReadTable("Аспекты", "Select Ситуации.[Номер ситуации], Ситуации.[Название ситуации] From Ситуации Where Показ=True order by [Номер ситуации];", DBName, "Select TempSit.[Номер ситуации], TempSit.[Название ситуации] From TempSit order by [Номер ситуации];");

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
Comm->CommandText="UPDATE Ситуации SET Ситуации.Del = True Where Показ=true;";
Comm->Execute();

MP<TADODataSet>Temp(this);
Temp->Connection=DB;
Temp->CommandText="Select * from TempSit";
Temp->Active=true;

MP<TADODataSet>Sit(this);
Sit->Connection=DB;
Sit->CommandText="Select * From Ситуации Where Показ=true";
Sit->Active=true;

for(Sit->First();!Sit->Eof;Sit->Next())
{
int Num=Sit->FieldByName("Номер ситуации")->Value;
 if(Temp->Locate("Номер ситуации", Num, SO))
 {
  Sit->Edit();
  Sit->FieldByName("Название ситуации")->Value=Temp->FieldByName("Название ситуации")->Value;
  Sit->FieldByName("Del")->Value=false;
  Sit->Post();

  Temp->Delete();
 }
 else
 {
  Sit->Edit();
  Sit->FieldByName("Del")->Value=true;
  Sit->Post();

Comm->CommandText="UPDATE Аспекты SET Аспекты.Ситуация = 0 WHERE (((Аспекты.Ситуация)="+IntToStr(Num)+"));";
Comm->Execute();
 }
}
Comm->CommandText="DELETE * From Ситуации Where Del=true AND Показ=true";
Comm->Execute();

Comm->CommandText="INSERT INTO Ситуации ( [Номер ситуации], [Название ситуации], Показ, Del ) SELECT TempSit.[Номер ситуации], TempSit.[Название ситуации], True AS Выражение1, False AS Выражение2 FROM TempSit;";
Comm->Execute();

Comm->CommandText="DELETE * From TempSit";
Comm->Execute();

 if(Role==2)
 {

Documents->Sit->Active=false;
Documents->Sit->Active=true;

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки ситуаций (главспец)","");
 }
 else
 {


Form1->Initialize();

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки ситуаций (пользователь)","");
 }




ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadVozd1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadVozd2");
Zast->MClient->ReadTable("Reference", "Select Узлы_3.[Номер узла], Узлы_3.[Родитель], Узлы_3.[Название] From Узлы_3 Order by Родитель, [Номер узла];", "Reference", "Select TempNode.[Номер узла], TempNode.[Родитель], TempNode.[Название] From TempNode Order by Родитель, [Номер узла];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadVozd2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeVozd");
Zast->MClient->ReadTable("Reference", "Select Ветви_3.[Номер ветви], Ветви_3.[Номер родителя], Ветви_3.[Название], Ветви_3.[Показ] From Ветви_3 Order by [Номер родителя], [Номер ветви];", "Reference", "Select TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.[Название], TempBranch.[Показ] From TempBranch Order by [Номер родителя], [Номер ветви];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeVozdExecute(TObject *Sender)
{
Documents->MergeNode("Узлы_3");
Documents->MergeBranch("Ветви_3");

Documents->LoadTab1();

String DBName;
if(Role==2)
{
 DBName="Аспекты";
}
else
{
 DBName="Аспекты_П";
}

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeVozd2");
Zast->MClient->ReadTable("Аспекты", "Select Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Воздействия.[Показ] From Воздействия order by [Номер воздействия];", DBName, "Select TempVozd.[Номер воздействия], TempVozd.[Наименование воздействия], TempVozd.[Показ] From TempVozd order by [Номер воздействия];");

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
Comm->CommandText="UPDATE Воздействия SET Воздействия.Del = False;";
Comm->Execute();

MP<TADODataSet>TempVozd(this);
TempVozd->Connection=DB;
TempVozd->CommandText="Select TempVozd.[Номер воздействия], TempVozd.[Наименование воздействия], TempVozd.[Показ] From TempVozd order by [Номер воздействия]";
TempVozd->Active=true;

MP<TADODataSet>Vozd(this);
Vozd->Connection=DB;
Vozd->CommandText="Select * From Воздействия";
Vozd->Active=true;

 for(Vozd->First();!Vozd->Eof;Vozd->Next())
 {
  int N=Vozd->FieldByName("Номер воздействия")->Value;

  if(TempVozd->Locate("Номер воздействия", N, SO))
  {
   Vozd->Edit();
   Vozd->FieldByName("Наименование воздействия")->Value=TempVozd->FieldByName("Наименование воздействия")->Value;
   Vozd->FieldByName("Показ")->Value=TempVozd->FieldByName("Показ")->Value;
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
Comm->CommandText="DELETE Воздействия.Del FROM Воздействия WHERE (((Воздействия.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO Воздействия ( [Номер воздействия], [Наименование воздействия], Показ ) SELECT TempVozd.[Номер воздействия], TempVozd.[Наименование воздействия], TempVozd.Показ FROM TempVozd;";
Comm->Execute();

Comm->CommandText="DELETE TempVozd.* FROM TempVozd;";
Comm->Execute();

if(Role==2)
{
Documents->Sit->Active=false;
Documents->Sit->Active=true;

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки воздействий (главспец)","");
}
else
{
Form1->Initialize();
Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки воздействий (пользователь)","");
}

ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------


void __fastcall TZast::ReadMeropr1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadMeropr2");
Zast->MClient->ReadTable("Reference", "Select Узлы_4.[Номер узла], Узлы_4.[Родитель], Узлы_4.[Название] From Узлы_4 Order by Родитель, [Номер узла];", "Reference", "Select TempNode.[Номер узла], TempNode.[Родитель], TempNode.[Название] From TempNode Order by Родитель, [Номер узла];");

}
//---------------------------------------------------------------------------
void __fastcall TZast::ReadMeropr2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeMeropr");
Zast->MClient->ReadTable("Reference", "Select Ветви_4.[Номер ветви], Ветви_4.[Номер родителя], Ветви_4.[Название], Ветви_4.[Показ] From Ветви_4 Order by [Номер родителя], [Номер ветви];", "Reference", "Select TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.[Название], TempBranch.[Показ] From TempBranch Order by [Номер родителя], [Номер ветви];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeMeroprExecute(TObject *Sender)
{
//Добавить перенос данных ветвей в соответствующую таблицу аспектов и пользовательских аспектов
Documents->MergeNode("Узлы_4");
Documents->MergeBranch("Ветви_4");
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
Comm->CommandText="UPDATE Ветви_4 SET Ветви_4.Del = True WHERE (((Ветви_4.Показ)=True));";
Comm->Execute();

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOConn;
Temp->CommandText="Select * From Ветви4";
Temp->Active=true;

MP<TADODataSet>Tab(this);
Tab->Connection=DB;
Tab->CommandText="Select * from ";
 */
if(Role==2)
{
Documents->LoadTab2();

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки мероприятий (главспец)","");
}
else
{
Form1->Initialize();
Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки мероприятий (пользователь)","");
}

ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadTerr1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadTerr2");
Zast->MClient->ReadTable("Reference", "Select Узлы_5.[Номер узла], Узлы_5.[Родитель], Узлы_5.[Название] From Узлы_5 Order by Родитель, [Номер узла];", "Reference", "Select TempNode.[Номер узла], TempNode.[Родитель], TempNode.[Название] From TempNode Order by Родитель, [Номер узла];");


}
//---------------------------------------------------------------------------
void __fastcall TZast::ReadTerr2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeTerr1");
Zast->MClient->ReadTable("Reference", "Select Ветви_5.[Номер ветви], Ветви_5.[Номер родителя], Ветви_5.[Название], Ветви_5.[Показ] From Ветви_5 Order by [Номер родителя], [Номер ветви];", "Reference", "Select TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.[Название], TempBranch.[Показ] From TempBranch Order by [Номер родителя], [Номер ветви];");


}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeTerr1Execute(TObject *Sender)
{
MDBConnector* DB;
String DBName;
if(Role==2)
{
 DB=ADOAspect;
 DBName="Аспекты";
}
else
{
 DB=ADOUsrAspect;
 DBName="Аспекты_П";
}
MP<TADODataSet>Ref(this);
Ref->Connection=Zast->ADOConn;
Ref->CommandText="Select * From Ветви_5";
Ref->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="UPDATE Территория SET Территория.Del = False;";
Comm->Execute();

Comm->CommandText="Delete * From TempTerr";
Comm->Execute();

MP<TADODataSet>TempTerr(this);
TempTerr->Connection=DB;
TempTerr->CommandText="Select TempTerr.[Номер территории], TempTerr.[Наименование территории], TempTerr.[Показ] From TempTerr order by [Номер территории]";
TempTerr->Active=true;

for(Ref->First();!Ref->Eof;Ref->Next())
{
 TempTerr->Append();
 TempTerr->FieldByName("Номер территории")->Value=Ref->FieldByName("Номер ветви")->Value;
 TempTerr->FieldByName("Наименование территории")->Value=Ref->FieldByName("Название")->Value;
 TempTerr->FieldByName("Показ")->Value=Ref->FieldByName("Показ")->Value;
 TempTerr->Post();
}

MP<TADODataSet>Terr(this);
Terr->Connection=DB;
Terr->CommandText="Select * From Территория";
Terr->Active=true;

 for(Terr->First();!Terr->Eof;Terr->Next())
 {
  int N=Terr->FieldByName("Номер территории")->Value;

  if(TempTerr->Locate("Номер территории", N, SO))
  {
   Terr->Edit();
   Terr->FieldByName("Наименование территории")->Value=TempTerr->FieldByName("Наименование территории")->Value;
   Terr->FieldByName("Показ")->Value=TempTerr->FieldByName("Показ")->Value;
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
Comm->CommandText="UPDATE Территория INNER JOIN Аспекты ON Территория.[Номер территории] = Аспекты.[Вид территории] SET Аспекты.[Вид территории] = 0 WHERE (((Территория.Del)=True));";
Comm->Execute();

Comm->CommandText="DELETE Территория.Del FROM Территория WHERE (((Территория.Del)=True) AND Показ=true);";
Comm->Execute();

Comm->CommandText="INSERT INTO Территория ( [Номер территории], [Наименование территории], Показ ) SELECT TempTerr.[Номер территории], TempTerr.[Наименование территории], TempTerr.Показ FROM TempTerr;";
Comm->Execute();

Comm->CommandText="DELETE TempTerr.* FROM TempTerr;";
Comm->Execute();

Documents->MergeNode("Узлы_5");
Documents->MergeBranch("Ветви_5");
Documents->LoadTab3();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeTerr2");
Zast->MClient->ReadTable("Аспекты", "Select Территория.[Номер территории], Территория.[Наименование территории], Территория.[Показ] From Территория order by [Номер территории];", DBName, "Select TempTerr.[Номер территории], TempTerr.[Наименование территории], TempTerr.[Показ] From TempTerr order by [Номер территории];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeTerr2Execute(TObject *Sender)
{
MDBConnector* DB;
String DBName;
if(Role==2)
{
 DB=ADOAspect;
 DBName="Аспекты";
}
else
{
 DB=ADOUsrAspect;
 DBName="Аспекты_П";
}

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="UPDATE Территория SET Территория.Del = False;";
Comm->Execute();

MP<TADODataSet>TempTerr(this);
TempTerr->Connection=DB;
TempTerr->CommandText="Select TempTerr.[Номер территории], TempTerr.[Наименование территории], TempTerr.[Показ] From TempTerr order by [Номер территории]";
TempTerr->Active=true;

MP<TADODataSet>Terr(this);
Terr->Connection=DB;
Terr->CommandText="Select * From Территория";
Terr->Active=true;

 for(Terr->First();!Terr->Eof;Terr->Next())
 {
  int N=Terr->FieldByName("Номер территории")->Value;

  if(TempTerr->Locate("Номер территории", N, SO))
  {
   Terr->Edit();
   Terr->FieldByName("Наименование территории")->Value=TempTerr->FieldByName("Наименование территории")->Value;
   Terr->FieldByName("Показ")->Value=TempTerr->FieldByName("Показ")->Value;
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
Comm->CommandText="DELETE Территория.Del FROM Территория WHERE (((Территория.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO Территория ( [Номер территории], [Наименование территории], Показ ) SELECT TempTerr.[Номер территории], TempTerr.[Наименование территории], TempTerr.Показ FROM TempTerr;";
Comm->Execute();

Comm->CommandText="DELETE TempTerr.* FROM TempTerr;";
Comm->Execute();

if(Role==2)
{
Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки территорий (главспец)","");
}
else
{
Form1->Initialize();
Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки территорий (пользователь)","");
}

ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadDeyat1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadDeyat2");
Zast->MClient->ReadTable("Reference", "Select Узлы_6.[Номер узла], Узлы_6.[Родитель], Узлы_6.[Название] From Узлы_6 Order by Родитель, [Номер узла];", "Reference", "Select TempNode.[Номер узла], TempNode.[Родитель], TempNode.[Название] From TempNode Order by Родитель, [Номер узла];");


}
//---------------------------------------------------------------------------
void __fastcall TZast::ReadDeyat2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeDeyat1");
Zast->MClient->ReadTable("Reference", "Select Ветви_6.[Номер ветви], Ветви_6.[Номер родителя], Ветви_6.[Название], Ветви_6.[Показ] From Ветви_6 Order by [Номер родителя], [Номер ветви];", "Reference", "Select TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.[Название], TempBranch.[Показ] From TempBranch Order by [Номер родителя], [Номер ветви];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeDeyat1Execute(TObject *Sender)
{

String DBName;
if(Role==2)
{
 DBName="Аспекты";
}
else
{
 DBName="Аспекты_П";
}

Documents->MergeNode("Узлы_6");
Documents->MergeBranch("Ветви_6");

if(Role==2)
{
Documents->LoadTab4();
}


 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeDeyat2");
Zast->MClient->ReadTable("Аспекты", "Select Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Деятельность.[Показ] From Деятельность order by [Номер деятельности];", DBName, "Select TempDeyat.[Номер деятельности], TempDeyat.[Наименование деятельности], TempDeyat.[Показ] From TempDeyat order by [Номер деятельности];");


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
 DBName="Аспекты";
}
else
{
 DB=ADOUsrAspect;
 DBName="Аспекты_П";
}
MP<TADODataSet>Ref(this);
Ref->Connection=Zast->ADOConn;
Ref->CommandText="Select * From Ветви_5";
Ref->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="UPDATE Территория SET Территория.Del = False;";
Comm->Execute();

Comm->CommandText="Delete * From TempTerr";
Comm->Execute();

MP<TADODataSet>TempTerr(this);
TempTerr->Connection=DB;
TempTerr->CommandText="Select TempTerr.[Номер территории], TempTerr.[Наименование территории], TempTerr.[Показ] From TempTerr order by [Номер территории]";
TempTerr->Active=true;

for(Ref->First();!Ref->Eof;Ref->Next())
{
 TempTerr->Append();
 TempTerr->FieldByName("Номер территории")->Value=Ref->FieldByName("Номер ветви")->Value;
 TempTerr->FieldByName("Наименование территории")->Value=Ref->FieldByName("Название")->Value;
 TempTerr->FieldByName("Показ")->Value=Ref->FieldByName("Показ")->Value;
 TempTerr->Post();
}
*/

String DBName;
if(Role==2)
{

 DBName="Аспекты";
}
else
{

 DBName="Аспекты_П";
}
MP<TADODataSet>Ref(this);
Ref->Connection=Zast->ADOConn;
Ref->CommandText="Select * From Ветви_6";
Ref->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="UPDATE Деятельность SET Деятельность.Del = False;";
Comm->Execute();

Comm->CommandText="Delete * From TempDeyat";
Comm->Execute();

MP<TADODataSet>TempDeyat(this);
TempDeyat->Connection=Zast->ADOAspect;
TempDeyat->CommandText="Select TempDeyat.[Номер деятельности], TempDeyat.[Наименование деятельности], TempDeyat.[Показ] From TempDeyat order by [Номер деятельности]";
TempDeyat->Active=true;

for(Ref->First();!Ref->Eof;Ref->Next())
{
 TempDeyat->Append();
 TempDeyat->FieldByName("Номер деятельности")->Value=Ref->FieldByName("Номер ветви")->Value;
 TempDeyat->FieldByName("Наименование деятельности")->Value=Ref->FieldByName("Название")->Value;
 TempDeyat->FieldByName("Показ")->Value=Ref->FieldByName("Показ")->Value;
 TempDeyat->Post();
}

MP<TADODataSet>Deyat(this);
Deyat->Connection=Zast->ADOAspect;
Deyat->CommandText="Select * From Деятельность";
Deyat->Active=true;

 for(Deyat->First();!Deyat->Eof;Deyat->Next())
 {
  int N=Deyat->FieldByName("Номер деятельности")->Value;

  if(TempDeyat->Locate("Номер деятельности", N, SO))
  {
   Deyat->Edit();
   Deyat->FieldByName("Наименование деятельности")->Value=TempDeyat->FieldByName("Наименование деятельности")->Value;
   Deyat->FieldByName("Показ")->Value=TempDeyat->FieldByName("Показ")->Value;
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
Comm->CommandText="UPDATE Деятельность INNER JOIN Аспекты ON Деятельность.[Номер деятельности] = Аспекты.[Деятельность] SET Аспекты.[Деятельность] = 0 WHERE (((Деятельность.Del)=True));";
Comm->Execute();

Comm->CommandText="DELETE деятельность.Del FROM Деятельность WHERE (((Деятельность.Del)=True AND Показ=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO Деятельность ( [Номер деятельности], [Наименование деятельности], Показ ) SELECT TempDeyat.[Номер деятельности], TempDeyat.[Наименование деятельности], TempDeyat.Показ FROM TempDeyat;";
Comm->Execute();

Comm->CommandText="DELETE TempDeyat.* FROM TempDeyat;";
Comm->Execute();

if(Role==2)
{
Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки деятельностей (главспец)","");
}
else
{
Form1->Initialize();
Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки деятельностей (пользователь)","");
}

ReadWriteDoc->Execute();

}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadAspect1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ReadAspect2");
Zast->MClient->ReadTable("Reference", "Select Узлы_7.[Номер узла], Узлы_7.[Родитель], Узлы_7.[Название] From Узлы_7 Order by Родитель, [Номер узла];", "Reference", "Select TempNode.[Номер узла], TempNode.[Родитель], TempNode.[Название] From TempNode Order by Родитель, [Номер узла];");


}
//---------------------------------------------------------------------------

void __fastcall TZast::ReadAspect2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeAspect1");
Zast->MClient->ReadTable("Reference", "Select Ветви_7.[Номер ветви], Ветви_7.[Номер родителя], Ветви_7.[Название], Ветви_7.[Показ] From Ветви_7 Order by [Номер родителя], [Номер ветви];", "Reference", "Select TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.[Название], TempBranch.[Показ] From TempBranch Order by [Номер родителя], [Номер ветви];");


}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeAspect1Execute(TObject *Sender)
{
Documents->MergeNode("Узлы_7");
Documents->MergeBranch("Ветви_7");
Documents->LoadTab5();



 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeAspect2");
Zast->MClient->ReadTable("Аспекты", "Select Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Аспект.[Показ] From Аспект order by [Номер аспекта];", "Аспекты", "Select TempAspect.[Номер аспекта], TempAspect.[Наименование аспекта], TempAspect.[Показ] From TempAspect order by [Номер аспекта];");


}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeAspect2Execute(TObject *Sender)
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="UPDATE Аспект SET Аспект.Del = False;";
Comm->Execute();

MP<TADODataSet>TempAspect(this);
TempAspect->Connection=Zast->ADOAspect;
TempAspect->CommandText="Select TempAspect.[Номер аспекта], TempAspect.[Наименование аспекта], TempAspect.[Показ] From TempAspect order by [Номер аспекта]";
TempAspect->Active=true;

MP<TADODataSet>Aspect(this);
Aspect->Connection=Zast->ADOAspect;
Aspect->CommandText="Select * From Аспект";
Aspect->Active=true;

 for(Aspect->First();!Aspect->Eof;Aspect->Next())
 {
  int N=Aspect->FieldByName("Номер аспекта")->Value;

  if(TempAspect->Locate("Номер аспекта", N, SO))
  {
   Aspect->Edit();
   Aspect->FieldByName("Наименование аспекта")->Value=TempAspect->FieldByName("Наименование аспекта")->Value;
   Aspect->FieldByName("Показ")->Value=TempAspect->FieldByName("Показ")->Value;
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
Comm->CommandText="DELETE Аспект.Del FROM Аспект WHERE (((Аспект.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO Аспект ( [Номер аспекта], [Наименование аспекта], Показ ) SELECT TempAspect.[Номер аспекта], TempAspect.[Наименование аспекта], TempAspect.Показ FROM TempAspect;";
Comm->Execute();

Comm->CommandText="DELETE TempDeyat.* FROM TempDeyat;";
Comm->Execute();

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки экологических аспектов (главспец)","");

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
 ShowMessage("Завершено");
}        
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteMetodikaExecute(TObject *Sender)
{

MClient->Act.ParamComm.clear();
MClient->Act.ParamComm.push_back("ReadWriteDoc");
MClient->WriteTable("Reference","Select Номер, Методика From Методика Order by номер", "Reference","Select Номер, Методика From Методика Order by номер");

Zast->MClient->WriteDiaryEvent("NetAspects","Конец записи методики (главспец)","");

//ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TZast::WritePodrExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerPodr");
MClient->WriteTable("Аспекты","Select Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Подразделения.[ServerNum] from Подразделения Order by [Номер подразделения]", "Аспекты", "Select TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения], TempПодразделения.[ServerNum] from TempПодразделения Order by [Номер подразделения]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerPodrExecute(TObject *Sender)
{

 MClient->Act.WaitCommand=11;
 String DB="Аспекты";
 String Mess="Command:11;1|"+IntToStr(DB.Length())+"#"+DB+"|";
 ClientSocket->Socket->SendText(Mess);
}
//---------------------------------------------------------------------------
void __fastcall TZast::WritePodr2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerPodr2");
MClient->ReadTable("Аспекты","Select TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения], TempПодразделения.[ServerNum] from TempПодразделения Order by [Номер подразделения]", "Аспекты", "Select TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения], TempПодразделения.[ServerNum] from TempПодразделения Order by [Номер подразделения]");

}
//---------------------------------------------------------------------------
void __fastcall TZast::MergeServerPodr2Execute(TObject *Sender)
{
MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOAspect;
Temp->CommandText="Select TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения], TempПодразделения.[ServerNum] from TempПодразделения Order by [Номер подразделения]";
Temp->Active=true;

MP<TADODataSet>TempDoc(this);
TempDoc->Connection=Zast->ADOAspect;
TempDoc->CommandText="Select Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Подразделения.[ServerNum] from Подразделения  Where ServerNum=0 Order by [Номер подразделения]";
TempDoc->Active=true;

for(TempDoc->First();!TempDoc->Eof;TempDoc->Next())
{
int N=TempDoc->FieldByName("Номер подразделения")->Value;
if(Temp->Locate("Номер подразделения", N, SO))
{
TempDoc->Edit();
TempDoc->FieldByName("ServerNum")->Value=Temp->FieldByName("ServerNum")->Value;
TempDoc->Post();
}
else
{
ShowMessage("Ошибка записи подразделений Номер="+IntToStr(N));
}
}
/*
Zast->MClient->SetCommandText("Аспекты","Delete * From TempПодразделения");
Zast->MClient->CommandExec("Аспекты");

Zast->MClient->WriteDiaryEvent("NetAspects","Конец сохранения подразделений (главспец)","");
*/
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * from TempПодразделения";
Comm->Execute();

 Zast->MClient->WriteDiaryEvent("NetAspects","Конец записи подразделений (главспец)","");

 Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------
void __fastcall TZast::WriteCritExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerCrit");
MClient->WriteTable("Reference","Select Значимость.[Номер значимости], Значимость.[Наименование значимости], Значимость.[Критерий1], Значимость.[Критерий], Значимость.[Мин граница], Значимость.[Макс граница], Значимость.[Необходимая мера] From Значимость Order by [Номер значимости]", "Reference", "Select TempZn.[Номер значимости], TempZn.[Наименование значимости], TempZn.[Критерий1], TempZn.[Критерий], TempZn.[Мин граница], TempZn.[Макс граница], TempZn.[Необходимая мера] From TempZn Order by [Номер значимости]");

}
//---------------------------------------------------------------------------
void __fastcall TZast::MergeServerCritExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=12;
 String DB="Reference";
 String DB2="Аспекты";
 String Mess="Command:12;2|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(DB2.Length())+"#"+DB2+"|";
 ClientSocket->Socket->SendText(Mess);
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteSitExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteSit2");
MClient->WriteTable("Reference","Select Ситуации.[Номер ситуации], Ситуации.[Название ситуации] from Ситуации Order by [Номер ситуации]", "Reference", "Select TempSit.[Номер ситуации], TempSit.[Название ситуации] from TempSit Order by [Номер ситуации]");

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
RefSit->CommandText="Select Ситуации.[Номер ситуации], Ситуации.[Название ситуации] from Ситуации Order by [Номер ситуации]";
RefSit->Active=true;

MP<TADODataSet>TempSit(this);
TempSit->Connection=ADOAspect;
TempSit->CommandText="Select * from TempSit";
TempSit->Active=true;

for(RefSit->First();!RefSit->Eof;RefSit->Next())
{
 TempSit->Append();
 TempSit->FieldByName("Номер ситуации")->Value=RefSit->FieldByName("Номер ситуации")->Value;
 TempSit->FieldByName("Название ситуации")->Value=RefSit->FieldByName("Название ситуации")->Value;
 TempSit->Post();
}

Comm->CommandText="UPDATE Ситуации SET Ситуации.Del = False;";
Comm->Execute();



MP<TADODataSet>AspSit(this);
AspSit->Connection=ADOAspect;
AspSit->CommandText="Select * from Ситуации Order by [Номер ситуации]";
AspSit->Active=true;

for(AspSit->First();!AspSit->Eof;AspSit->Next())
{
 int Num=AspSit->FieldByName("Номер ситуации")->Value;
 if(TempSit->Locate("Номер ситуации", Num, SO))
 {
  AspSit->Edit();
  AspSit->FieldByName("Название ситуации")->Value=TempSit->FieldByName("Название ситуации")->Value;
  AspSit->FieldByName("Показ")->Value=true;
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

Comm->CommandText="Delete * from Ситуации Where Del=true AND [Показ]=true";
Comm->Execute();

Comm->CommandText="INSERT INTO Ситуации ( [Номер ситуации], [Название ситуации], Показ ) SELECT TempSit.[Номер ситуации], TempSit.[Название ситуации], True AS Выражение1 FROM TempSit;";
Comm->Execute();

Comm->CommandText="UPDATE Ситуации SET Ситуации.Del = False;";
Comm->Execute();

Comm->CommandText="Delete * From TempSit";
Comm->Execute();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerSit");
MClient->WriteTable("Аспекты","Select Ситуации.[Номер ситуации], Ситуации.[Название ситуации] from Ситуации Order by [Номер ситуации]", "Аспекты", "Select TempSit.[Номер ситуации], TempSit.[Название ситуации] from TempSit Order by [Номер ситуации]");

}
//---------------------------------------------------------------------------
void __fastcall TZast::MergeServerSitExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=13;
 String DB="Reference";
 String DB2="Аспекты";
 String Mess="Command:13;2|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(DB2.Length())+"#"+DB2+"|";
 ClientSocket->Socket->SendText(Mess);
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteVozd1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteVozd2");
MClient->WriteTable("Reference","Select Узлы_3.[Номер узла], Узлы_3.[Родитель], Узлы_3.[Название] From Узлы_3 Order by [Родитель], [Номер узла]", "Reference", "Select TempNode.[Номер узла], TempNode.[Родитель], TempNode.[Название] From TempNode Order by [Родитель], [Номер узла]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteVozd2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerVozd");
MClient->WriteTable("Reference","Select Ветви_3.[Номер ветви], Ветви_3.[Номер родителя], Ветви_3.[Название], Ветви_3.[Показ] From Ветви_3 Order by [Номер родителя], [Номер ветви]", "Reference", "Select TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.[Название], TempBranch.[Показ] From TempBranch Order by [Номер родителя], [Номер ветви]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerVozdExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=14;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerVozd");

 //"Reference", "Узлы_3","Ветви_3","Аспекты","Воздействия","Воздействие","Номер воздействия","Наименование воздействия"
 String DB="Reference";
 String Node="Узлы_3";
 String Branch="Ветви_3";
 String DB2="Аспекты";
 String Table="Воздействия";
 String AspField="Воздействие";
 String KeyTarget="Номер воздействия";
 String NameTarget="Наименование воздействия";

 String Mess="Command:14;8|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|"+IntToStr(DB2.Length())+"#"+DB2+"|"+IntToStr(Table.Length())+"#"+Table+"|"+IntToStr(AspField.Length())+"#"+AspField+"|"+IntToStr(KeyTarget.Length())+"#"+KeyTarget+"|"+IntToStr(NameTarget.Length())+"#"+NameTarget+"|";
 ClientSocket->Socket->SendText(Mess);
}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerVozdExecute(TObject *Sender)
{
Zast->MClient->WriteDiaryEvent("NetAspects","Конец записи воздействий (главспец)","");
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteMeropr1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteMeropr2");
MClient->WriteTable("Reference","Select Узлы_4.[Номер узла], Узлы_4.[Родитель], Узлы_4.[Название] From Узлы_4 Order by [Родитель], [Номер узла]", "Reference", "Select TempNode.[Номер узла], TempNode.[Родитель], TempNode.[Название] From TempNode Order by [Родитель], [Номер узла]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteMeropr2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerMeropr");
MClient->WriteTable("Reference","Select Ветви_4.[Номер ветви], Ветви_4.[Номер родителя], Ветви_4.[Название], Ветви_4.[Показ] From Ветви_4 Order by [Номер родителя], [Номер ветви]", "Reference", "Select TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.[Название], TempBranch.[Показ] From TempBranch Order by [Номер родителя], [Номер ветви]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerMeroprExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=15;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerMeropr");

 //"Reference", "Узлы_4","Ветви_4"
 String DB="Reference";
 String Node="Узлы_4";
 String Branch="Ветви_4";


 String Mess="Command:15;3|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|";
 ClientSocket->Socket->SendText(Mess);

}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerMeroprExecute(TObject *Sender)
{
Zast->MClient->WriteDiaryEvent("NetAspects","Конец записи мероприятий (главспец)","");
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteTerr1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteTerr2");
MClient->WriteTable("Reference","Select Узлы_5.[Номер узла], Узлы_5.[Родитель], Узлы_5.[Название] From Узлы_5 Order by [Родитель], [Номер узла]", "Reference", "Select TempNode.[Номер узла], TempNode.[Родитель], TempNode.[Название] From TempNode Order by [Родитель], [Номер узла]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteTerr2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerTerr");
MClient->WriteTable("Reference","Select Ветви_5.[Номер ветви], Ветви_5.[Номер родителя], Ветви_5.[Название], Ветви_5.[Показ] From Ветви_5 Order by [Номер родителя], [Номер ветви]", "Reference", "Select TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.[Название], TempBranch.[Показ] From TempBranch Order by [Номер родителя], [Номер ветви]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerTerrExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=14;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerVozd");

 //"Reference", "Узлы_5","Ветви_5","Аспекты","Территория","Вид территории","Номер территории","Наименование территории"
 String DB="Reference";
 String Node="Узлы_5";
 String Branch="Ветви_5";
 String DB2="Аспекты";
 String Table="Территория";
 String AspField="Вид территории";
 String KeyTarget="Номер территории";
 String NameTarget="Наименование территории";

 String Mess="Command:14;8|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|"+IntToStr(DB2.Length())+"#"+DB2+"|"+IntToStr(Table.Length())+"#"+Table+"|"+IntToStr(AspField.Length())+"#"+AspField+"|"+IntToStr(KeyTarget.Length())+"#"+KeyTarget+"|"+IntToStr(NameTarget.Length())+"#"+NameTarget+"|";
 ClientSocket->Socket->SendText(Mess);

}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerTerrExecute(TObject *Sender)
{
Zast->MClient->WriteDiaryEvent("NetAspects","Конец записи территорий (главспец)","");
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteDeyat1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteDeyat2");
MClient->WriteTable("Reference","Select Узлы_6.[Номер узла], Узлы_6.[Родитель], Узлы_6.[Название] From Узлы_6 Order by [Родитель], [Номер узла]", "Reference", "Select TempNode.[Номер узла], TempNode.[Родитель], TempNode.[Название] From TempNode Order by [Родитель], [Номер узла]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteDeyat2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerDeyat");
MClient->WriteTable("Reference","Select Ветви_6.[Номер ветви], Ветви_6.[Номер родителя], Ветви_6.[Название], Ветви_6.[Показ] From Ветви_6 Order by [Номер родителя], [Номер ветви]", "Reference", "Select TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.[Название], TempBranch.[Показ] From TempBranch Order by [Номер родителя], [Номер ветви]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerDeyatExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=14;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerDeyat");

 //"Reference", "Узлы_6","Ветви_6","Аспекты","Деятельность","Деятельность","Номер деятельности","Наименование деятельности"
 String DB="Reference";
 String Node="Узлы_6";
 String Branch="Ветви_6";
 String DB2="Аспекты";
 String Table="Деятельность";
 String AspField="Деятельность";
 String KeyTarget="Номер деятельности";
 String NameTarget="Наименование деятельности";

 String Mess="Command:14;8|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|"+IntToStr(DB2.Length())+"#"+DB2+"|"+IntToStr(Table.Length())+"#"+Table+"|"+IntToStr(AspField.Length())+"#"+AspField+"|"+IntToStr(KeyTarget.Length())+"#"+KeyTarget+"|"+IntToStr(NameTarget.Length())+"#"+NameTarget+"|";
 ClientSocket->Socket->SendText(Mess);

}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerDeyatExecute(TObject *Sender)
{
Zast->MClient->WriteDiaryEvent("NetAspects","Конец записи видов деятельности (главспец)","");
}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteAspect1Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("WriteAspect2");
MClient->WriteTable("Reference","Select Узлы_7.[Номер узла], Узлы_7.[Родитель], Узлы_7.[Название] From Узлы_7 Order by [Родитель], [Номер узла]", "Reference", "Select TempNode.[Номер узла], TempNode.[Родитель], TempNode.[Название] From TempNode Order by [Родитель], [Номер узла]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::WriteAspect2Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeServerAspect");
MClient->WriteTable("Reference","Select Ветви_7.[Номер ветви], Ветви_7.[Номер родителя], Ветви_7.[Название], Ветви_7.[Показ] From Ветви_7 Order by [Номер родителя], [Номер ветви]", "Reference", "Select TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.[Название], TempBranch.[Показ] From TempBranch Order by [Номер родителя], [Номер ветви]");

}
//---------------------------------------------------------------------------

void __fastcall TZast::MergeServerAspectExecute(TObject *Sender)
{
 MClient->Act.WaitCommand=14;
 MClient->Act.ParamComm.clear();
 MClient->Act.ParamComm.push_back("EndMergeServerAspect");

 //"Reference", "Узлы_7","Ветви_7","Аспекты","Аспект","Аспект","Номер аспекта","Наименование аспекта"
 String DB="Reference";
 String Node="Узлы_7";
 String Branch="Ветви_7";
 String DB2="Аспекты";
 String Table="Аспект";
 String AspField="Аспект";
 String KeyTarget="Номер аспекта";
 String NameTarget="Наименование аспекта";

 String Mess="Command:14;8|"+IntToStr(DB.Length())+"#"+DB+"|"+IntToStr(Node.Length())+"#"+Node+"|"+IntToStr(Branch.Length())+"#"+Branch+"|"+IntToStr(DB2.Length())+"#"+DB2+"|"+IntToStr(Table.Length())+"#"+Table+"|"+IntToStr(AspField.Length())+"#"+AspField+"|"+IntToStr(KeyTarget.Length())+"#"+KeyTarget+"|"+IntToStr(NameTarget.Length())+"#"+NameTarget+"|";
 ClientSocket->Socket->SendText(Mess);

}
//---------------------------------------------------------------------------

void __fastcall TZast::EndMergeServerAspectExecute(TObject *Sender)
{
Zast->MClient->WriteDiaryEvent("NetAspects","Конец записи списка экологических аспектов (главспец)","");
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
//Стартовая загрузка подразделений
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadObslOtd");
Zast->MClient->ReadTable("Аспекты", "Select Подразделения.[Номер подразделения], Подразделения.[Название подразделения] From Подразделения Order by [Номер подразделения];", "Аспекты", "Select TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения] From TempПодразделения Order by [Номер подразделения];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::StartLoadObslOtdExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartMergeLoginsPodr");
Zast->MClient->ReadTable("Аспекты", "Select ObslOtdel.Login, ObslOtdel.NumObslOtdel From ObslOtdel Order by Login, NumObslOtdel;", "Аспекты", "Select TempObslOtdel.Login, TempObslOtdel.NumObslOtdel From TempObslOtdel Order by Login, NumObslOtdel;");

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
//Объединение загруженых с сервера логинов, подразделений и распределения подразделений
//при старте старте программы под главспецом
//Новые не сохраненные подразделения не должны быть удалены

//Объединить таблицу подразделений
MergeOtdels();
//Объединить таблицу логинов, одновременно корректируя таблицу TempObslOtdel
MergeLogins();

//Объединить таблицу обслуживаемых подразделений
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
Podr->CommandText="Select * from Подразделения";
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
  Pod=Podr->FieldByName("Номер подразделения")->Value;

  ToTable->Edit();
  ToTable->FieldByName("Login")->Value=Log;
  ToTable->FieldByName("NumObslOtdel")->Value=Pod;
  ToTable->Post();
 }
 else
 {
  ShowMessage("Ошибка перекодировки подразделений N="+IntToStr(N1));
 }
 }
 else
 {
  ShowMessage("Ошибка перекодировки логинов N="+IntToStr(N));
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
S.Text="Чтение методики...";
S.Num=1;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadPodrazd";
S.Text="Чтение подразделений...";
S.Num=2;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadCrit";
S.Text="Чтение критериев...";
S.Num=3;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadSit";
S.Text="Чтение ситуаций...";
S.Num=4;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadVozd1";
S.Text="Чтение списка воздействий...";
S.Num=5;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadMeropr1";
S.Text="Чтение списка мероприятий...";
S.Num=6;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadTerr1";
S.Text="Чтение списка территорий...";
S.Num=7;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadDeyat1";
S.Text="Чтение списка видов деятельности...";
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
Comm->CommandText="UPDATE Подразделения SET Подразделения.Del = False;";
Comm->Execute();

MP<TADODataSet>Otdels(this);
Otdels->Connection=DB;
Otdels->CommandText="Select * From Подразделения Order by [Номер подразделения]";
Otdels->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=DB;
Temp->CommandText="Select * From TempПодразделения Order by [Номер подразделения]";
Temp->Active=true;

for(Otdels->First();!Otdels->Eof;Otdels->Next())
{
int SNum=Otdels->FieldByName("ServerNum")->AsInteger;
if(Temp->Locate("Номер подразделения", SNum, SO))
{
//найден
//Обновить название подразделения
//удалить запись из TempПодразделения
Otdels->Edit();
Otdels->FieldByName("Название подразделения")->Value=Temp->FieldByName("Название подразделения")->Value;
Otdels->Post();

Temp->Delete();
}
else
{
//Ненайден
//Это подразделение на сервере удалено
//Пометить к удалению
/*
Otdels->Edit();
Otdels->FieldByName("Del")->Value=true;
Otdels->Post();
*/
}
}

//Удалить лишние подразделения
/*
Comm->CommandText="Delete * from Подразделения Where Del=true";
Comm->Execute();
*/

//если в TempПодразделения остались записи то это новые подразделения на сервере
//Перенести их в Подразделения

for(Temp->First();!Temp->Eof;Temp->Next())
{
Otdels->Insert();
Otdels->FieldByName("Название подразделения")->Value=Temp-> FieldByName("Название подразделения")->Value;
Otdels->FieldByName("ServerNum")->Value=Temp-> FieldByName("Номер подразделения")->Value;
Otdels->FieldByName("Del")->Value=false;
Otdels->Post();
}
//очистить TempПодразделения
Comm->CommandText="Delete * from TempПодразделения";
Comm->Execute();


}
//----------------------------------------------------
void __fastcall TZast::CompareMSpecAspectsExecute(TObject *Sender)
{
//Сравнение таблицы аспектов для главспеца по числу записей и составу
//для принятия решения о необходимости чтения таблицы аспектов
MP<TADODataSet>Tab(this);
Tab->Connection=ADOAspect;
/*
" SELECT Аспекты.[Номер аспекта], Аспекты.Подразделение, Аспекты.Ситуация, Аспекты.[Вид территории], Аспекты.Деятельность, Аспекты.Специальность, Аспекты.Аспект, Аспекты.Воздействие, Аспекты.G, Аспекты.O, Аспекты.R, Аспекты.S, Аспекты.T, Аспекты.L, Аспекты.N, Аспекты.Z, Аспекты.Значимость, Аспекты.[Проявление воздействия], Аспекты.[Тяжесть последствий], Аспекты.Приоритетность, Аспекты.[Выполняющиеся мероприятия], Аспекты.[Предлагаемые мероприятия], Аспекты.[Мониторинг и контроль], Аспекты.[Предлагаемый мониторинг и контроль], Аспекты.Исполнитель, Аспекты.[Дата создания], Аспекты.[Начало действия], Аспекты.[Конец действия], Аспекты.ServerNum FROM Аспекты;"
*/
Tab->CommandText="SELECT Аспекты.[Номер аспекта], Аспекты.Ситуация, Аспекты.[Вид территории], Аспекты.Деятельность, Аспекты.Аспект, Аспекты.Воздействие,  Аспекты.Z, Аспекты.Значимость FROM Аспекты ORDER BY Аспекты.[Номер аспекта];";
Tab->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=ADOAspect;
Temp->CommandText="SELECT TempAspects.[Номер аспекта], TempAspects.Ситуация, TempAspects.[Вид территории], TempAspects.Деятельность, TempAspects.Аспект, TempAspects.Воздействие, TempAspects.Z, TempAspects.Значимость FROM TempAspects ORDER BY TempAspects.[Номер аспекта]";
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
    Mess="Содержание аспектов на сервере не совпадает с содержанием аспектов в локальной базе данных\rОбновить список аспектов?";
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
 Mess="Количество аспектов на сервере не совпадает с количеством аспектов в локальной базе данных\rОбновить список аспектов?";
}

if(!Res)
{
 if(Application->MessageBoxA(Mess.c_str(),"Обновление аспектов",MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON1)==IDYES)
 {
  //Обновление списка аспектов для главспеца
MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOAspect;
Temp->CommandText="Select * From TempAspects";
Temp->Active=true;

MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOAspect;
Podr->CommandText="Select * From Подразделения";
Podr->Active=true;

for(Temp->First();!Temp->Eof;Temp->Next())
{
 int N=Temp->FieldByName("Подразделение")->Value;

 if(Podr->Locate("ServerNum", N, SO))
 {
  int Num=Podr->FieldByName("Номер подразделения")->Value;

  Temp->Edit();
  Temp->FieldByName("Подразделение")->Value=Num;
  Temp->Post();
 }
 else
 {
  ShowMessage("Ошибка копирования аспектов");
 }
}

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;



Comm->CommandText="Delete * From Аспекты";
Comm->Execute();


String CT="INSERT INTO Аспекты ( [Номер аспекта], Подразделение, Ситуация, [Вид территории], Деятельность, Специальность, Аспект, Воздействие, G, O, R, S, T, L, N, Z, Значимость, [Проявление воздействия], [Тяжесть последствий], Приоритетность, [Выполняющиеся мероприятия], [Предлагаемые мероприятия], [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Исполнитель, [Дата создания], [Начало действия], [Конец действия] ) ";
CT=CT+" SELECT TempAspects.[Номер аспекта], TempAspects.Подразделение, TempAspects.Ситуация, TempAspects.[Вид территории], TempAspects.Деятельность, TempAspects.Специальность, TempAspects.Аспект, TempAspects.Воздействие, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.Значимость, TempAspects.[Проявление воздействия], TempAspects.[Тяжесть последствий], TempAspects.Приоритетность, TempAspects.[Выполняющиеся мероприятия], TempAspects.[Предлагаемые мероприятия], TempAspects.[Мониторинг и контроль], TempAspects.[Предлагаемый мониторинг и контроль], TempAspects.Исполнитель, TempAspects.[Дата создания], TempAspects.[Начало действия], TempAspects.[Конец действия] ";
CT=CT+" FROM TempAspects;";
Comm->CommandText=CT;
Comm->Execute();

Comm->CommandText="Delete * From TempAspects";
Comm->Execute();
 }
}

MAsp->MoveAspects->Active=false;
String CT="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";
CT=CT+" Order by Аспекты.[Номер аспекта]; ";

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
Zast->MClient->WriteDiaryEvent("NetAspects","Конец записи аспектов (главспец)","");

Zast->ReadWriteDoc->Execute();
//ShowMessage("Запись данных завершена");
}
//---------------------------------------------------------------------------

void __fastcall TZast::CompareMSpecAspects2Execute(TObject *Sender)
{
//Сравнение таблицы аспектов для главспеца по числу записей и составу
//для принятия решения о необходимости записи таблицы аспектов
MP<TADODataSet>Tab(this);
Tab->Connection=ADOAspect;
/*
" SELECT Аспекты.[Номер аспекта], Аспекты.Подразделение, Аспекты.Ситуация, Аспекты.[Вид территории], Аспекты.Деятельность, Аспекты.Специальность, Аспекты.Аспект, Аспекты.Воздействие, Аспекты.G, Аспекты.O, Аспекты.R, Аспекты.S, Аспекты.T, Аспекты.L, Аспекты.N, Аспекты.Z, Аспекты.Значимость, Аспекты.[Проявление воздействия], Аспекты.[Тяжесть последствий], Аспекты.Приоритетность, Аспекты.[Выполняющиеся мероприятия], Аспекты.[Предлагаемые мероприятия], Аспекты.[Мониторинг и контроль], Аспекты.[Предлагаемый мониторинг и контроль], Аспекты.Исполнитель, Аспекты.[Дата создания], Аспекты.[Начало действия], Аспекты.[Конец действия], Аспекты.ServerNum FROM Аспекты;"
*/
Tab->CommandText="SELECT Аспекты.[Номер аспекта], Подразделения.ServerNum FROM Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение ORDER BY Аспекты.[Номер аспекта];";
Tab->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=ADOAspect;
Temp->CommandText="SELECT TempAspects.[Номер аспекта], TempAspects.Подразделение From TempAspects ORDER BY [Номер аспекта]";
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
    Mess="Содержание аспектов на сервере не совпадает с содержанием аспектов в локальной базе данных\rЗаписать список аспектов на сервер?";
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
 Mess="Количество аспектов на сервере не совпадает с количеством аспектов в локальной базе данных\rЗаписать список аспектов на сервер?";
}

if(!Res)
{
 if(Application->MessageBoxA(Mess.c_str(),"Запись аспектов",MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON1)==IDYES)
 {
  //Запись списка аспектов главспецом
Documents->ReadWrite.clear();
Str_RW S;
Documents->ReadWrite.push_back(S);


Zast->MClient->Act.ParamComm.clear();
Zast->MClient->Act.ParamComm.push_back("SaveAspectsMSpec2");
Zast->MClient->WriteTable("Аспекты","SELECT Аспекты.[Номер аспекта], Подразделения.ServerNum FROM Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение Order by Аспекты.[Номер аспекта]; ", "Аспекты", "SELECT TempAspects.[Номер аспекта], TempAspects.Подразделение From TempAspects order by [Номер аспекта];");

Zast->ReadWriteDoc->Execute();
 }
 else
 {
 Zast->MClient->WriteDiaryEvent("NetAspects","Отказ от сохранения движения аспектов (главспец)","");
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
Zast->MClient->WriteDiaryEvent("NetAspects","Конец записи аспектов (главспец)","");
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
Podr->CommandText="Select * From Подразделения";
Podr->Active=true;

for(Temp->First();!Temp->Eof;Temp->Next())
{
 int N=Temp->FieldByName("Подразделение")->Value;

 if(Podr->Locate("ServerNum", N, SO))
 {
  int Num=Podr->FieldByName("Номер подразделения")->Value;

  Temp->Edit();
  Temp->FieldByName("Подразделение")->Value=Num;
  Temp->Post();
 }
 else
 {
  ShowMessage("Ошибка копирования аспектов");
 }
}

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;



Comm->CommandText="Delete * From Аспекты";
Comm->Execute();


String CT="INSERT INTO Аспекты ( [Номер аспекта], Подразделение, Ситуация, [Вид территории], Деятельность, Специальность, Аспект, Воздействие, G, O, R, S, T, L, N, Z, Значимость, [Проявление воздействия], [Тяжесть последствий], Приоритетность, [Выполняющиеся мероприятия], [Предлагаемые мероприятия], [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Исполнитель, [Дата создания], [Начало действия], [Конец действия] ) ";
CT=CT+" SELECT TempAspects.[Номер аспекта], TempAspects.Подразделение, TempAspects.Ситуация, TempAspects.[Вид территории], TempAspects.Деятельность, TempAspects.Специальность, TempAspects.Аспект, TempAspects.Воздействие, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.Значимость, TempAspects.[Проявление воздействия], TempAspects.[Тяжесть последствий], TempAspects.Приоритетность, TempAspects.[Выполняющиеся мероприятия], TempAspects.[Предлагаемые мероприятия], TempAspects.[Мониторинг и контроль], TempAspects.[Предлагаемый мониторинг и контроль], TempAspects.Исполнитель, TempAspects.[Дата создания], TempAspects.[Начало действия], TempAspects.[Конец действия] ";
CT=CT+" FROM TempAspects;";
Comm->CommandText=CT;
Comm->Execute();

Comm->CommandText="Delete * From TempAspects";
Comm->Execute();

MAsp->MoveAspects->Active=false;
CT="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";
CT=CT+" Order by Аспекты.[Номер аспекта]; ";

MAsp->MoveAspects->CommandText=CT;
MAsp->MoveAspects->Connection=Zast->ADOAspect;
MAsp->MoveAspects->Active=true;

MAsp->ChangeCPodr();

ShowMessage("Чтение завершено");
}
//---------------------------------------------------------------------------

void __fastcall TZast::SaveAspectsMSpec0Execute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("SaveAspectsMSpec");
Zast->MClient->WriteTable("Аспекты","SELECT Аспекты.[Номер аспекта], Подразделения.ServerNum FROM Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение Order by Аспекты.[Номер аспекта]; ", "Аспекты", "SELECT TempAspects.[Номер аспекта], TempAspects.Подразделение From TempAspects order by [Номер аспекта];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::StopProgramExecute(TObject *Sender)
{
Zast->Close();
}
//---------------------------------------------------------------------------

void __fastcall TZast::StartLoadPodrUSRExecute(TObject *Sender)
{
//Стартовая загрузка подразделений
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartLoadObslOtdUSR");
Zast->MClient->ReadTable("Аспекты", "Select Подразделения.[Номер подразделения], Подразделения.[Название подразделения] From Подразделения Order by [Номер подразделения];", "Аспекты_П", "Select TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения] From TempПодразделения Order by [Номер подразделения];");

}
//---------------------------------------------------------------------------

void __fastcall TZast::StartLoadObslOtdUSRExecute(TObject *Sender)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("StartMergeLoginsPodr");
Zast->MClient->ReadTable("Аспекты", "Select ObslOtdel.Login, ObslOtdel.NumObslOtdel From ObslOtdel Order by Login, NumObslOtdel;", "Аспекты_П", "Select TempObslOtdel.Login, TempObslOtdel.NumObslOtdel From TempObslOtdel Order by Login, NumObslOtdel;");

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
LPodr->CommandText="SELECT Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Подразделения.ServerNum FROM Logins INNER JOIN (Подразделения INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+"));";
LPodr->Active=true;
*/
MergeAspects(Form1->NumLogin);
}
//---------------------------------------------------------------------------
void  TZast::MergeAspects(int NumLogin)
{
//Zast->MClient->WriteDiaryEvent("NetAspects","Обновление аспектов","");
try
{
MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOUsrAspect;
Podr->CommandText="Select * From Подразделения";
Podr->Active=true;

MP<TADODataSet>TempAsp(this);
TempAsp->Connection=Zast->ADOUsrAspect;
TempAsp->CommandText="Select * From TempAspects";
TempAsp->Active=true;

for(TempAsp->First();!TempAsp->Eof;TempAsp->Next())
{
 int N=TempAsp->FieldByName("Подразделение")->Value;

 if(Podr->Locate("ServerNum",N,SO))
 {
  int Num=Podr->FieldByName("Номер подразделения")->Value;

  TempAsp->Edit();
  TempAsp->FieldByName("Подразделение")->Value=Num;
  TempAsp->Post();
 }
 else
 {
 Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой обновления аспектов","");
  ShowMessage("Ошибка объединения аспектов");
 }
}

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * from Аспекты";
Comm->Execute();

String ST="INSERT INTO Аспекты ( Подразделение, Ситуация, [Вид территории], Деятельность, Специальность, Аспект, Воздействие, G, O, R, S, T, L, N, Z, Значимость, [Проявление воздействия], [Тяжесть последствий], Приоритетность, [Выполняющиеся мероприятия], [Предлагаемые мероприятия], [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Исполнитель, [Дата создания], [Начало действия], [Конец действия], ServerNum ) ";
ST=ST+" SELECT TempAspects.Подразделение, TempAspects.Ситуация, TempAspects.[Вид территории], TempAspects.Деятельность, TempAspects.Специальность, TempAspects.Аспект, TempAspects.Воздействие, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.Значимость, TempAspects.[Проявление воздействия], TempAspects.[Тяжесть последствий], TempAspects.Приоритетность, TempAspects.[Выполняющиеся мероприятия], TempAspects.[Предлагаемые мероприятия], TempAspects.[Мониторинг и контроль], TempAspects.[Предлагаемый мониторинг и контроль], TempAspects.Исполнитель, TempAspects.[Дата создания], TempAspects.[Начало действия], TempAspects.[Конец действия], TempAspects.[Номер аспекта] ";
ST=ST+" FROM TempAspects; ";
Comm->CommandText=ST;
Comm->Execute();
Form1->Initialize();
ShowMessage("Завершено");
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка обновления аспектов"," Ошибка "+IntToStr(GetLastError()));

}
}
//----------------------------------------------------------------------
void __fastcall TZast::WriteAspectsUsrExecute(TObject *Sender)
{
Zast->ClientSocket->Socket->SendText("Command:18;1|"+IntToStr(IntToStr(Form1->NumLogin).Length())+"#"+IntToStr(Form1->NumLogin));
}
//---------------------------------------------------------------------------

