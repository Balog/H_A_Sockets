//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Main.h"
#include "inifiles.hpp";
#include "MasterPointer.h"
#include "CodeText.h"
#include "MDBConnector.h"
#include "EditLogin.h"
#include "Diary.h"

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

Cl=new Clients(this);


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

Cl->ConnectDiary(BI.Patch);
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

UpdateTempLogin();
Users->ItemIndex=0;

String File1=GetFileDatabase("Diary");
/*
int Num1;
for(unsigned int i=0;i<Cl->VBases.size();i++)
{
 if(Cl->VBases[i].FileName==File1)
 {
  Num1=i;
  break;
 }
}
*/


//MyClients=new mClients(Form1, VBases[Num1].Patch);
//MyClients->SetDatabasePatc(ExtractFilePath(VBases[Base->ItemIndex].Patch));
//MyClients->SetDiaryBase(VBases[Num1].Patch);
//MyClients->OnCountClientChanged=ChangeCountClient;

//MyClients->CountClientChanged=ChangeCountClient;

VerifyLicense();



ServerSocket->Active=true;
Cl->DiaryEvent->WriteEvent(Now(),"�� ���������", "�� ��������", "���������", "������ �����������", "");

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

Label1->Caption=IntToStr(Cl->VClients.size());
Cl->DiaryEvent->WriteEvent(Now(),"�� ���������", "�� ��������", "���������", "������ ���������");
this->ChangeCountClient(this);
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
Cl->DiaryEvent->WriteEvent(Now(),Cl->VClients[i]->IP,Cl->VClients[i]->Login , "���������", "������ ��������","����: "+Cl->VClients[i]->AppPatch);
  Cl->VClients.erase(Cl->IVC);
 }
 Cl->IVC++;
}
Label1->Caption=IntToStr(Cl->VClients.size());

  Form1->ListBox1->Clear();
for(unsigned int i=0; i<Cl->VClients.size();i++)
{
String App;
if(ExtractFileName(Cl->VClients[i]->AppPatch)=="AdminARM.exe")
{
 App="�����";
}
if(ExtractFileName(Cl->VClients[i]->AppPatch)=="NetAspects.exe")
{
 App="�������";
}
if(ExtractFileName(Cl->VClients[i]->AppPatch)=="Hazards.exe")
{
 App="���������";
}
 Form1->ListBox1->Items->Add(Cl->VClients[i]->IP+" "+App);


}
this->ChangeCountClient(this);
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
ImageList1->GetIcon(2, Application->Icon);
NID.hIcon=Application->Icon->Handle;
Shell_NotifyIcon(NIM_MODIFY, &NID);

String Mess="";
String Mess1="";
do
{
Mess1=Socket->ReceiveText();
//String Mess1=Socket->ReceiveText();
Mess=Mess+Mess1;
}
while(Mess1.Length()!=0);

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
Parameters.clear();

Shell_NotifyIcon(NIM_DELETE, &NID);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N1Click(TObject *Sender)
{
//�������� �����
LicCount=GetLicenseCount(Base->Text);
if(Role->ItemIndex==2 & Users->Items->Count>=LicCount & LicCount!=-1)
{
 ShowMessage("���������� ������������ ���������� �������������!");
}
else
{
MP<TADODataSet>Tab(this);
Tab->Connection=Form1->Database;
Tab->CommandText="Select * From Logins";
Tab->Active=true;

Tab->Insert();
Tab->FieldByName("Login")->Value=" ";
Tab->FieldByName("Role")->Value=3;
Tab->Post();
Tab->Active=false;
Tab->Active=true;
Tab->Last();
EditLoginNumber=Tab->FieldByName("Num")->Value;

EditLogins->Mode=1;
EditLogins->Login="";
EditLogins->ShowModal();
}
}
//---------------------------------------------------------------------------
long TForm1::GetLicenseCount(String DBName)
{
int Ret=0;

for(unsigned int i=0;i<Cl->VBases.size();i++)
{
 if(Cl->VBases[i].Name==DBName)
 {
  Ret=Cl->VBases[i].LicCount;
  break;
 }
}
return Ret;
}
//-------------------------------------------------------------------------
void TForm1::UpdateTempLogin()
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" order by Login";
Tab->Active=true;

Users->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
 Users->Items->Add(Tab->FieldByName("Login")->AsString);
}
}
//-----------------------------------------------------------------
void TForm1::VerifyLicense()
{
//int Num;
//int LicCount;
MP<TADOConnection>DB(this);
DB->LoginPrompt=false;
for(unsigned int i=0;i<Cl->VBases.size();i++)
{
 if(Cl->VBases[i].MainSpec==0)
 {
  //Num=i;
  LicCount=Cl->VBases[i].LicCount;

  DB->Connected=false;
  DB->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+Cl->VBases[i].Patch+";Persist Security Info=False";
  DB->Connected=true;

if(LicCount>=0)
{
MP<TADODataSet>Logins(this);
Logins->Connection=DB;
Logins->CommandText="Select * From Logins Where Role=3 Order by Num";
Logins->Active=true;
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "��������", "�������� ����� �������������", "DB: "+VBases[i].Name+" ���������="+IntToStr(LicCount)+" �������="+IntToStr(Logins->RecordCount)); //ShowMessage("���� �������� ������");
if(LicCount<Logins->RecordCount)
{
int j=0;
for(Logins->First();!Logins->Eof;Logins->Next())
{
//ShowMessage("i="+IntToStr(i)+" LicCount="+IntToStr(LicCount));
if(j>=LicCount)
{
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "�������� ������", "�������� ���������� �������������", "DB: "+VBases[i].Name+" Login: "+Logins->FieldByName("Login")->AsString); //ShowMessage("���� �������� ������");

//Logins->Delete();
Logins->Edit();
Logins->FieldByName("Del")->Value=true;
Logins->Post();
}
j++;
}

MP<TADOCommand>Comm(this);
Comm->Connection=DB;
Comm->CommandText="Delete * From Logins Where Del=true";
Comm->Execute();
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "�������� ������", "�������� ���������� �������������", "DB: "+VBases[i].Name+" ������� ������������� - "+IntToStr(i-LicCount)); //ShowMessage("���� �������� ������");

}
else
{
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "��������", "�������� ����� �������������", "DB: "+VBases[i].Name+" ����������� ����� �������������"); //ShowMessage("���� �������� ������");

}
}
else
{
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "��������", "�������������� ��������", "DB: "+VBases[i].Name); //ShowMessage("���� �������� ������");

}
 }
}

}
//------------------------------------------------------------------------
void __fastcall TForm1::RoleClick(TObject *Sender)
{
UpdateTempLogin();
Users->ItemIndex=0;
UpdateOtdel(0);
}
//---------------------------------------------------------------------------
void TForm1::UpdateOtdel(int NumLogin)
{
if(NumLogin<0)
{
Otdels->Clear();
}
else
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(NumLogin);
int N1=Tab->FieldByName("Num")->AsInteger;

//int NumLogin=TempLogin->FieldByName("Num")->Value;
MP<TADODataSet>Verify(this);
Verify->Connection=Database;

Verify->Active=false;
Verify->CommandText="SELECT ObslOtdel.Login, �������������.[�������� �������������] FROM ������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel  where Login="+IntToStr(N1)+" Order by NumObslOtdel";
Verify->Active=true;
Otdels->Clear();
for(Verify->First();!Verify->Eof;Verify->Next())
{
Otdels->Items->Add(Verify->FieldByName("�������� �������������")->AsString);
}
}
}
//----------------------------------------------------------------

void __fastcall TForm1::BaseClick(TObject *Sender)
{
int N=Users->ItemIndex;

Database->Connected=false;
Database->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+Cl->VBases[Base->ItemIndex].Patch+";Persist Security Info=False";
Database->LoginPrompt=false;
Database->Connected=true;

UpdateTempLogin();

Users->ItemIndex=N;

UpdateOtdel(N);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button1Click(TObject *Sender)
{
MP<TADODataSet>Tab(this);
Tab->Connection=Form1->Database;
Tab->CommandText="Select * From Logins Where Role="+IntToStr(Role->ItemIndex+1)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(Users->ItemIndex);

String Code=Tab->FieldByName("Code1")->AsString+Tab->FieldByName("Code2")->AsString;
String TabLogin=Tab->FieldByName("Login")->Value;
CodeText* CT=new CodeText();
String Str=CT->EnCrypt(Code, Edit1->Text, 256);
delete CT;
String TLogin=Str.SubString(256-TabLogin.Length()*2+1,TabLogin.Length()*2);
String LG=TLogin.SubString(1,TLogin.Length()/2);
if(LG==TabLogin)
{
ShowMessage("������ �����");
}
else
{
ShowMessage("������ ��������");
}        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N2Click(TObject *Sender)
{
//������������� �����

MP<TADODataSet>Tab(this);
Tab->Connection=Form1->Database;
Tab->CommandText="Select * From Logins Where Role="+IntToStr(Role->ItemIndex+1)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(Users->ItemIndex);

EditLoginNumber=Tab->FieldByName("Num")->Value;

EditLogins->Mode=2;
EditLogins->Login=Tab->FieldByName("Login")->AsString;
EditLogins->ShowModal();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N3Click(TObject *Sender)
{
//������� �����
MP<TADODataSet>Tab(this);
Tab->Connection=Form1->Database;
Tab->CommandText="Select * From Logins Where Role=3 order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(Users->ItemIndex);

Tab->Delete();
UpdateTempLogin();
Otdels->Clear();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N4Click(TObject *Sender)
{
//�������� ������
MP<TADODataSet>Tab(this);
Tab->Connection=Form1->Database;
Tab->CommandText="Select * From Logins Where Role="+IntToStr(Role->ItemIndex+1)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(Users->ItemIndex);

EditLoginNumber=Tab->FieldByName("Num")->Value;

EditLogins->Mode=3;
EditLogins->Login=Tab->FieldByName("Login")->AsString;
EditLogins->ShowModal();        
}
//---------------------------------------------------------------------------
void TForm1::SaveCode(String Login, String Password)
{
if(Role->ItemIndex==0)
{

Database->Connected=false;
String CS=Database->ConnectionString;
Database->Connected=true;

for(unsigned int i=0;i<Cl->VBases.size();i++)
{
if(Cl->VBases[i].MainSpec==0)
{
Database->Connected=false;
Database->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+Cl->VBases[i].Patch+";Persist Security Info=False";
Database->LoginPrompt=false;
Database->Connected=true;

MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From Logins Where Role<>1 AND Login='"+Login+"'";
Tab->Active=true;
if(Tab->RecordCount==0)
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;


Sleep(500);

CodeText *CT=new CodeText();
String T=Login+Login;
String CodeSTR=CT->Crypt(T, Password);
delete CT;

Tab->CommandText="Select * From Logins Where Role=1";
Tab->Active=true;

Tab->Edit();
Tab->FieldByName("Login")->Value=Login;
Tab->FieldByName("Code1")->Value=CodeSTR.SubString(1,128);
Tab->FieldByName("Code2")->Value=CodeSTR.SubString(129,128);
Tab->Post();
}
else
{
MP<TADOCommand>Comm(this);
Comm->Connection=Database;
Comm->CommandText="Delete * from Logins where Login=' '";
Comm->Execute();

ShowMessage("����� ����� ��� ���������������!");
}
}
}
Database->Connected=false;
Database->ConnectionString=CS;
Database->Connected=true;
UpdateTempLogin();

}
else
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From Logins Where Num<>"+IntToStr(EditLoginNumber)+" AND Login='"+Login+"'";
Tab->Active=true;
if(Tab->RecordCount==0)
{
Tab->Active=false;

CodeText *CT=new CodeText();
String T=Login+Login;
String CodeSTR=CT->Crypt(T, Password);
delete CT;


Tab->CommandText="Select * From Logins Where Num="+IntToStr(EditLoginNumber);
Tab->Active=true;
//Tab->MoveBy(Users->ItemIndex);

Tab->Edit();
Tab->FieldByName("Login")->Value=Login;
Tab->FieldByName("Code1")->Value=CodeSTR.SubString(1,128);
Tab->FieldByName("Code2")->Value=CodeSTR.SubString(129,128);
Tab->Post();

UpdateTempLogin();
}
else
{
MP<TADOCommand>Comm(this);
Comm->Connection=Database;
Comm->CommandText="Delete * from Logins where Login=' '";
Comm->Execute();

ShowMessage("����� ����� ��� ���������������!");
}
}
}
//------------------------------------------------------


void __fastcall TForm1::UsersClick(TObject *Sender)
{
int N=Users->ItemIndex;

UpdateOtdel(N);        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::UsersMouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
int N=Users->ItemAtPos(Point(X,Y), true);
Users->ItemIndex=N;
UpdateOtdel(N);
if(Button==mbRight)
{
if(Role->ItemIndex>1)
{
if(N>=0)
{
N1->Enabled=true;
N2->Enabled=true;
N3->Enabled=true;
N4->Enabled=true;
}
else
{
N1->Enabled=true;
N2->Enabled=false;
N3->Enabled=false;
N4->Enabled=false;
}
}
else
{
if(N>=0)
{
N1->Enabled=false;
N2->Enabled=true;
N3->Enabled=false;
N4->Enabled=true;
}
else
{
N1->Enabled=false;
N2->Enabled=false;
N3->Enabled=false;
N4->Enabled=false;
}
}
}        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::OtdelsMouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
int N=Otdels->ItemAtPos(Point(X,Y),true);
Otdels->ItemIndex=N;

FreeOtdel->Items->Clear();
if(N>=0)
{
TMenuItem *M=new TMenuItem(FreeOtdel);
M->Caption="�������";
M->OnClick=SelectOtdel;
M->Tag=-1;
FreeOtdel->Items->Add(M);

M=new TMenuItem(FreeOtdel);
M->Caption="-";
M->Tag=0;
FreeOtdel->Items->Add(M);
}

MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="SELECT �������������.[����� �������������], �������������.[�������� �������������], ObslOtdel.NumObslOtdel FROM ������������� LEFT JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel WHERE (((ObslOtdel.NumObslOtdel) Is Null));";
Tab->Active=true;

for(Tab->First();!Tab->Eof;Tab->Next())
{
TMenuItem *M=new TMenuItem(FreeOtdel);
M->Caption=Tab->FieldByName("�������� �������������")->AsString;
M->OnClick=SelectOtdel;
M->Tag=Tab->FieldByName("����� �������������")->AsInteger;
FreeOtdel->Items->Add(M);
}

if(Role->ItemIndex==2)
{
if(Tab->RecordCount==0)
{
TMenuItem *M=new TMenuItem(FreeOtdel);
M=new TMenuItem(FreeOtdel);
M->Caption="��� ��������� �������";
M->Tag=0;
M->Enabled=false;
FreeOtdel->Items->Add(M);
}
}

}
//---------------------------------------------------------------------------
void __fastcall TForm1::SelectOtdel(TObject *Sender)
{

TMenuItem *Menu=(TMenuItem*)Sender;
int Num=Menu->Tag;

MP<TADODataSet>Tab1(this);
Tab1->Connection=Database;
Tab1->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" order by Login";
Tab1->Active=true;
Tab1->First();
Tab1->MoveBy(Users->ItemIndex);
int N1=Tab1->FieldByName("Num")->Value;

if(Num>0)
{


MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From ObslOtdel";
Tab->Active=true;

Tab->Insert();
Tab->FieldByName("Login")->Value=N1;
Tab->FieldByName("NumObslOtdel")->Value=Num;
Tab->Post();

UpdateOtdel(Users->ItemIndex);
}
else
{
MP<TADODataSet>Tab(this);
Tab->Connection=Database;
Tab->CommandText="Select * From ObslOtdel Where Login="+IntToStr(N1)+" Order by NumObslOtdel";
//Tab->CommandText="SELECT ObslOtdel.Login, �������������.[�������� �������������] FROM ������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel  where Login="+IntToStr(N1);
Tab->Active=true;

Tab->First();
Tab->MoveBy(Otdels->ItemIndex);
//ShowMessage(Tab->FieldByName("Login")->AsString+" "+Tab->FieldByName("NumObslOtdel")->AsString);
Tab->Delete();

UpdateOtdel(Users->ItemIndex);
}
}
//--------------------------------------------------------------------------



void __fastcall TForm1::N5Click(TObject *Sender)
{
int N=Cl->VClients.size();
if(N==0)
{
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "���������", "������ ������� �������", "");

this->Close();
}
else
{
ShowMessage("� ������� ���������� "+IntToStr(N)+" ��������!\r��������� ���� ��� ����������.");
}        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::FormCreate(TObject *Sender)
{
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "���������", "����� ������� �������", "");

Application->ShowMainForm=false;
NID.cbSize = sizeof(TNotifyIconData);
NID.hWnd=Handle;
NID.uID=1;
NID.uFlags=NIF_ICON | NIF_MESSAGE | NIF_TIP;
NID.uCallbackMessage=MyTrayIcon;


ImageList1->GetIcon(0, Application->Icon);

NID.hIcon=Application->Icon->Handle;

strcpy(NID.szTip,"������");
Shell_NotifyIcon(NIM_ADD, &NID);
}
//---------------------------------------------------------------------------
void __fastcall TForm1::MTIcon(TMessage&a)
{
POINT P;
switch( a.LParam)
{
case 514:

 Show();
 SetForegroundWindow(Handle);
 break;

case 516:
 GetCursorPos(&P);
 PopupMenu2->Popup(P.x,P.y);

}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::WMSysCommand(TMessage & Msg)
{
  switch (Msg.WParam)
  {
/*
case SC_MINIMIZE:
{
 Shell_NotifyIcon(NIM_ADD, &NID);
  BCreate->Enabled = false;
 BDelete->Enabled = true;
 BHide->Enabled = true;
Form1->Visible=false;
//ShowMessage("SC_MINIMIZE");
                     break;
}
*/
case SC_CLOSE:
{

  Shell_NotifyIcon(NIM_ADD, &NID);
//  BCreate->Enabled = false;
// BDelete->Enabled = true;
// BHide->Enabled = true;
Form1->Visible=false;
//SendMessage(Handle,WM_SYSCOMMAND,SC_MINIMIZE,0);

//ShowMessage("SC_Close");
break;
}

default:
DefWindowProc(Handle,Msg.Msg,Msg.WParam,Msg.LParam);

 }

}
//------------------------------------------------------------------
void __fastcall TForm1::N7Click(TObject *Sender)
{
//Form1->MyClients->DiaryEvent->WriteEvent(Now(),Form1->MyClients->NameComp, Form1->MyClients->Login, "���������", "������ ������� �������", "");

this->Close();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N6Click(TObject *Sender)
{
FDiary->ShowModal();        
}
//---------------------------------------------------------------------------
void __fastcall TForm1::ChangeCountClient(TObject *Sender)
{
//Label1->Caption=MyClients->ClientCount();

String S="������. ��������: "+Label1->Caption;
strcpy(NID.szTip,S.c_str());
/*
if(Net_Error)
{
ImageList1->GetIcon(2, Application->Icon);
}
else
{
if(MyClients->ClientCount()==0)
{
ImageList1->GetIcon(0, Application->Icon);
}
else
{
ImageList1->GetIcon(1, Application->Icon);
}
}
*/
if(Cl->VClients.size()==0)
{
ImageList1->GetIcon(0, Application->Icon);
}
else
{
ImageList1->GetIcon(1, Application->Icon);
}


NID.hIcon=Application->Icon->Handle;
Shell_NotifyIcon(NIM_MODIFY, &NID);
/*
ListBox1->Clear();
for(int i=0;i<MyClients->ClientCount();i++)
{
 ListBox1->Items->Add(IntToStr(MyClients->GetIDC(i))+". "+MyClients->GetName(i));
}
*/
}
//-------------------------------------


