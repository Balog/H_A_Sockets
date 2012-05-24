//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Main.h"
#include "inifiles.hpp";
#include "CodeText.h"
#include "MasterPointer.h"
#include "PassForm.h"
#include "Zastavka.h"
//#include "EditLogin.h"
#include "Progress.h"
#include "About.h"
using namespace std;
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
Reg=false;
}
//---------------------------------------------------------------------------


void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{

switch (Application->MessageBoxA("�������� ��� ��������� �� ������?","����� �� ���������",MB_YESNOCANCEL	+MB_ICONQUESTION+MB_DEFBUTTON1))
{
 case IDYES:
 {
Zast->Stop=true;
Zast->MClient->Act.WaitCommand=0;

Zast->Saved=false;
Zast->PrepareSaveLogins->Execute();


 Action=caNone;
 break;
 }
 case IDNO:
 {
Zast->MClient->WriteDiaryEvent("AdminARM","���������� ������ ����� �� ������","");
Zast->Stop=true;
Sleep(2000);
Zast->Close();
 Action=caNone;
 break;
 }
 case IDCANCEL:
 {
 Action=caNone;
 break;
 }
}


}
//---------------------------------------------------------------------------

void __fastcall TForm1::RoleChange(TObject *Sender)
{
UpdateTempLogin();
Users->ItemIndex=0;
UpdateOtdel(0);
}
//---------------------------------------------------------------------------
void TForm1::UpdateTempLogin()
{
MP<TADODataSet>Tab(this);
//Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;

Users->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
 Users->Items->Add(Tab->FieldByName("Login")->AsString);
}
}
//-----------------------------------------------------------------
void TForm1::UpdateOtdel(int NumLogin)
{
if(NumLogin<0)
{
Otdels->Clear();
}
else
{
MP<TADODataSet>Tab(this);
//Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(NumLogin);
int N1=Tab->FieldByName("Num")->AsInteger;

MP<TADODataSet>Verify(this);
//Verify->Connection=Zast->MClient->Database;

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

void __fastcall TForm1::FormShow(TObject *Sender)
{
Zast->UpdateLogins->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CBDatabaseClick(TObject *Sender)
{
UpdateTempLogin();
Form1->Users->ItemIndex=0;
UpdateOtdel(0);
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

void __fastcall TForm1::N3Click(TObject *Sender)
{
/*
//������� �����
MP<TADODataSet>Tab(this);
//Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * From Logins Where Role="+IntToStr(Role->ItemIndex+1)+"  AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->MoveBy(Users->ItemIndex);

Tab->Delete();
UpdateTempLogin();

int N=Users->ItemIndex;

UpdateOtdel(N);
*/
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N1Click(TObject *Sender)
{
//�������� �����
/*
int Lic=Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumLicense;
if(Lic<=Users->Count & Role->ItemIndex==2 & Lic!=-1 )
{
 ShowMessage("���������� ������������ ���������� �������������!");
}
else
{
EditLogins->Log->Text="";
EditLogins->Pass1->Text="";
EditLogins->Pass2->Text="";

EditLogins->Mode=false;
EditLogins->NumLogin=-1;

EditLogins->ShowModal();
}
*/
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N2Click(TObject *Sender)
{
//������������� �����
/*
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * from Logins Where Role="+IntToStr(Form1->Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->RecNo=Users->ItemIndex+1;


EditLogins->Log->Text=Tab->FieldByName("Login")->AsString;
EditLogins->Pass1->Text="";
EditLogins->Pass2->Text="";

EditLogins->Mode=false;
EditLogins->NumLogin=Users->ItemIndex;

EditLogins->ShowModal();
*/
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N4Click(TObject *Sender)
{
//������������� ������
/*
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * from Logins Where Role="+IntToStr(Form1->Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->RecNo=Users->ItemIndex+1;


EditLogins->Log->Text=Tab->FieldByName("Login")->AsString;
EditLogins->Pass1->Text="";
EditLogins->Pass2->Text="";

EditLogins->Mode=true;
EditLogins->NumLogin=Users->ItemIndex;

EditLogins->ShowModal();
*/
}
//---------------------------------------------------------------------------

void __fastcall TForm1::OtdelsMouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
/*
FreeOtdel->Items->Clear();
if(Users->ItemIndex>=0)
{
int N=Otdels->ItemAtPos(Point(X,Y),true);
Otdels->ItemIndex=N;


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
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="SELECT �������������.[����� �������������], �������������.[�������� �������������], �������������.Del, �������������.NumDatabase FROM ������������� LEFT JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.[NumObslOtdel] WHERE (((�������������.NumDatabase)="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+") AND ((ObslOtdel.NumObslOtdel) Is Null));";
Tab->Active=true;

for(Tab->First();!Tab->Eof;Tab->Next())
{
TMenuItem *M=new TMenuItem(FreeOtdel);
M->Caption=Tab->FieldByName("�������� �������������")->AsString;
M->OnClick=SelectOtdel;
M->Tag=Tab->FieldByName("����� �������������")->AsInteger;
FreeOtdel->Items->Add(M);
}

if(Role->ItemIndex>=2)
{
if(Tab->RecordCount==0)
{
TMenuItem *M=new TMenuItem(FreeOtdel);
M->Caption="��� ��������� �������������";
M->Tag=0;
M->Enabled=false;
FreeOtdel->Items->Add(M);
}
}
}
else
{
TMenuItem *M=new TMenuItem(FreeOtdel);
M->Caption="�������� �����";
M->Tag=0;
M->Enabled=false;
FreeOtdel->Items->Add(M);
}
*/
}
//---------------------------------------------------------------------------
void __fastcall TForm1::SelectOtdel(TObject *Sender)
{
/*
TMenuItem *Menu=(TMenuItem*)Sender;
int Num=Menu->Tag;

MP<TADODataSet>Tab1(this);
Tab1->Connection=Zast->MClient->Database;
Tab1->CommandText="Select * From Logins where Role="+IntToStr(Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab1->Active=true;
Tab1->First();
Tab1->MoveBy(Users->ItemIndex);
int N1=Tab1->FieldByName("Num")->Value;

if(Num>0)
{


MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * From ObslOtdel";
Tab->Active=true;

Tab->Insert();
Tab->FieldByName("Login")->Value=N1;
Tab->FieldByName("NumObslOtdel")->Value=Num;
Tab->FieldByName("NumDatabase")->Value=Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase;
Tab->Post();

UpdateOtdel(Users->ItemIndex);
}
else
{
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * From ObslOtdel Where Login="+IntToStr(N1)+" Order by NumObslOtdel";
Tab->Active=true;

Tab->First();
Tab->MoveBy(Otdels->ItemIndex);
Tab->Delete();

UpdateOtdel(Users->ItemIndex);
}
*/
}
//--------------------------------------------------------------------------
void __fastcall TForm1::N5Click(TObject *Sender)
{
Prog->PB->Min=0;
Prog->PB->Position=0;
Prog->PB->Max=Zast->MClient->VDB.size()*5;
Prog->Show();
Zast->PrepareRead->Execute();
}
//---------------------------------------------------------------------------



void __fastcall TForm1::N6Click(TObject *Sender)
{
Prog->PB->Min=0;
Prog->PB->Position=0;
Prog->PB->Max=Zast->MClient->VDB.size()*5;
Prog->Show();
Zast->Saved=false;
Zast->PrepareSaveLogins->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::UsersClick(TObject *Sender)
{
int N=Users->ItemIndex;

UpdateOtdel(N);        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N7Click(TObject *Sender)
{
Prog->Show();
Prog->PB->Visible=true;
Prog->PB->Max=6;
Prog->PB->Position=0;


//FDiary->LoadDiary();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N10Click(TObject *Sender)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"ADMINARM.HLP";
Application->HelpJump("IDH_�����������");        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N8Click(TObject *Sender)
{
//FAbout->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N9Click(TObject *Sender)
{
this->Close();
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
case SC_MINIMIZE:
{

//  Shell_NotifyIcon(NIM_ADD, &NID);
//  BCreate->Enabled = false;
// BDelete->Enabled = true;
// BHide->Enabled = true;
//Form1->Visible=false;
//SendMessage(Handle,WM_SYSCOMMAND,SC_MINIMIZE,0);

//ShowMessage("SC_Close");
Application->Minimize();
break;
}

default:
DefWindowProc(Handle,Msg.Msg,Msg.WParam,Msg.LParam);

 }

}
//-------------------------------------------------------------------

