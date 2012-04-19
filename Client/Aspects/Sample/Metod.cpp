//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//#include "Form_SendFile.h"
#include "Metod.h"
//#include "About.h"
#include "Zastavka.h"
#include "About.h"
//#include "LoadKeyFile.h"
//#include "Svod.h"
//#include "File_operations.cpp"
#include "SetZn.h"
#include "MasterPointer.h"
#include "About.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TMetodika *Metodika;
//---------------------------------------------------------------------------
__fastcall TMetodika::TMetodika(TComponent* Owner)
        : TForm(Owner)
{
Registered=false;
//ADODataSet1->Active=true;
/*
if(Zast->WrMetod==1)
{
 DBMemo1->ReadOnly=false;
 Button1->Visible=true;
}
else
{
 DBMemo1->ReadOnly=true;
 Button1->Visible=false;
}
*/
}
//---------------------------------------------------------------------------
void __fastcall TMetodika::N6Click(TObject *Sender)
{
Form1->Visible=true;
this->Visible=false;
/*
if(Zast->Connect==true)
{
Form1->Visible=true;
this->Visible=false;
}
else
{
ShowMessage("��� ���� ������!");
}
*/
}
//---------------------------------------------------------------------------
void __fastcall TMetodika::N4Click(TObject *Sender)
{
Vvedenie->Show();
Vvedenie->Position=poScreenCenter;
this->Visible=false;        
}
//---------------------------------------------------------------------------


void __fastcall TMetodika::N7Click(TObject *Sender)
{

//FAbout->ShowModal();

}
//---------------------------------------------------------------------------

void __fastcall TMetodika::FormClose(TObject *Sender, TCloseAction &Action)
{
Form1->Close();
}
//---------------------------------------------------------------------------

void __fastcall TMetodika::N8Click(TObject *Sender)
{
//Zast->Close();
Form1->Close();
}
//---------------------------------------------------------------------------

void __fastcall TMetodika::N1Click(TObject *Sender)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_�����_������������");
}
//---------------------------------------------------------------------------





void __fastcall TMetodika::N2Click(TObject *Sender)
{
FAbout->ShowModal();
}
//---------------------------------------------------------------------------



void __fastcall TMetodika::FormShow(TObject *Sender)
{
if(!Registered)
{
Zast->MClient->Start();
Zast->MClient->RegForm(this);
ReadMet();
Registered=true;
Zast->MClient->Stop();
}


//NameForm(this, "�������� ������");
}
//---------------------------------------------------------------------------
String TMetodika::ReadMet()
{
Zast->MClient->WriteDiaryEvent("NetAspects","������ �������� ��������","");
try
{
if(LoadMet())
{
ADODataSet1->Active=false;
ADODataSet1->Connection=Zast->ADOConn;
ADODataSet1->CommandText="select * from ��������";
ADODataSet1->Active=true;
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ��������","");
return "���������";

}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects","���� �������� ��������","");
return "������";
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� ��������"," ������ "+IntToStr(GetLastError()));
return "������";
}
}
//-----------------------------------------------------------
bool TMetodika::LoadMet()
{
bool Ret=false;
if(Zast->MClient->PrepareLoadMetod("Reference"))
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
Comm->CommandText="Delete * From TempMet";
Comm->Execute();

MP<TADODataSet>TempMet(this);
TempMet->Connection=Zast->ADOConn;
TempMet->CommandText="Select TempMet.[�����], TempMet.[��������] From TempMet order by �����";
TempMet->Active=true;

Table* STMet=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Reference")].ServerDB);

STMet->SetCommandText("Select TempMet.[�����], TempMet.[��������] From TempMet order by �����");
STMet->Active(true);

Zast->MClient->LoadTable(STMet, TempMet);

if(Zast->MClient->VerifyTable(STMet, TempMet)==0)
{
Comm->CommandText="Delete * from  ��������";
Comm->Execute();

MP<TADODataSet>Metod(this);
Metod->Connection=Zast->ADOConn;
Metod->CommandText="Select * from �������� ";
Metod->Active=true;

MP<TMemo>M(this);
M->Visible=false;
M->Parent=this;

TStrings* TT=M->Lines;
TT->Clear();
for(TempMet->First();!TempMet->Eof;TempMet->Next())
{
String S=TempMet->FieldByName("��������")->AsString;
TT->Append(S);
}
Metod->Insert();
Metod->FieldByName("�����")->Value=1;
Metod->FieldByName("��������")->Assign(TT);
Metod->Post();

Ret=true;
}
Zast->MClient->DeleteTable(this, STMet);
}
return Ret;
}
//----------------------------------------------------------
void __fastcall TMetodika::N3Click(TObject *Sender)
{
/*
if(Zast->MEcolog!=1)
{
ShowMessage("������ ������ � ������� ��������������");
}
else
{
if(Zast->Connect==true)
{
 this->Visible=false;
Setting->Visible=true;
}
else
{
 ShowMessage("��� ���� ������!");
}
}
*/
}
//---------------------------------------------------------------------------

void __fastcall TMetodika::N11Click(TObject *Sender)
{
//SaveSF();
}
//---------------------------------------------------------------------------

void __fastcall TMetodika::N12Click(TObject *Sender)
{
/*
if(Zast->Connect)
{
 if (LoadKey())
 {
  this->Hide();
  this->Show();
 }

 }
 else
 {
  ShowMessage("��� ���� ������!");
 }
 */
}
//---------------------------------------------------------------------------

void __fastcall TMetodika::FormActivate(TObject *Sender)
{
//this->Top=0;
//this->Left=0;
this->Width=1024;
this->Height=742;        
}
//---------------------------------------------------------------------------

void __fastcall TMetodika::N13Click(TObject *Sender)
{
/*
FSvod->Visible=true;
this->Visible=false;
*/
}
//---------------------------------------------------------------------------

void __fastcall TMetodika::N15Click(TObject *Sender)
{
/*
CreateNew();
NameForm(this, "��������");
*/
}
//---------------------------------------------------------------------------

void __fastcall TMetodika::N16Click(TObject *Sender)
{
/*
Connect();
NameForm(this, "��������");
*/
}
//---------------------------------------------------------------------------

void __fastcall TMetodika::N17Click(TObject *Sender)
{
/*
Rename();
NameForm(this, "��������");
*/
}
//---------------------------------------------------------------------------

void __fastcall TMetodika::N18Click(TObject *Sender)
{
/*
SaveAs();
NameForm(this, "��������");
*/
}
//---------------------------------------------------------------------------




void __fastcall TMetodika::N19Click(TObject *Sender)
{
/*
if(Zast->MEcolog!=1)
{
ShowMessage("������ ������ � ������� ��������������");
}
else
{
SZn->ShowModal();
}
*/       
}
//---------------------------------------------------------------------------

void __fastcall TMetodika::FormHide(TObject *Sender)
{
//ADODataSet1->UpdateBatch();        
}
//---------------------------------------------------------------------------
void __fastcall TMetodika::WMSysCommand(TMessage & Msg)
{
  switch (Msg.WParam)
  {
case SC_MINIMIZE:
{
Application->Minimize();
break;
}

default:
DefWindowProc(Handle,Msg.Msg,Msg.WParam,Msg.LParam);

 }
}
void __fastcall TMetodika::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_��������_������");
}
}
//---------------------------------------------------------------------------

