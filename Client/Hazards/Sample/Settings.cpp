//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Settings.h"
#include "About.h"
//#include "LoadKeyFile.h"
//#include "Svod.h"
//#include "File_operations.cpp"
#include "SetZn.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TSetting *Setting;
//---------------------------------------------------------------------------
__fastcall TSetting::TSetting(TComponent* Owner)
        : TForm(Owner)
{
/*
if(Zast->Connect==true)
{
Podr->Active=true;
Sit->Active=true;
}
*/
}
//---------------------------------------------------------------------------
void __fastcall TSetting::N6Click(TObject *Sender)
{
/*
Form1->Visible=true;
this->Visible=false;
*/
}
//---------------------------------------------------------------------------
void __fastcall TSetting::FormClose(TObject *Sender, TCloseAction &Action)
{
Zast->Close();
}
//---------------------------------------------------------------------------
void __fastcall TSetting::N4Click(TObject *Sender)
{
/*
Vvedenie->Visible=true;
this->Visible=false;
*/
}
//---------------------------------------------------------------------------
void __fastcall TSetting::N5Click(TObject *Sender)
{
/*
Metodika->Visible=true;
this->Visible=false;
*/
}
//---------------------------------------------------------------------------
void __fastcall TSetting::N2Click(TObject *Sender)
{
/*
if(Zast->MEcolog!=1)
{
ShowMessage("������ ������ � ������� ��������������");
}
else
{
this->Visible=false;
Documents->Visible=true;
}
*/
}
//---------------------------------------------------------------------------
void __fastcall TSetting::N8Click(TObject *Sender)
{
Zast->Close();
}
//---------------------------------------------------------------------------
void __fastcall TSetting::N9Click(TObject *Sender)
{
Podr->Insert();
Podr->FieldByName("�������� �������������")->Value="����� �������������";
Podr->Post();
}
//---------------------------------------------------------------------------

void __fastcall TSetting::N10Click(TObject *Sender)
{
/*
Temp->Active=false;
int A=Podr->FieldByName("����� �������������")->Value;
Temp->CommandText="Select * From ������� Where �������������="+IntToStr(A);
Temp->Active=true;
if (Temp->RecordCount!=0)
{
AnsiString A="�� �������?\n��������� ������ ���������� "+IntToStr(Temp->RecordCount)+" ��������!";
int Q=Application->MessageBoxA(A.c_str(),"�������� �������������",MB_YESNO);
if(Q==IDYES)
{
Podr->Delete();
}
}
else
{
Podr->Delete();
Form1->Aspects->Active=false;
Form1->Aspects->Active=true;
Form1->Podrazdel->Active=false;
Form1->Podrazdel->Active=true;
}
*/
}
//---------------------------------------------------------------------------

void __fastcall TSetting::N11Click(TObject *Sender)
{
Sit->Insert();
Sit->FieldByName("�������� ��������")->Value="����� ��������";
Sit->Post();
}
//---------------------------------------------------------------------------

void __fastcall TSetting::N12Click(TObject *Sender)
{
/*
Temp->Active=false;
int A=Sit->FieldByName("����� ��������")->Value;
Temp->CommandText="Select * From ������� Where ��������="+IntToStr(A);
Temp->Active=true;
if (Temp->RecordCount!=0)
{
AnsiString A="�� �������?\n��������� ������ ���������� "+IntToStr(Temp->RecordCount)+" ��������!";
int Q=Application->MessageBoxA(A.c_str(),"�������� ��������",MB_YESNO);
if(Q==IDYES)
{
Sit->Delete();
}
}
else
{
Sit->Delete();
Form1->Aspects->Active=false;
Form1->Aspects->Active=true;
Form1->Situaciya->Active=false;
Form1->Situaciya->Active=true;
}
*/
}
//---------------------------------------------------------------------------

void __fastcall TSetting::FormShow(TObject *Sender)
{
/*
NameForm(this, "������");
Initialize();
PageControl1->TabIndex=0;
*/
}
//---------------------------------------------------------------------------

void __fastcall TSetting::N15Click(TObject *Sender)
{
//SaveSF();
}
//---------------------------------------------------------------------------


void __fastcall TSetting::N7Click(TObject *Sender)
{
//FAbout->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TSetting::N16Click(TObject *Sender)
{
/*
if (LoadKey())
{
this->Hide();
this->Show();
}
*/
}
//---------------------------------------------------------------------------

void __fastcall TSetting::FormActivate(TObject *Sender)
{
this->Top=0;
this->Left=0;
this->Width=1024;
this->Height=742;
//PageControl1->TabIndex=0;
PageControl1->Pages[0]->Show();
Podr->Active=false;
Podr->Active=true;
Sit->Active=false;
Sit->Active=true;
}
//---------------------------------------------------------------------------

void __fastcall TSetting::N17Click(TObject *Sender)
{
/*
FSvod->Visible=true;
this->Visible=false;
*/
}
//---------------------------------------------------------------------------



void __fastcall TSetting::N19Click(TObject *Sender)
{
/*
CreateNew();
NameForm(this, "������");
Initialize();
*/
}
//---------------------------------------------------------------------------

void __fastcall TSetting::N20Click(TObject *Sender)
{
/*
Connect();
NameForm(this, "������");
Initialize();
*/
}
//---------------------------------------------------------------------------

void __fastcall TSetting::N21Click(TObject *Sender)
{
/*
Rename();
NameForm(this, "������");
Initialize();
*/
}
//---------------------------------------------------------------------------

void __fastcall TSetting::N22Click(TObject *Sender)
{
/*
SaveAs();
NameForm(this, "������");
Initialize();
*/
}
//---------------------------------------------------------------------------
void TSetting::Initialize()
{
/*
Podr->Active=true;
Sit->Active=true;
*/
}

void __fastcall TSetting::N23Click(TObject *Sender)
{
/*
FullPatch();
*/
}
//---------------------------------------------------------------------------

void __fastcall TSetting::N1Click(TObject *Sender)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"HELP_ASPECTS.HLP";
Application->HelpJump("������");        
}
//---------------------------------------------------------------------------

void __fastcall TSetting::N24Click(TObject *Sender)
{
/*
if(Zast->MEcolog!=1)
{
ShowMessage("������ ������ � ������� ��������������");
}
else
{
if(Zast->MEcolog!=1)
{
ShowMessage("������ ������ � ������� ��������������");
}
else
{
SZn->ShowModal();
}
}
*/
}
//---------------------------------------------------------------------------

