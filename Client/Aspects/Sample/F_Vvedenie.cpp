//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//#include "Form_SendFile.h"
#include "F_Vvedenie.h"
#include "About.h"
//#include "LoadKeyFile.h"
//#include "Svod.h"
//#include "File_operations.cpp"
#include "SetZn.h"
#include "MainForm.h"
#include "Metod.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TVvedenie *Vvedenie;
//---------------------------------------------------------------------------
__fastcall TVvedenie::TVvedenie(TComponent* Owner)
        : TForm(Owner)
{

}
//---------------------------------------------------------------------------
void __fastcall TVvedenie::N6Click(TObject *Sender)
{
/*
if(Zast->Connect==true)
{
Form1->Visible=true;
this->Visible=false;
}
else
{
ShowMessage("Нет базы данных!");
}
*/
this->Hide();
Form1->Show();
}
//---------------------------------------------------------------------------
void __fastcall TVvedenie::N8Click(TObject *Sender)
{
//Zast->Close();
Form1->Close();
}
//---------------------------------------------------------------------------
void __fastcall TVvedenie::Button1Click(TObject *Sender)
{
Zast->Close();
}
//---------------------------------------------------------------------------
void __fastcall TVvedenie::N5Click(TObject *Sender)
{

Metodika->Visible=true;
Metodika->Position=poScreenCenter;
this->Visible=false;

}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::N7Click(TObject *Sender)
{

//FAbout->ShowModal();

}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::FormClose(TObject *Sender, TCloseAction &Action)
{
Form1->Close();       
}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::N1Click(TObject *Sender)
{

Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_СТАРТ_ПОЛЬЗОВАТЕЛЬ");

}
//---------------------------------------------------------------------------



void __fastcall TVvedenie::N9Click(TObject *Sender)
{

//SaveSF();
}
//---------------------------------------------------------------------------



void __fastcall TVvedenie::N2Click(TObject *Sender)
{
/*
if(Zast->MEcolog!=1)
{
ShowMessage("Доступ только с правами Администратора");
}
else
{
this->Visible=false;
Documents->Visible=true;
}
*/
}
//---------------------------------------------------------------------------


void __fastcall TVvedenie::FormShow(TObject *Sender)
{

//NameForm(this, "Введение");
}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::N10Click(TObject *Sender)
{
/*
if(Zast->MEcolog!=1)
{
ShowMessage("Доступ только с правами Администратора");
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
 ShowMessage("Нет базы данных!");
}
}
*/
}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::N11Click(TObject *Sender)
{
//
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
  ShowMessage("Нет базы данных!");
 }
 */
}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::Button2Click(TObject *Sender)
{
//ShowMessage(this->Top);

}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::FormActivate(TObject *Sender)
{
//this->Top=0;
//this->Left=0;
this->Width=1024;
this->Height=742;
}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::N12Click(TObject *Sender)
{
/*
FSvod->Visible=true;
this->Visible=false;
*/
}
//---------------------------------------------------------------------------


void __fastcall TVvedenie::N14Click(TObject *Sender)
{
/*
CreateNew();
NameForm(this, "Введение");
*/
}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::N17Click(TObject *Sender)
{
/*
Connect();
NameForm(this, "Введение");
*/
}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::N15Click(TObject *Sender)
{
/*
Rename();
NameForm(this, "Введение");
*/
}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::N16Click(TObject *Sender)
{
/*
SaveAs();
NameForm(this, "Введение");
*/
}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::N18Click(TObject *Sender)
{
//FullPatch();
}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::N19Click(TObject *Sender)
{
/*
if(Zast->MEcolog!=1)
{
ShowMessage("Доступ только с правами Администратора");
}
else
{
SZn->ShowModal();
}
*/
}
//---------------------------------------------------------------------------
void __fastcall TVvedenie::WMSysCommand(TMessage & Msg)
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


void __fastcall TVvedenie::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_ВВЕДЕНИЕ");
}
}
//---------------------------------------------------------------------------

