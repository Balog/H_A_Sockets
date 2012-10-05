//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
#include "F_Vvedenie.h"
#include "About.h"
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
this->Hide();
Form1->Show();
}
//---------------------------------------------------------------------------
void __fastcall TVvedenie::N8Click(TObject *Sender)
{
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


void __fastcall TVvedenie::FormClose(TObject *Sender, TCloseAction &Action)
{
Form1->Close();       
}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::N1Click(TObject *Sender)
{

Application->HelpFile=ExtractFilePath(Application->ExeName)+"Hazards.HLP";
Application->HelpJump("IDH_ÑÒÀÐÒ_ÏÎËÜÇÎÂÀÒÅËÜ");

}
//---------------------------------------------------------------------------
void __fastcall TVvedenie::FormActivate(TObject *Sender)
{

this->Width=1024;
this->Height=742;
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
Application->HelpJump("IDH_ÂÂÅÄÅÍÈÅ");
}
}
//---------------------------------------------------------------------------

void __fastcall TVvedenie::N7Click(TObject *Sender)
{
FAbout->ShowModal();        
}
//---------------------------------------------------------------------------

