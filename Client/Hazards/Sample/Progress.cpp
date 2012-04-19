//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Progress.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFProg *FProg;
//---------------------------------------------------------------------------
__fastcall TFProg::TFProg(TComponent* Owner)
        : TForm(Owner)
{

}
//---------------------------------------------------------------------------
void __fastcall TFProg::FormShow(TObject *Sender)
{
ShowWindow(Application->Handle,SW_RESTORE);
SetWindowPos(Handle, HWND_TOPMOST,0,0,0,0,             SWP_NOMOVE|SWP_NOSIZE);
this->BringToFront();
Timer1->Enabled=true;}
//---------------------------------------------------------------------------

void __fastcall TFProg::FormHide(TObject *Sender)
{
SetWindowPos(Handle, HWND_NOTOPMOST, 0,0,0,0,
             SWP_NOMOVE|SWP_NOSIZE);
Timer1->Enabled=false;
}
//---------------------------------------------------------------------------

void __fastcall TFProg::Timer1Timer(TObject *Sender)
{
Label1->Repaint();
Progress->Repaint();
}
//---------------------------------------------------------------------------

void __fastcall TFProg::FormActivate(TObject *Sender)
{
Progress->Min=0;
Progress->Position=Progress->Min;
}
//---------------------------------------------------------------------------

