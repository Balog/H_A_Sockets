//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Progress.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TProg *Prog;
//---------------------------------------------------------------------------
__fastcall TProg::TProg(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TProg::FormCreate(TObject *Sender)
{
ShowWindow(Application->Handle,SW_RESTORE);
SetWindowPos(Handle, HWND_TOPMOST,0,0,0,0,             SWP_NOMOVE|SWP_NOSIZE);
this->BringToFront();        
}
//---------------------------------------------------------------------------

