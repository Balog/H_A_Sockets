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
SignComplete=true;
}
//---------------------------------------------------------------------------
void __fastcall TProg::FormCreate(TObject *Sender)
{
ShowWindow(Application->Handle,SW_RESTORE);
SetWindowPos(Handle, HWND_TOPMOST,0,0,0,0,
this->BringToFront();        
}
//---------------------------------------------------------------------------
