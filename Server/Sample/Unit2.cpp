//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Unit2.h"
#include "UProba.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm2 *Form2;
//---------------------------------------------------------------------------
__fastcall TForm2::TForm2(TComponent* Owner)
        : TForm(Owner)
{
MForm=new Form(Form1->ServerName);
int N=MForm->RegForm(Form1->MClient->IDC(), this->Name);
N=N;
}
//---------------------------------------------------------------------------


void __fastcall TForm2::FormDestroy(TObject *Sender)
{
delete MForm;
}
//---------------------------------------------------------------------------
