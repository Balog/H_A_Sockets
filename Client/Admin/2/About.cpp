//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "About.h"
//#include "AdmMain.h"
#include "Zastavka.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFAbout *FAbout;
//---------------------------------------------------------------------------
__fastcall TFAbout::TFAbout(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TFAbout::Button1Click(TObject *Sender)
{
this->Close();
}
//---------------------------------------------------------------------------



void __fastcall TFAbout::FormClick(TObject *Sender)
{
this->Close();
}
//---------------------------------------------------------------------------




void __fastcall TFAbout::FormShow(TObject *Sender)
{

if(Zast->MClient->Reg)
{
Label1->Visible=false;
}
else
{
Label1->Visible=true;
}

}
//---------------------------------------------------------------------------

void __fastcall TFAbout::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"ADMINARM.HLP";
Application->HelpJump("IDH_�_���������");
}
}
//---------------------------------------------------------------------------
