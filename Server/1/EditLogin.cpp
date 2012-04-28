//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "EditLogin.h"
#include "Main.h"
#include "MasterPointer.h"
//#include "AdmMain.h"
#include "EditLogin.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TEditLogins *EditLogins;
//---------------------------------------------------------------------------
__fastcall TEditLogins::TEditLogins(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TEditLogins::FormShow(TObject *Sender)
{
Log->Text=Login;
Button1->Enabled=false;
if(Mode==3)
{
Log->Text=Login;
Log->Enabled=false;
}
else
{
Log->Enabled=true;
}
Pass1->Text="";
Pass2->Text="";

}
//---------------------------------------------------------------------------
void __fastcall TEditLogins::Pass1Change(TObject *Sender)
{
if(Pass1->Text==Pass2->Text & Log->Text!="" & Pass1->Text!="" & Pass2->Text!="")
{
Button1->Enabled=true;
}
else
{
Button1->Enabled=false;
}
}
//---------------------------------------------------------------------------
void __fastcall TEditLogins::Button1Click(TObject *Sender)
{

Form1->SaveCode(Log->Text, Pass1->Text);
this->Close();

}
//---------------------------------------------------------------------------
void __fastcall TEditLogins::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"ADMINARM.HLP";
Application->HelpJump("IDH_тнплю_педюйрхпнбюмхъ_свермни_гюохях");
}
}
//---------------------------------------------------------------------------

