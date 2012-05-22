//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "About.h"
#include "Zastavka.h"
#include "Main.h"
#include "MasterPointer.h"
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
if(Form1->CBDatabase->ItemIndex==0)
{
 Image1->Visible=true;
 Image2->Visible=false;
}
else
{
 Image1->Visible=false;
 Image2->Visible=true;
}
LicAvail->Caption=IntToStr(Zast->MClient->VDB[Form1->CBDatabase->ItemIndex].NumLicense);
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * From Logins where Role=3 AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
LicUsed->Caption=IntToStr(Tab->RecordCount);
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

