//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
#include "Metod.h"
#include "Zastavka.h"
#include "About.h"
#include "MasterPointer.h"
#include "About.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TMetodika *Metodika;
//---------------------------------------------------------------------------
__fastcall TMetodika::TMetodika(TComponent* Owner)
        : TForm(Owner)
{
Registered=false;
}
//---------------------------------------------------------------------------
void __fastcall TMetodika::N6Click(TObject *Sender)
{
Form1->Visible=true;
this->Visible=false;
}
//---------------------------------------------------------------------------
void __fastcall TMetodika::N4Click(TObject *Sender)
{
Vvedenie->Show();
Vvedenie->Position=poScreenCenter;
this->Visible=false;        
}
//---------------------------------------------------------------------------
void __fastcall TMetodika::FormClose(TObject *Sender, TCloseAction &Action)
{
Form1->Close();
}
//---------------------------------------------------------------------------

void __fastcall TMetodika::N8Click(TObject *Sender)
{

Form1->Close();
}
//---------------------------------------------------------------------------
void __fastcall TMetodika::N1Click(TObject *Sender)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_�����_������������");
}
//---------------------------------------------------------------------------
void __fastcall TMetodika::N2Click(TObject *Sender)
{
FAbout->ShowModal();
}
//---------------------------------------------------------------------------
void __fastcall TMetodika::FormShow(TObject *Sender)
{

ReadMet();

}
//---------------------------------------------------------------------------
void TMetodika::ReadMet()
{
ADODataSet1->Active=false;
ADODataSet1->Connection=Zast->ADOConn;
ADODataSet1->CommandText="select * from ��������";
ADODataSet1->Active=true;

}
//----------------------------------------------------------
void __fastcall TMetodika::FormActivate(TObject *Sender)
{

this->Width=1024;
this->Height=742;        
}
//---------------------------------------------------------------------------
void __fastcall TMetodika::WMSysCommand(TMessage & Msg)
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
void __fastcall TMetodika::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_��������_������");
}
}
//---------------------------------------------------------------------------

