//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "PassForm.h"
#include "Main.h"
#include "Zastavka.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TPass *Pass;
//---------------------------------------------------------------------------
__fastcall TPass::TPass(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------

void __fastcall TPass::FormShow(TObject *Sender)
{

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ViewLogins");

 Zast->MClient->ReadTable("�������","Select Login, Code1, Code2 from Logins Where Role=1", "Select Login, Code1, Code2 From TempLogins");

}
//---------------------------------------------------------------------------
void TPass::ViewLogins()
{
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select Login, Code1, Code2 From TempLogins";
Tab->Active=true;

CbLogin->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
CbLogin->Items->Add(Tab->FieldByName("Login")->AsString);
}
CbLogin->ItemIndex=0;
}
//---------------------------------------------------------------------------
void __fastcall TPass::Button1Click(TObject *Sender)
{
this->Hide();
Form1->Show();
}
//---------------------------------------------------------------------------

