//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "PassForm.h"
#include "Main.h"
#include "Zastavka.h"
#include "CodeText.h"
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

 Zast->MClient->ReadTable("Аспекты","Select Login, Code1, Code2 from Logins Where Role=1", "Select Login, Code1, Code2 From TempLogins");

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
EdPass->Text="";
EdPass->SetFocus();
}
//---------------------------------------------------------------------------
void __fastcall TPass::Button1Click(TObject *Sender)
{
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * From TempLogins Where Login='"+CbLogin->Text+"'";
Tab->Active=true;

String TabLogin=Tab->FieldByName("Login")->AsString;
String Code=Tab->FieldByName("Code1")->AsString+Tab->FieldByName("Code2")->AsString;

CodeText* CT=new CodeText();
String Str=CT->EnCrypt(Code, EdPass->Text, 256);
delete CT;

String TLogin=Str.SubString(256-TabLogin.Length()*2+1,TabLogin.Length()*2);
String LG=TLogin.SubString(1,TLogin.Length()/2);
if(LG==TabLogin)
{
//Main->MClient->WriteDiaryEvent("AdminARM","Пароль принят","");
//Zast->MClient->SetLogin(TabLogin);
//Main->Start=true;
//Main->InputPass=true;
this->Close();
Form1->Show();
}
else
{
//Main->Start=false;
this->Hide();
//Main->MClient->WriteDiaryEvent("AdminARM","Неверный пароль","Pass="+EdPass->Text);
ShowMessage("Пароль ошибочен");
//Main->InputPass=false;
Zast->Close();
}




this->Hide();
Form1->Show();
}
//---------------------------------------------------------------------------


void __fastcall TPass::EdPassKeyPress(TObject *Sender, char &Key)
{
if(Key==13)
{
Button1->Click();
}
}
//---------------------------------------------------------------------------

