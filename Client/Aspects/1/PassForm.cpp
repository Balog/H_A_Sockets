//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "PassForm.h"
#include "Main.h"
#include "Zastavka.h"
#include "CodeText.h"
#include "Progress.h"
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
void __fastcall TPass::Button1Click(TObject *Sender)
{
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->ADOAspect;
Tab->CommandText="Select * From TempLogins Where Login='"+CbLogin->Text+"'";
Tab->Active=true;

Tab->First();
Tab->MoveBy(CbLogin->ItemIndex);
int NumRole=Tab->FieldByName("Role")->AsInteger;

String TabLogin=Tab->FieldByName("Login")->AsString;
String Code=Tab->FieldByName("Code1")->AsString+Tab->FieldByName("Code2")->AsString;

CodeText* CT=new CodeText();
String Str=CT->EnCrypt(Code, EdPass->Text, 256);
delete CT;

String TLogin=Str.SubString(256-TabLogin.Length()*2+1,TabLogin.Length()*2);
String LG=TLogin.SubString(1,TLogin.Length()/2);
if(LG==TabLogin)
{



Zast->MClient->Login=TabLogin;

Zast->MClient->LoginResult(TabLogin,EdPass->Text, NumRole, true);
this->Close();

}
else
{
this->Hide();
ShowMessage("Пароль ошибочен");
Zast->MClient->LoginResult(TabLogin,EdPass->Text, NumRole,  false);
Sleep(2000);
Zast->Close();
}




this->Hide();
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

void __fastcall TPass::FormCreate(TObject *Sender)
{
ShowWindow(Application->Handle,SW_RESTORE);
SetWindowPos(Handle, HWND_TOPMOST,0,0,0,0,             SWP_NOMOVE|SWP_NOSIZE);
this->BringToFront();
}
//---------------------------------------------------------------------------

void __fastcall TPass::CbLoginClick(TObject *Sender)
{
EdPass->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TPass::FormShow(TObject *Sender)
{
Pass->EdPass->SetFocus();
Prog->Hide();
}
//---------------------------------------------------------------------------

