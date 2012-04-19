//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Password.h"
#include "AdmMain.h"
#include "MasterPointer.h"
#include "CodeText.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
//#include "ClientClass.h"
TPass *Pass;
//---------------------------------------------------------------------------
__fastcall TPass::TPass(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TPass::FormShow(TObject *Sender)
{
Main->MClient->RegForm(this);

Table* Logins=Main->MClient->CreateTable(this, Main->ServerName, Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);

Logins->SetCommandText("Select * from Logins Where Role=1");
bool B=Logins->Active(true);

if(B)
{
Variant Name;
Name=Logins->FieldByName("Login");

//Logins->FieldByName("Login")->Value;
CbLogin->Clear();
CbLogin->Items->Add((String)Name);
CbLogin->ItemIndex=0;
EdPass->Text="";

Main->MClient->DeleteTable(this, Logins);

EdPass->SetFocus();

Main->MClient->Stop();
}
else
{
CbLogin->Clear();
CbLogin->Items->Add("Сервер недоступен");
CbLogin->ItemIndex=0;
EdPass->Text="";
CbLogin->Enabled=false;
EdPass->Enabled=false;
}
}
//---------------------------------------------------------------------------
void __fastcall TPass::Button1Click(TObject *Sender)
{
if(Main->MClient->S_Error)
{
InputPassword();
}
else
{
Main->Close();
}
}
//---------------------------------------------------------------------------
void __fastcall TPass::EdPassKeyPress(TObject *Sender, char &Key)
{
if(Key==13)
{
InputPassword();
}
}
//---------------------------------------------------------------------------
void TPass::InputPassword()
{

Main->MClient->Start();

Table* Logins=Main->MClient->CreateTable(this, Main->ServerName, Main->VDB[Main->GetIDDBName(Main->CBDatabase->Text)].ServerDB);

Logins->SetCommandText("Select * from Logins Where Role=1");
Logins->Active(true);
Logins->First();

String TabLogin;
TabLogin=(String)Logins->FieldByName("Login");

String Code;
Code=(String)Logins->FieldByName("Code1")+(String)Logins->FieldByName("Code2");

CodeText* CT=new CodeText();
String Str=CT->EnCrypt(Code, EdPass->Text, 256);
delete CT;
String TLogin=Str.SubString(256-TabLogin.Length()*2+1,TabLogin.Length()*2);
String LG=TLogin.SubString(1,TLogin.Length()/2);
if(LG==TabLogin)
{
Main->MClient->WriteDiaryEvent("AdminARM","Пароль принят","");
Main->MClient->SetLogin(TabLogin);
Main->Start=true;
Main->InputPass=true;
this->Close();
}
else
{
Main->Start=false;
this->Hide();
Main->MClient->WriteDiaryEvent("AdminARM","Неверный пароль","Pass="+EdPass->Text);
ShowMessage("Пароль ошибочен");
Main->InputPass=false;
Main->Close();
}

Main->MClient->DeleteTable(this, Logins);

//Main->MClient->Stop();

}
//---------------------------------------------------------------------------
void __fastcall TPass::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"ADMINARM.HLP";
Application->HelpJump("IDH_АВТОРИЗАЦИЯ");
}
}
//---------------------------------------------------------------------------
