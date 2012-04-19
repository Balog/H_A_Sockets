//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Password.h"
#include "Docs.h"
#include "MasterPointer.h"
#include "CodeText.h"
#include "MainForm.h"
#include "F_Vvedenie.h"
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
Zast->MClient->Start();
Zast->MClient->RegForm(this);
//Table* ROtdels=MClient->CreateTable(this, ServerName, VDB[CBDatabase->ItemIndex].ServerDB);

Table* Logins=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Logins->SetCommandText("SELECT Logins.* FROM Logins WHERE (((Logins.Role)=2)) OR (((Logins.Role)=3)) OR (((Logins.Role)=4)) ORDER BY Logins.Role, Logins.Login;");
//Logins->SetCommandText("Select * from Logins Where Role=2 ");
bool B=Logins->Active(true);

if(B)
{
Variant Name;


//Logins->FieldByName("Login")->Value;
CbLogin->Clear();
for(Logins->First();!Logins->eof();Logins->Next())
{
Name=Logins->FieldByName("Login");
CbLogin->Items->Add((String)Name);
}
CbLogin->ItemIndex=0;
EdPass->Text="";

Zast->MClient->DeleteTable(this, Logins);

EdPass->SetFocus();

Zast->MClient->Stop();
}
else
{
CbLogin->Clear();
CbLogin->Items->Add("������ ����������");
CbLogin->ItemIndex=0;
EdPass->Text="";
CbLogin->Enabled=false;
EdPass->Enabled=false;
}

}
//---------------------------------------------------------------------------
void __fastcall TPass::Button1Click(TObject *Sender)
{

if(Zast->MClient->S_Error)
{
try
{
 this->Hide();
InputPassword();
}
catch(...)
{
 ShowMessage("������ �������� ������");
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� ������"," ������ "+IntToStr(GetLastError()));
}
}
else
{
Zast->Close();
}

}
//---------------------------------------------------------------------------
void __fastcall TPass::EdPassKeyPress(TObject *Sender, char &Key)
{
if(Key==13)
{
if(Zast->MClient->S_Error)
{
 this->Hide();
try
{
InputPassword();
}
catch(...)
{
 ShowMessage("������ �������� ������");
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� ������"," ������ "+IntToStr(GetLastError()));
 Zast->Close();
}
}
else
{
Zast->Close();
}
}
}
//---------------------------------------------------------------------------
void TPass::InputPassword()
{
 //ShowMessage("�������� ������");
Zast->MClient->Start();

Table* Logins=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Logins->SetCommandText("Select * from Logins Where Role=2");
Logins->Active(true);
Logins->First();

String TabLogin;
//TabLogin=(String)Logins->FieldByName("Login");
TabLogin=CbLogin->Text;
Table* Group=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Group->SetCommandText("Select * from Logins Where Login='"+TabLogin+"'");
Group->Active(true);

String Code;
Code=(String)Group->FieldByName("Code1")+(String)Group->FieldByName("Code2");

CodeText* CT=new CodeText();
String Str=CT->EnCrypt(Code, EdPass->Text, 256);
delete CT;
String TLogin=Str.SubString(256-TabLogin.Length()*2+1,TabLogin.Length()*2);
String LG=TLogin.SubString(1,TLogin.Length()/2);

 //ShowMessage("����� ��������");
if(LG==TabLogin)
{
 Zast->MClient->SetLogin(TabLogin);
Zast->MClient->WriteDiaryEvent("NetAspects","������ ������","");

int Role=Group->FieldByName("Role");
Form1->NumLogin=Group->FieldByName("Num");

Zast->MClient->DeleteTable(this, Group);
Zast->MClient->DeleteTable(this, Logins);


Form1->Role=Role;

Zast->Start=true;


switch (Role)
{
 case 2:
 {
 if(Zast->LoadLogins(Zast->ADOAspect))
 {
 //Zast->MClient->Stop();
 Zast->ADOUsrAspect->Connected=false;
 //ShowMessage("������� � ���������");

 Documents->Show();
 this->Hide();
 this->Close();
 }
 else
 {
 ShowMessage("������ ������������� �������");
 Zast->Close();
 }
 break;
 }
 case 3: case 4:
 {

 Form1->Demo=(Role==4);
 if(Zast->LoadLogins(Zast->ADOUsrAspect))
 {
 //Zast->MClient->Stop();
 Zast->ADOAspect->Connected=false;
 //ShowMessage("������� � ������������");
 Zast->MClient->SetLogin(TabLogin);
 Zast->MClient->Stop();
 Vvedenie->Show();
 this->Hide(); 
 this->Close();
 }
 else
 {
 ShowMessage("������ ������������� �������");
 Zast->Close();
 }
 break;
 }
}




}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects","�������� ������","Pass="+EdPass->Text);
Zast->MClient->DeleteTable(this, Group);
Zast->MClient->DeleteTable(this, Logins);
Zast->MClient->Stop();
Zast->Start=false;
ShowMessage("������ ��������");
Zast->Close();
}



}
//---------------------------------------------------------------------------
void __fastcall TPass::CbLoginClick(TObject *Sender)
{
EdPass->SetFocus();
}
//---------------------------------------------------------------------------
void __fastcall TPass::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_�����������");
}
}
//---------------------------------------------------------------------------

