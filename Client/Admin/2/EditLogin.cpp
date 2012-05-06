//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "EditLogin.h"
#include "MasterPointer.h"
#include "Main.h"
#include "Zastavka.h"
#include "CodeText.h"
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
void TEditLogins::NewLogin()
{
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * from Logins Where Role="+IntToStr(Form1->Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase);
Tab->Active=true;

CodeText *CT=new CodeText();
String T=Log->Text+Log->Text;
String CodeSTR=CT->Crypt(T, Pass1->Text);
delete CT;

Tab->Insert();
Tab->FieldByName("Login")->Value=Log->Text;
Tab->FieldByName("Code1")->Value=CodeSTR.SubString(1,128);
Tab->FieldByName("Code2")->Value=CodeSTR.SubString(129,128);
Tab->FieldByName("Role")->Value=Form1->Role->ItemIndex+1;
Tab->FieldByName("NumDatabase")->Value=Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase;
Tab->Post();

Tab->Last();

int Res=Tab->FieldByName("Num")->AsInteger;
Form1->UpdateTempLogin();
Tab->Active=false;
Tab->CommandText="Select * from Logins Where Role="+IntToStr(Form1->Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->Locate("Num",Res,SO);
Form1->Users->ItemIndex=Tab->RecNo-1;
Form1->UpdateOtdel(Res);
this->Close();
}
//---------------------------------------------------------------------------
void TEditLogins::EditLogin(int Num)
{
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->MClient->Database;
Tab->CommandText="Select * from Logins Where Role="+IntToStr(Form1->Role->ItemIndex+1)+" AND NumDatabase="+IntToStr(Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase)+" order by Login";
Tab->Active=true;
Tab->First();
Tab->RecNo=Num+1;

CodeText *CT=new CodeText();
String T=Log->Text+Log->Text;
String CodeSTR=CT->Crypt(T, Pass1->Text);
delete CT;

Tab->Edit();
Tab->FieldByName("Login")->Value=Log->Text;
Tab->FieldByName("Code1")->Value=CodeSTR.SubString(1,128);
Tab->FieldByName("Code2")->Value=CodeSTR.SubString(129,128);
Tab->FieldByName("Role")->Value=Form1->Role->ItemIndex+1;
Tab->FieldByName("NumDatabase")->Value=Zast->MClient->VDB[Zast->MClient->GetIDDBName(Form1->CBDatabase->Text)].NumDatabase;
Tab->Post();

Form1->UpdateTempLogin();
this->Close();
}
//---------------------------------------------------------------------------
void TEditLogins::EditPass(int Num)
{

}
//---------------------------------------------------------------------------
void __fastcall TEditLogins::Button1Click(TObject *Sender)
{
if(NumLogin==-1)
{
 //����� �����
 NewLogin();
}
else
{
 //�������������� ������/������
EditLogin(NumLogin);

}
}
//---------------------------------------------------------------------------
void __fastcall TEditLogins::Pass1Change(TObject *Sender)
{
if(Pass1->Text==Pass2->Text & Log->Text!="")
{
 Button1->Enabled=true;
}
else
{
 Button1->Enabled=false;
}
}
//---------------------------------------------------------------------------
void __fastcall TEditLogins::FormShow(TObject *Sender)
{
Button1->Enabled=false;
if(Mode)
{
Log->Enabled=false;
}
else
{
Log->Enabled=true;
}        
}
//---------------------------------------------------------------------------
