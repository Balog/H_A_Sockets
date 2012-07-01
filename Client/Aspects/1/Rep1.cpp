//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Rep1.h"
#include "MainForm.h"
#include "Reports.h"
#include "Zastavka.h"
#include "MasterPointer.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TReport1 *Report1;
//---------------------------------------------------------------------------
__fastcall TReport1::TReport1(TComponent* Owner)
        : TForm(Owner)
{
Registration=false;
}
//---------------------------------------------------------------------------
void __fastcall TReport1::FormShow(TObject *Sender)
{
Podr->Active=false;
Podr->Connection=Report1->RepBase;
Podr->CommandText="Select * from Temp������������� Order by [�������� �������������]";
Podr->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=Report1->RepBase;


 Report1->CPodrazdel->Clear();
Report1->CPodrazdel->Items->Add("��� �������������");
for(Report1->Podr->First();!Report1->Podr->Eof;Report1->Podr->Next())
{
Temp->Active=false;
if(Flt!="")
{
Temp->CommandText="SELECT TempAspects.*, Logins.ServerNum, �������������.ServerNum FROM (Logins INNER JOIN (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login) INNER JOIN TempAspects ON �������������.ServerNum = TempAspects.������������� Where �������������="+IntToStr(Report1->Podr->FieldByName("ServerNum")->AsInteger)+" AND "+Flt+" Order By [����� �������]";
}
else
{
Temp->CommandText="SELECT TempAspects.*, Logins.ServerNum, �������������.ServerNum FROM (Logins INNER JOIN (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login) INNER JOIN TempAspects ON �������������.ServerNum = TempAspects.������������� Where �������������="+IntToStr(Report1->Podr->FieldByName("ServerNum")->AsInteger)+" Order By [����� �������]";

}
Temp->Active=true;
Report1->Podr->Edit();
if(Temp->RecordCount!=0)
{
 Report1->CPodrazdel->Items->Add(Report1->Podr->FieldByName("�������� �������������")->Value);
Report1->Podr->FieldByName("����� �������������")->Value=0;
}
else
{
Report1->Podr->FieldByName("����� �������������")->Value=-1;
}
Report1->Podr->Post();
}
 Report1->CPodrazdel->ItemIndex=0;
if(Report1->CPodrazdel->Items->Count==1)
{
CPodrazdel->Enabled=false;
Button1->Enabled=false;
}
else
{
CPodrazdel->Enabled=true;
Button1->Enabled=true;
}


Date1->Date=Now();
Date2->Date=Now();
}
//---------------------------------------------------------------------------
void __fastcall TReport1::Button1Click(TObject *Sender)
{
 Report1->CreateRep();
}
//---------------------------------------------------------------------------
void __fastcall TReport1::CPodrazdelKeyPress(TObject *Sender, char &Key)
{
Key=0;
}
//---------------------------------------------------------------------------

void __fastcall TReport1::Date1KeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TReport1::Date2KeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TReport1::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
if(Role==2)
{
  Application->HelpJump("IDH_������_1_2");
}
else
{
  Application->HelpJump("IDH_������������_�������");
}
}
}
//---------------------------------------------------------------------------
void TReport1::CreateRep()
{
MP<TADODataSet>Temp(this);
Temp->Connection=Report1->RepBase;
Temp->CommandText="Select * from Temp������������� Where [����� �������������]=0 Order by [�������� �������������]";
Temp->Active=true;

int NumPodr=0;
if(CPodrazdel->ItemIndex!=0)
{
Temp->First();
Temp->MoveBy(CPodrazdel->ItemIndex-1);
NumPodr=Temp->FieldByName("ServerNum")->Value;
}



if(NumRep==1)
{
Reports *R=new Reports();
R->Connect=RepBase;
R->Role=Role;
R->CreateReport1(NumPodr, Date1->Date, Date2->Date, Edit1->Text, Flt, FltName, CheckBox1->Checked);
delete R;
}
else
{
Reports *R=new Reports();
R->Connect=RepBase;
R->Role=Role;
R->CreateReport2(NumPodr, Date1->Date, Date2->Date, Edit1->Text, Flt, FltName, CheckBox1->Checked);
delete R;
}
}
//--------------------------------------------------------------------------
