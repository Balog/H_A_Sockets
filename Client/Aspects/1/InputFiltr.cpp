//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "InputFiltr.h"
#include "InpDocs.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFilter *Filter;
//---------------------------------------------------------------------------
__fastcall TFilter::TFilter(TComponent* Owner)
        : TForm(Owner)
{
Filtr1="";
}
//---------------------------------------------------------------------------
void __fastcall TFilter::Button6Click(TObject *Sender)
{
this->Close();
}
//---------------------------------------------------------------------------
void __fastcall TFilter::RadioGroup1Click(TObject *Sender)
{
/*
Edit1->Text="";
Edit2->Text="";
Edit3->Text="";
Edit4->Text="";
*/
 Button5->Enabled=false;
switch (RadioGroup1->ItemIndex)
{
 case 0:
 {
  ComboBox1->Visible=false;
 Edit1->Visible=false;
 Edit2->Visible=false;
 Edit3->Visible=false;
 Edit4->Visible=false;
 Button1->Visible=false;
 Button2->Visible=false;
 Button3->Visible=false;
 Button4->Visible=false;
 CText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE Logins.ServerNum="+IntToStr(Form1->NumLogin)+" Order By [����� �������]";
  Button5->Enabled=true;
Edit1->Text="";
Edit2->Text="";
Edit3->Text="";
Edit4->Text="";
ComboBox1->ItemIndex=-1;
ComboBox2->Visible=false;
ComboBox3->Visible=false;
Filtr1="";
Filtr2="����������=True";
 break;
 }
 case 1:
 {
  ComboBox1->Visible=true;
 Edit1->Visible=false;
 Edit2->Visible=false;
 Edit3->Visible=false;
 Edit4->Visible=false;
ComboBox2->Visible=false;
ComboBox3->Visible=false;
Button1->Visible=false;
Button2->Visible=false;
Button3->Visible=false;
Button4->Visible=false;

 break;
 }
 case 2:
 {
  ComboBox1->Visible=false;
 Edit1->Visible=true;
 Edit2->Visible=false;
 Edit3->Visible=false;
 Edit4->Visible=false;
 Button1->Visible=true;
 Button2->Visible=false;
 Button3->Visible=false;
 Button4->Visible=false;
ComboBox2->Visible=false;
ComboBox3->Visible=false;


 break;
 }
 case 3:
 {
  ComboBox1->Visible=false;
 Edit1->Visible=false;
 Edit2->Visible=true;
 Edit3->Visible=false;
 Edit4->Visible=false;
 Button1->Visible=false;
 Button2->Visible=true;
 Button3->Visible=false;
 Button4->Visible=false;
ComboBox2->Visible=false;
ComboBox3->Visible=false;


 break;
 }
 case 4:
 {
  ComboBox1->Visible=false;
 Edit1->Visible=false;
 Edit2->Visible=false;
 Edit3->Visible=true;
 Edit4->Visible=false;
 Button1->Visible=false;
 Button2->Visible=false;
 Button3->Visible=true;
 Button4->Visible=false;
ComboBox2->Visible=false;
ComboBox3->Visible=false;


 break;
 }
 case 5:
 {
 ComboBox1->Visible=false;
 Edit1->Visible=false;
 Edit2->Visible=false;
 Edit3->Visible=false;
 Edit4->Visible=true;
 Button1->Visible=false;
 Button2->Visible=false;
 Button3->Visible=false;
 Button4->Visible=true;
 ComboBox2->Visible=false;
ComboBox3->Visible=false;
 break;
 }
 case 6:
 {
 ComboBox1->Visible=false;
 Edit1->Visible=false;
 Edit2->Visible=false;
 Edit3->Visible=false;
 Edit4->Visible=false;
 Button1->Visible=false;
 Button2->Visible=false;
 Button3->Visible=false;
 Button4->Visible=false;
 ComboBox2->Visible=true;
ComboBox3->Visible=false;
 break;
 }
 case 7:
 {
 ComboBox1->Visible=false;
 Edit1->Visible=false;
 Edit2->Visible=false;
 Edit3->Visible=false;
 Edit4->Visible=false;
 Button1->Visible=false;
 Button2->Visible=false;
 Button3->Visible=false;
 Button4->Visible=false;
 ComboBox2->Visible=false;
ComboBox3->Visible=true;
 break;
 }

}
}
//---------------------------------------------------------------------------
void __fastcall TFilter::FormShow(TObject *Sender)
{
//RadioGroup1->ItemIndex=5;
/*
if (Zast->Connect==true)
{

ComboBox1->Clear();
Podr->First();
for(int i=0;i<Podr->RecordCount;i++)
{
 AnsiString T=Podr->FieldByName("�������� �������������")->Value;
 ComboBox1->Items->Add(T);
 Podr->Next();
}
}
*/
Podr->Active=false;
Podr->CommandText=Form1->Podrazdel->CommandText;
Podr->Connection=Zast->ADOUsrAspect;
Podr->Active=true;

ComboBox1->Clear();
for(Podr->First();!Podr->Eof;Podr->Next())
{
ComboBox1->Items->Add(Podr->FieldByName("�������� �������������")->AsString);
}
}
//---------------------------------------------------------------------------
void __fastcall TFilter::Button1Click(TObject *Sender)
{

InputDocs->IForm=3;
InputDocs->Mode=1;
InputDocs->ShowModal();
}
//---------------------------------------------------------------------------
void __fastcall TFilter::Button2Click(TObject *Sender)
{

InputDocs->IForm=3;
InputDocs->Mode=2;
InputDocs->ShowModal();
}
//---------------------------------------------------------------------------
void __fastcall TFilter::Button3Click(TObject *Sender)
{

InputDocs->IForm=3;
InputDocs->Mode=3;
InputDocs->ShowModal();
}
//---------------------------------------------------------------------------
void __fastcall TFilter::Button4Click(TObject *Sender)
{

InputDocs->IForm=3;
InputDocs->Mode=4;
InputDocs->ShowModal();
}
//---------------------------------------------------------------------------
void TFilter::InpTer()
{
Edit1->Text=InputDocs->TextBr;
Index=InputDocs->NumBr;
CText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE [��� ����������]="+IntToStr(Index)+" AND (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+"))  Order By [����� �������]";
Filtr1="[��� ����������]="+IntToStr(Index);
Filtr2=" ����������=True AND [��� ����������]="+IntToStr(Index);
Form1->Aspects->Active=false;
Form1->Aspects->CommandText=CText;
Form1->Aspects->Active=true;
 Button5->Enabled=true;
}
//------------
void TFilter::InpDeyat()
{
Edit2->Text=InputDocs->TextBr;
Index=InputDocs->NumBr;
//"SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.AdmNum)="+IntToStr(NumLogin)+")) ORDER BY �������.[����� �������]"
CText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND ������������="+IntToStr(Index)+" Order By [����� �������]";
Filtr1="������������="+IntToStr(Index);
Filtr2="����������=True AND ������������="+IntToStr(Index);
Form1->Aspects->Active=false;
Form1->Aspects->CommandText=CText;
Form1->Aspects->Active=true;
 Button5->Enabled=true;
}
//------------
void TFilter::InpAsp()
{
Edit3->Text=InputDocs->TextBr;
Index=InputDocs->NumBr;
CText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  ������="+IntToStr(Index)+" Order By [����� �������]";
Filtr1="������="+IntToStr(Index);
Filtr2="����������=True AND ������="+IntToStr(Index);

Form1->Aspects->Active=false;
Form1->Aspects->CommandText=CText;
Form1->Aspects->Active=true;
 Button5->Enabled=true;
}
//------------
void TFilter::InpVozd()
{
Edit4->Text=InputDocs->TextBr;
Index=InputDocs->NumBr;
//CText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.Num)="+IntToStr(Form1->NumLogin)+")) AND  �����������="+IntToStr(Index)+" Order By [����� �������]";

CText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  �����������="+IntToStr(Index)+" Order By [����� �������]";

Filtr1="�����������="+IntToStr(Index);
Filtr2="����������=True AND �����������="+IntToStr(Index);

Form1->Aspects->Active=false;
Form1->Aspects->CommandText=CText;
Form1->Aspects->Active=true;
 Button5->Enabled=true;
}
//------------
void __fastcall TFilter::ComboBox1Click(TObject *Sender)
{
Podr->First();
Podr->MoveBy(ComboBox1->ItemIndex);
Index=Podr->FieldByName("����� �������������")->Value;
CText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  �������������="+IntToStr(Index)+" Order By [����� �������]";
Filtr1="";  //�� ������������� ������ ����������� �������
Filtr2="����������=True";

 Button5->Enabled=true;

}
//---------------------------------------------------------------------------
void __fastcall TFilter::Button5Click(TObject *Sender)
{
Form1->Filtr1=Filtr1;
Form1->Filtr2=Filtr2;
Form1->TempAspects->Active=false;
Form1->TempAspects->CommandText=CText;
Form1->TempAspects->Active=true;


if(Form1->TempAspects->RecordCount!=0)
{
Form1->Aspects->Active=false;
Form1->Aspects->CommandText=CText;
Form1->Aspects->Active=true;
this->Close();
Form1->InitCombo();

switch (RadioGroup1->ItemIndex)
{
 case 0:
 {
Form1->LFiltr->Caption="��������";
 break;
 }
 case 1:
 {
Form1->LFiltr->Caption="�� �������������";
 break;
 }
 case 2:
 {
Form1->LFiltr->Caption="�� ���� ����������";
 break;
 }
 case 3:
 {
Form1->LFiltr->Caption="�� ���� ������������";
 break;
 }
 case 4:
 {
Form1->LFiltr->Caption="�� �������";
 break;
 }
 case 5:
 {
Form1->LFiltr->Caption="�� �����������";
 break;
 }
 case 6:
 {
Form1->LFiltr->Caption="�� ����������";
 break;
 }
 case 7:
 {
Form1->LFiltr->Caption="�� ����������";
 break;
 }

}
}
else
{
 ShowMessage("��� �������� ��������������� ����� �������");

}
}
//---------------------------------------------------------------------------
void __fastcall TFilter::FormClose(TObject *Sender, TCloseAction &Action)
{
AnsiString A=Form1->Aspects->CommandText;
if (A=="Select * From ������� Order By [����� �������]")
{
  //Form1->BitBtn5->Enabled=true;
}
else
{
  //Form1->BitBtn5->Enabled=false;
}
}
//---------------------------------------------------------------------------
void __fastcall TFilter::ComboBox2Click(TObject *Sender)
{
if (ComboBox2->ItemIndex==0)
{
CText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.AdmNum)="+IntToStr(Form1->NumLogin)+")) AND  �������������<>0 AND ������������<>0 AND ������<>0 AND �����������<>0 AND ��������<>0 Order By [����� �������]";
Filtr1="(�������������<>0 AND ������������<>0 AND ������<>0 AND �����������<>0 AND ��������<>0)";
Filtr2="����������=True AND (�������������<>0 AND ������������<>0 AND ������<>0 AND �����������<>0 AND ��������<>0)";

}
else
{
CText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.AdmNum)="+IntToStr(Form1->NumLogin)+")) AND  (�������������=0 OR ������������=0 OR ������=0 OR �����������=0 OR ��������=0) Order By [����� �������]";
Filtr1="(�������������=0 OR ������������=0 OR ������=0 OR �����������=0 OR ��������=0)";
Filtr2="����������=True AND (�������������=0 OR ������������=0 OR ������=0 OR �����������=0 OR ��������=0)";

}
 Button5->Enabled=true;
}
//---------------------------------------------------------------------------
void __fastcall TFilter::ComboBox3Click(TObject *Sender)
{
if (ComboBox3->ItemIndex==0)
{
CText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  ����������=True ORDER BY �������.[����� �������]";
Filtr1="����������=True";
Filtr2="����������=True";


}
else
{
CText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  ����������=False ORDER BY �������.[����� �������]";
Filtr1="����������=False";
Filtr2="����������=False";


}
 Button5->Enabled=true;
}
//---------------------------------------------------------------------------


void __fastcall TFilter::ComboBox1KeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox2KeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox3KeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------


void __fastcall TFilter::Edit1DblClick(TObject *Sender)
{
Button1->Click();
}
//---------------------------------------------------------------------------

void __fastcall TFilter::Edit2DblClick(TObject *Sender)
{
Button2->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::Edit3DblClick(TObject *Sender)
{
Button3->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::Edit4DblClick(TObject *Sender)
{
Button4->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_����������_��������");
}
}
//---------------------------------------------------------------------------
