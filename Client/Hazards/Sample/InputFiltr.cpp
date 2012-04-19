//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "InputFiltr.h"
#include "MasterPointer.h"
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

ComboBox4->Visible=false;
ComboBox5->Visible=false;
ComboBox6->Visible=false;

 CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE Logins.ServerNum="+IntToStr(Form1->NumLogin)+" Order By [Номер аспекта]";
  Button5->Enabled=true;
Edit1->Text="";
Edit2->Text="";
Edit3->Text="";
Edit4->Text="";
ComboBox1->ItemIndex=-1;
ComboBox2->Visible=false;
ComboBox3->Visible=false;
Filtr1="";
Filtr2="Значимость=True";
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

ComboBox4->Visible=false;
ComboBox5->Visible=false;
ComboBox6->Visible=false;


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


ComboBox4->Visible=false;
ComboBox5->Visible=false;
ComboBox6->Visible=false;
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


ComboBox4->Visible=false;
ComboBox5->Visible=false;
ComboBox6->Visible=false;
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


ComboBox4->Visible=false;
ComboBox5->Visible=false;
ComboBox6->Visible=false;
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

ComboBox4->Visible=false;
ComboBox5->Visible=false;
ComboBox6->Visible=false;
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
 ComboBox2->Visible=false;
ComboBox3->Visible=false;

ComboBox4->Visible=true;
ComboBox5->Visible=false;
ComboBox6->Visible=false;
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

ComboBox4->Visible=false;
ComboBox5->Visible=false;
ComboBox6->Visible=false;
 break;
 }
 case 8:
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
ComboBox3->Visible=false;

ComboBox4->Visible=false;
ComboBox5->Visible=true;
ComboBox6->Visible=false;
 break;
 }
 case 9:
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
ComboBox3->Visible=false;

ComboBox4->Visible=false;
ComboBox5->Visible=false;
ComboBox6->Visible=true;
 break;
 }
 case 10:
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

ComboBox4->Visible=false;
ComboBox5->Visible=false;
ComboBox6->Visible=false;
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
 AnsiString T=Podr->FieldByName("Название подразделения")->Value;
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
ComboBox1->Items->Add(Podr->FieldByName("Название подразделения")->AsString);
}

MP<TADODataSet>Sit(this);
Sit->Connection=Zast->ADOAspect;
Sit->CommandText="Select * from Ситуации Where Показ=true order by [Номер ситуации]";
Sit->Active=true;

ComboBox4->Clear();
for(Sit->First();!Sit->Eof;Sit->Next())
{
 ComboBox4->Items->Add(Sit->FieldByName("Название ситуации")->AsString);
}

MP<TADODataSet>TP(this);
TP->Connection=Zast->ADOAspect;
TP->CommandText="Select * from [Тяжесть последствий] order by [Номер последствия]";
TP->Active=true;

ComboBox5->Clear();
for(TP->First();!TP->Eof;TP->Next())
{
 ComboBox5->Items->Add(TP->FieldByName("Наименование последствия")->AsString);
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
CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE [Вид территории]="+IntToStr(Index)+" AND (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+"))  Order By [Номер аспекта]";
Filtr1="[Вид территории]="+IntToStr(Index);
Filtr2=" Значимость=True AND [Вид территории]="+IntToStr(Index);
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
//"SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.AdmNum)="+IntToStr(NumLogin)+")) ORDER BY Аспекты.[Номер аспекта]"
CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND Деятельность="+IntToStr(Index)+" Order By [Номер аспекта]";
Filtr1="Деятельность="+IntToStr(Index);
Filtr2="Значимость=True AND Деятельность="+IntToStr(Index);
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
CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  Аспект="+IntToStr(Index)+" Order By [Номер аспекта]";
Filtr1="Аспект="+IntToStr(Index);
Filtr2="Значимость=True AND Аспект="+IntToStr(Index);

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
//CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.Num)="+IntToStr(Form1->NumLogin)+")) AND  Воздействие="+IntToStr(Index)+" Order By [Номер аспекта]";

CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  Воздействие="+IntToStr(Index)+" Order By [Номер аспекта]";

Filtr1="Воздействие="+IntToStr(Index);
Filtr2="Значимость=True AND Воздействие="+IntToStr(Index);

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
Index=Podr->FieldByName("Номер подразделения")->Value;
CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  Подразделение="+IntToStr(Index)+" Order By [Номер аспекта]";
Filtr1="";  //По подразделению фильтр накладываем позднее
Filtr2="Значимость=True";

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
Form1->LFiltr->Caption="Отключен";
 break;
 }
 case 1:
 {
Form1->LFiltr->Caption="По подразделению";
 break;
 }
 case 2:
 {
Form1->LFiltr->Caption="По участку/установке";
 break;
 }
 case 3:
 {
Form1->LFiltr->Caption="По объекту оценки";
 break;
 }
 case 4:
 {
Form1->LFiltr->Caption="По опасности";
 break;
 }
 case 5:
 {
Form1->LFiltr->Caption="По последствию";
 break;
 }
 case 6:
 {
Form1->LFiltr->Caption="По ситуации";
 break;
 }
 case 7:
 {
Form1->LFiltr->Caption="По степени риска";
 break;
 }
 case 8:
 {
Form1->LFiltr->Caption="По категории риска";
 break;
 }
 case 9:
 {
Form1->LFiltr->Caption="По приоритетности";
 break;
 }
 case 10:
 {
Form1->LFiltr->Caption="По валидности";
 break;
 }


}
}
else
{
 ShowMessage("Нет рисков удовлетворяющим этому условию");

}
}
//---------------------------------------------------------------------------
void __fastcall TFilter::FormClose(TObject *Sender, TCloseAction &Action)
{
AnsiString A=Form1->Aspects->CommandText;
if (A=="Select * From Аспекты Order By [Номер аспекта]")
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
CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.AdmNum)="+IntToStr(Form1->NumLogin)+")) AND  Подразделение<>0 AND Деятельность<>0 AND Аспект<>0 AND Воздействие<>0 AND Ситуация<>0 Order By [Номер аспекта]";
Filtr1="(Подразделение<>0 AND Деятельность<>0 AND Аспект<>0 AND Воздействие<>0 AND Ситуация<>0)";
Filtr2="Значимость=True AND (Подразделение<>0 AND Деятельность<>0 AND Аспект<>0 AND Воздействие<>0 AND Ситуация<>0)";

}
else
{
CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.AdmNum)="+IntToStr(Form1->NumLogin)+")) AND  (Подразделение=0 OR Деятельность=0 OR Аспект=0 OR Воздействие=0 OR Ситуация=0) Order By [Номер аспекта]";
Filtr1="(Подразделение=0 OR Деятельность=0 OR Аспект=0 OR Воздействие=0 OR Ситуация=0)";
Filtr2="Значимость=True AND (Подразделение=0 OR Деятельность=0 OR Аспект=0 OR Воздействие=0 OR Ситуация=0)";

}
 Button5->Enabled=true;
}
//---------------------------------------------------------------------------
void __fastcall TFilter::ComboBox3Click(TObject *Sender)
{
if (ComboBox3->ItemIndex==0)
{
CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  Значимость=True ORDER BY Аспекты.[Номер аспекта]";
Filtr1="Значимость=True";
Filtr2="Значимость=True";


}
else
{
CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  Значимость=False ORDER BY Аспекты.[Номер аспекта]";
Filtr1="Значимость=False";
Filtr2="Значимость=False";


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
Application->HelpFile=ExtractFilePath(Application->ExeName)+"HAZARDS.HLP";
Application->HelpJump("IDH_ФИЛЬТРАЦИЯ_АСПЕКТОВ");
}
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox4Click(TObject *Sender)
{
/*
CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  Значимость=True ORDER BY Аспекты.[Номер аспекта]";
Filtr1="Значимость=True";
Filtr2="Значимость=True";
*/
MP<TADODataSet>Sit(this);
Sit->Connection=Zast->ADOAspect;
Sit->CommandText="Select * From Ситуации where Показ=true order by [Номер ситуации]";
Sit->Active=true;
Sit->First();
Sit->MoveBy(ComboBox4->ItemIndex);
int NumSit=Sit->FieldByName("Номер ситуации")->Value;

CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  Ситуация="+IntToStr(NumSit)+" ORDER BY Аспекты.[Номер аспекта]";
Filtr1=" Ситуация="+IntToStr(NumSit);
Filtr2=" Ситуация="+IntToStr(NumSit);

 Button5->Enabled=true;
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox4KeyPress(TObject *Sender, char &Key)
{
Key=0;
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox5Click(TObject *Sender)
{
MP<TADODataSet>TP(this);
TP->Connection=Zast->ADOAspect;
TP->CommandText="Select * from [Тяжесть последствий] order by [Номер последствия]";
TP->Active=true;

TP->First();
TP->MoveBy(ComboBox5->ItemIndex);
int NumTP=TP->FieldByName("Номер последствия")->Value;

CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  [Тяжесть последствий]="+IntToStr(NumTP)+" ORDER BY Аспекты.[Номер аспекта]";
Filtr1=" [Тяжесть последствий]="+IntToStr(NumTP);
Filtr2=" [Тяжесть последствий]="+IntToStr(NumTP);

 Button5->Enabled=true;
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox5KeyPress(TObject *Sender, char &Key)
{
Key=0;
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox6KeyPress(TObject *Sender, char &Key)
{
Key=0;
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox6Click(TObject *Sender)
{
MP<TADODataSet>Prior(this);
Prior->Connection=Zast->ADOAspect;
Prior->CommandText="Select * from Приоритетность order by [Номер приоритетности]";
Prior->Active=true;
Prior->First();
Prior->MoveBy(ComboBox6->ItemIndex);
int NumP=Prior->FieldByName("Номер приоритетности")->Value;

CText="SELECT Аспекты.* ";
CText=CText+" FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация ";
CText=CText+" Where [Приоритетность]="+IntToStr(NumP)+" ORDER BY Аспекты.[Номер аспекта]";
Filtr1=" [Приоритетность]="+IntToStr(NumP);
Filtr2=" [Приоритетность]="+IntToStr(NumP);

 Button5->Enabled=true;
}
//---------------------------------------------------------------------------

