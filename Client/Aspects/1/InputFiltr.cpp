//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "InputFiltr.h"
#include "InpDocs.h"
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

//SetFilterGroup(RadioGroup1->ItemIndex);
 Button5->Enabled=false;

switch (RadioGroup1->ItemIndex)
{
 case 0:
 {
  ComboBox1->Visible=false;
 ComboBox4->Visible=false;
 ComboBox5->Visible=false;
 ComboBox6->Visible=false;
 ComboBox7->Visible=false;
 Button1->Visible=false;
 Button2->Visible=false;
 Button3->Visible=false;
 Button4->Visible=false;
 CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE Logins.ServerNum="+IntToStr(Form1->NumLogin)+" Order By [Номер аспекта]";
  Button5->Enabled=true;
ComboBox4->Text="";
ComboBox5->Text="";
ComboBox6->Text="";
ComboBox7->Text="";
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
 ComboBox4->Visible=false;
 ComboBox5->Visible=false;
 ComboBox6->Visible=false;
 ComboBox7->Visible=false;
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
 ComboBox4->Visible=true;
 ComboBox4->Text="";
 ComboBox4->SetFocus();
 ComboBox5->Visible=false;
 ComboBox6->Visible=false;
 ComboBox7->Visible=false;
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
 ComboBox4->Visible=false;
 ComboBox5->Visible=true;
 ComboBox5->Text="";
 ComboBox5->SetFocus();
 ComboBox6->Visible=false;
 ComboBox7->Visible=false;
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
 ComboBox4->Visible=false;
 ComboBox5->Visible=false;
 ComboBox6->Visible=true;
 ComboBox6->Text="";
 ComboBox6->SetFocus();
 ComboBox7->Visible=false;
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
 ComboBox4->Visible=false;
 ComboBox5->Visible=false;
 ComboBox6->Visible=false;
 ComboBox7->Visible=true;
 ComboBox7->SetFocus();
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
 ComboBox4->Visible=false;
 ComboBox5->Visible=false;
 ComboBox6->Visible=false;
 ComboBox7->Visible=false;
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
 ComboBox4->Visible=false;
 ComboBox5->Visible=false;
 ComboBox6->Visible=false;
 ComboBox7->Visible=false;
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
void TFilter::SetFilterGroup(int NumFilter)
{

}
//----------------------------------------------------------------------------
void __fastcall TFilter::FormShow(TObject *Sender)
{

Podr->Active=false;
Podr->CommandText=Form1->Podrazdel->CommandText;
Podr->Connection=Zast->ADOUsrAspect;
Podr->Active=true;

ComboBox1->Clear();
for(Podr->First();!Podr->Eof;Podr->Next())
{
ComboBox1->Items->Add(Podr->FieldByName("Название подразделения")->AsString);
}

String Text="";
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
Tab->CommandText="Select [Номер территории], [Наименование территории] From Территория Order by [Номер территории]";//"Select [Номер территории], [Наименование территории] From Территория Where [Наименование территории] Like '%з%' Order by [Номер территории]";
Tab->Active=true;

ComboBox4->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox4->Items->Add(Tab->FieldByName("Наименование территории")->AsString);
}


Tab->Active=false;
String CT="Select [Номер деятельности], [Наименование деятельности] From Деятельность Where [Наименование деятельности] Like '%"+Text+"%' AND Показ=true Order by [Номер деятельности]";
Tab->CommandText=CT;
Tab->Active=true;
ComboBox5->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox5->Items->Add(Tab->FieldByName("Наименование деятельности")->AsString);
}


Tab->Active=false;
CT="Select [Номер аспекта], [Наименование аспекта] From Аспект Where [Наименование аспекта] Like '%"+Text+"%' AND Показ=true Order by [Номер аспекта]";
Tab->CommandText=CT;
Tab->Active=true;

ComboBox6->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox6->Items->Add(Tab->FieldByName("Наименование аспекта")->AsString);
}

Tab->Active=false;
CT="Select [Номер воздействия], [Наименование воздействия] From Воздействия Where [Наименование воздействия] Like '%"+Text+"%' AND Показ=true Order by [Номер воздействия]";
Tab->CommandText=CT;
Tab->Active=true;
ComboBox7->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox7->Items->Add(Tab->FieldByName("Наименование воздействия")->AsString);
}

RadioGroup1->ItemIndex=NumFiltr;
switch (NumFiltr)
{
 case 1:
 {
 ComboBox1->ItemIndex=NumItemCombo;
 break;
 }
 case 2:
 {
 ComboBox4->Text=InputDocs->TextBr;
 break;
 }
 case 3:
 {
 ComboBox5->Text=InputDocs->TextBr;
 break;
 }
 case 4:
 {
 ComboBox6->Text=InputDocs->TextBr;
 break;
 }
 case 5:
 {
 ComboBox7->Text=InputDocs->TextBr;
 break;
 }
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
ComboBox4->Text=InputDocs->TextBr;
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
ComboBox5->Text=InputDocs->TextBr;
Index=InputDocs->NumBr;
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
ComboBox6->Text=InputDocs->TextBr;
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
ComboBox7->Text=InputDocs->TextBr;
Index=InputDocs->NumBr;
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
NumItemCombo=ComboBox1->ItemIndex;
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
NumFiltr=0;
 break;
 }
 case 1:
 {
Form1->LFiltr->Caption="По подразделению";
NumFiltr=1;
 break;
 }
 case 2:
 {
Form1->LFiltr->Caption="По виду территории";
NumFiltr=2;
 break;
 }
 case 3:
 {
Form1->LFiltr->Caption="По виду деятельности";
NumFiltr=3;
 break;
 }
 case 4:
 {
Form1->LFiltr->Caption="По аспекту";
NumFiltr=4;
 break;
 }
 case 5:
 {
Form1->LFiltr->Caption="По воздействию";
NumFiltr=5;
 break;
 }
 case 6:
 {
Form1->LFiltr->Caption="По валидности";
NumFiltr=6;
 break;
 }
 case 7:
 {
Form1->LFiltr->Caption="По значимости";
NumFiltr=7;
 break;
 }

}
}
else
{
 ShowMessage("Нет аспектов удовлетворяющим этому условию");

}
}
//---------------------------------------------------------------------------
void __fastcall TFilter::FormClose(TObject *Sender, TCloseAction &Action)
{
AnsiString A=Form1->Aspects->CommandText;

}
//---------------------------------------------------------------------------
void __fastcall TFilter::ComboBox2Click(TObject *Sender)
{
if (ComboBox2->ItemIndex==0)
{
CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  [Вид территории]<>0 AND Деятельность<>0 AND Аспект<>0 AND Воздействие<>0 AND Ситуация<>0 Order By [Номер аспекта]";
Filtr1="([Вид территории]<>0 AND Деятельность<>0 AND Аспект<>0 AND Воздействие<>0 AND Ситуация<>0)";
Filtr2="Значимость=True AND ([Вид территории]<>0 AND Деятельность<>0 AND Аспект<>0 AND Воздействие<>0 AND Ситуация<>0)";

}
else
{
CText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  ([Вид территории]=0 OR Деятельность=0 OR Аспект=0 OR Воздействие=0 OR Ситуация=0) Order By [Номер аспекта]";
Filtr1="([Вид территории]=0 OR Деятельность=0 OR Аспект=0 OR Воздействие=0 OR Ситуация=0)";
Filtr2="Значимость=True AND ([Вид территории]=0 OR Деятельность=0 OR Аспект=0 OR Воздействие=0 OR Ситуация=0)";

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
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_ФИЛЬТРАЦИЯ_АСПЕКТОВ");
}
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox4Change(TObject *Sender)
{
String Text=ComboBox4->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер территории], [Наименование территории] From Территория Where [Наименование территории] Like '%"+Text+"%' AND Показ=true Order by [Номер территории]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox4->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox4->DropDownCount=30;
}

ComboBox4->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox4->Items->Add(Tab->FieldByName("Наименование территории")->AsString);
}



ComboBox4->DroppedDown=true;
ComboBox4->AutoComplete=false;


ComboBox4->Text=Text;
 keybd_event(35,0,0,0);

 keybd_event(35,0,KEYEVENTF_KEYUP,0);        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox4DropDown(TObject *Sender)
{
String Text=ComboBox4->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер территории], [Наименование территории] From Территория Where [Наименование территории] Like '%"+Text+"%' AND Показ=true Order by [Номер территории]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox4->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox4->DropDownCount=30;
}

ComboBox4->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox4->Items->Add(Tab->FieldByName("Наименование территории")->AsString);
}        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox4KeyPress(TObject *Sender, char &Key)
{
if(Key==13)
{
String Text=ComboBox4->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер территории], [Наименование территории] From Территория Where [Наименование территории] Like '%"+Text+"%' AND Показ=true Order by [Номер территории]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox4->ItemIndex);
ComboBox4->Text=Tab->FieldByName("Наименование территории")->AsString;
if(ComboBox4->DroppedDown & ComboBox4->ItemIndex!=-1)
{
//MAsp->Index=Tab->FieldByName("Номер территории")->AsInteger;
InputDocs->NumBr=Tab->FieldByName("Номер территории")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("Наименование территории")->AsString;
InpTer();
}
}        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox4Select(TObject *Sender)
{
String Text=ComboBox4->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер территории], [Наименование территории] From Территория Where [Наименование территории] Like '%"+Text+"%' AND Показ=true Order by [Номер территории]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox4->ItemIndex);
ComboBox4->Text=Tab->FieldByName("Наименование территории")->AsString;
if(!ComboBox4->DroppedDown & ComboBox4->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("Номер территории")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("Наименование территории")->AsString;
InpTer();
}
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox4Enter(TObject *Sender)
{
ComboBox4->DroppedDown=true;
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox5Change(TObject *Sender)
{
String Text=ComboBox5->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер деятельности], [Наименование деятельности] From деятельность Where [Наименование деятельности] Like '%"+Text+"%' AND Показ=true Order by [Номер деятельности]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox5->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox5->DropDownCount=30;
}

ComboBox5->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox5->Items->Add(Tab->FieldByName("Наименование деятельности")->AsString);
}



ComboBox5->DroppedDown=true;
ComboBox5->AutoComplete=false;


ComboBox5->Text=Text;
 keybd_event(35,0,0,0);

 keybd_event(35,0,KEYEVENTF_KEYUP,0);      
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox5DropDown(TObject *Sender)
{
String Text=ComboBox5->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер деятельности], [Наименование деятельности] From Деятельность Where [Наименование деятельности] Like '%"+Text+"%' AND Показ=true Order by [Номер деятельности]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox5->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox5->DropDownCount=30;
}

ComboBox5->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox5->Items->Add(Tab->FieldByName("Наименование деятельности")->AsString);
}
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox5KeyPress(TObject *Sender, char &Key)
{
if(Key==13)
{
String Text=ComboBox5->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер деятельности], [Наименование деятельности] From Деятельность Where [Наименование деятельности] Like '%"+Text+"%' AND Показ=true Order by [Номер деятельности]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox5->ItemIndex);
ComboBox5->Text=Tab->FieldByName("Наименование деятельности")->AsString;
if(ComboBox5->DroppedDown & ComboBox5->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("Номер деятельности")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("Наименование деятельности")->AsString;
InpDeyat();
}
}


}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox5Select(TObject *Sender)
{
String Text=ComboBox5->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер деятельности], [Наименование деятельности] From Деятельность Where [Наименование деятельности] Like '%"+Text+"%' AND Показ=true Order by [Номер деятельности]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox5->ItemIndex);
ComboBox5->Text=Tab->FieldByName("Наименование деятельности")->AsString;
if(!ComboBox5->DroppedDown & ComboBox5->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("Номер деятельности")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("Наименование деятельности")->AsString;
InpDeyat();
}
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox5Enter(TObject *Sender)
{
ComboBox5->DroppedDown=true;        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox6Change(TObject *Sender)
{
String Text=ComboBox6->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер аспекта], [Наименование аспекта] From Аспект Where [Наименование аспекта] Like '%"+Text+"%' AND Показ=true Order by [Номер аспекта]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox6->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox6->DropDownCount=30;
}

ComboBox6->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox6->Items->Add(Tab->FieldByName("Наименование аспекта")->AsString);
}



ComboBox6->DroppedDown=true;
ComboBox6->AutoComplete=false;


ComboBox6->Text=Text;
 keybd_event(35,0,0,0);

 keybd_event(35,0,KEYEVENTF_KEYUP,0);        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox6DropDown(TObject *Sender)
{
String Text=ComboBox6->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер аспекта], [Наименование аспекта] From Аспект Where [Наименование аспекта] Like '%"+Text+"%' AND Показ=true Order by [Номер аспекта]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox6->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox6->DropDownCount=30;
}

ComboBox6->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox6->Items->Add(Tab->FieldByName("Наименование аспекта")->AsString);
}        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox6KeyPress(TObject *Sender, char &Key)
{
if(Key==13)
{
String Text=ComboBox6->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер аспекта], [Наименование аспекта] From Аспект Where [Наименование аспекта] Like '%"+Text+"%' AND Показ=true Order by [Номер аспекта]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox6->ItemIndex);
ComboBox6->Text=Tab->FieldByName("Наименование аспекта")->AsString;
if(ComboBox6->DroppedDown & ComboBox6->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("Номер аспекта")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("Наименование аспекта")->AsString;
InpAsp();
}
}        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox6Select(TObject *Sender)
{
String Text=ComboBox6->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер аспекта], [Наименование аспекта] From Аспект Where [Наименование аспекта] Like '%"+Text+"%' AND Показ=true Order by [Номер аспекта]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox6->ItemIndex);
ComboBox6->Text=Tab->FieldByName("Наименование аспекта")->AsString;
if(!ComboBox6->DroppedDown & ComboBox6->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("Номер аспекта")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("Наименование аспекта")->AsString;
InpAsp();
}
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox6Enter(TObject *Sender)
{
ComboBox6->DroppedDown=true;          
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox7Change(TObject *Sender)
{
String Text=ComboBox7->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер воздействия], [Наименование воздействия] From Воздействия Where [Наименование воздействия] Like '%"+Text+"%' AND Показ=true Order by [Номер воздействия]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox7->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox7->DropDownCount=30;
}

ComboBox7->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox7->Items->Add(Tab->FieldByName("Наименование воздействия")->AsString);
}



ComboBox7->DroppedDown=true;
ComboBox7->AutoComplete=false;


ComboBox7->Text=Text;
 keybd_event(35,0,0,0);

 keybd_event(35,0,KEYEVENTF_KEYUP,0);        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox7DropDown(TObject *Sender)
{
String Text=ComboBox7->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер воздействия], [Наименование воздействия] From Воздействия Where [Наименование воздействия] Like '%"+Text+"%' AND Показ=true Order by [Номер воздействия]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox7->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox7->DropDownCount=30;
}

ComboBox7->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox7->Items->Add(Tab->FieldByName("Наименование воздействия")->AsString);
}        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox7KeyPress(TObject *Sender, char &Key)
{
if(Key==13)
{
String Text=ComboBox7->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер воздействия], [Наименование воздействия] From Воздействия Where [Наименование воздействия] Like '%"+Text+"%' AND Показ=true Order by [Номер воздействия]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox7->ItemIndex);
ComboBox7->Text=Tab->FieldByName("Наименование воздействия")->AsString;
if(ComboBox7->DroppedDown & ComboBox7->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("Номер воздействия")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("Наименование воздействия")->AsString;
InpVozd();
}
}        
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox7Select(TObject *Sender)
{
String Text=ComboBox7->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [Номер воздействия], [Наименование воздействия] From Воздействия Where [Наименование воздействия] Like '%"+Text+"%' AND Показ=true Order by [Номер воздействия]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox7->ItemIndex);
ComboBox7->Text=Tab->FieldByName("Наименование воздействия")->AsString;
if(!ComboBox7->DroppedDown & ComboBox7->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("Номер воздействия")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("Наименование воздействия")->AsString;
InpVozd();
}
}
//---------------------------------------------------------------------------

void __fastcall TFilter::ComboBox7Enter(TObject *Sender)
{
ComboBox7->DroppedDown=true;
}
//---------------------------------------------------------------------------
void TFilter::SetDefFiltr()
{
CText="SELECT  Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) ORDER BY Аспекты.[Номер аспекта];";
}
//----------------------------------------------------------------------------
