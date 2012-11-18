//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "FMoveAsp.h"
#include "MasterPointer.h"
#include "Zastavka.h"
#include "InpDocs.h"
#include "Progress.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TMAsp *MAsp;
//---------------------------------------------------------------------------
__fastcall TMAsp::TMAsp(TComponent* Owner)
        : TForm(Owner)
{
Cont=true;
}
//---------------------------------------------------------------------------
void __fastcall TMAsp::FormShow(TObject *Sender)
{
Prog->Close();
Podr->Active=false;
Podr->CommandText="Select * From Подразделения";
Podr->Connection=Zast->ADOAspect;
Podr->Active=true;
ComboBox1->Clear();
Podr->First();
for(int i=0;i<Podr->RecordCount;i++)
{
 AnsiString T=Podr->FieldByName("Название подразделения")->Value;
 ComboBox1->Items->Add(T);
 Podr->Next();
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

MoveAspects->Active=false;
CT="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение ";
CT=CT+" FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация ";
CT=CT+" Order by Аспекты.[Номер аспекта]";
MoveAspects->CommandText=CT;
MoveAspects->Connection=Zast->ADOAspect;
MoveAspects->Active=true;

ChangeCPodr();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::BitBtn1Click(TObject *Sender)
{
MoveAspects->First();
ChangeCPodr();
}
//---------------------------------------------------------------------------
void TMAsp::ChangeCPodr()
{
if(MoveAspects->RecordCount!=0)
{
CPodr->Enabled=true;
MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOAspect;
Podr->CommandText="Select * From Подразделения order by [Номер подразделения]";
Podr->Active=true;

CPodr->Clear();
for(Podr->First();!Podr->Eof;Podr->Next())
{
 CPodr->Items->Add(Podr->FieldByName("Название подразделения")->Value);
}

int N=MoveAspects->FieldByName("Подразделение")->AsInteger;

Podr->Locate("Номер подразделения", N, SO);
int Num=Podr->RecNo;

CPodr->ItemIndex=Num-1;
}
else
{
CPodr->Clear();
CPodr->Items->Add("Нет списка аспектов");
CPodr->ItemIndex=0;
CPodr->Enabled=false;
}

if(MoveAspects->RecNo>0)
{
Label8->Caption=MoveAspects->RecNo;
}
else
{
Label8->Caption=0;
}
Label10->Caption=MoveAspects->RecordCount;
}
//----------------------------------------------
void __fastcall TMAsp::BitBtn2Click(TObject *Sender)
{
MoveAspects->Prior();
ChangeCPodr();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::BitBtn3Click(TObject *Sender)
{
MoveAspects->Next();
ChangeCPodr();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::BitBtn4Click(TObject *Sender)
{
MoveAspects->Last();
ChangeCPodr();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::CPodrClick(TObject *Sender)
{
MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOAspect;
Podr->CommandText="Select * From Подразделения order by [Номер подразделения]";
Podr->Active=true;

Podr->First();
Podr->MoveBy(CPodr->ItemIndex);
int N=Podr->FieldByName("Номер подразделения")->Value;

MoveAspects->Edit();
MoveAspects->FieldByName("Подразделение")->Value=N;
MoveAspects->Post();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::BitBtn5Click(TObject *Sender)
{
Prog->SignComplete=true;
//Zast->SaveAspectsMSpec0->Execute();
Zast->BlockMK(true);
try
{
Zast->MClient->BlockServer("SaveAspectsMSpec0");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::BitBtn6Click(TObject *Sender)
{
if(Application->MessageBoxA("Вы действительно хотите прочесть список аспектов с сервера?\rВсе несохраненные изменения по перемещению аспектов будут удалены!","Чтение списка аспектов",MB_YESNO+MB_ICONWARNING+MB_DEFBUTTON2+MB_APPLMODAL)==IDYES)
{
Zast->BlockMK(true);
try
{
Zast->MClient->BlockServer("ReadAspectsMSpec");
}
catch(...)
{
 Zast->BlockMK(false);
}
}

}
//---------------------------------------------------------------------------
void TMAsp::MergeAspects()
{
Zast->MClient->WriteDiaryEvent("NetAspects","Начало объединения аспектов (главспец)","");

try
{

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * From Аспекты";
Comm->Execute();

Comm->CommandText="INSERT INTO Аспекты ( ServerNum, Подразделение, Ситуация, [Вид территории], Деятельность, Специальность, Аспект, Воздействие, Z, Значимость, [Проявление воздействия], [Тяжесть последствий], Приоритетность ) SELECT TempAspects.[Номер аспекта], TempAspects.Подразделение, TempAspects.Ситуация, TempAspects.[Вид территории], TempAspects.Деятельность, TempAspects.Специальность, TempAspects.Аспект, TempAspects.Воздействие, TempAspects.Z, TempAspects.Значимость, TempAspects.[Проявление воздействия], TempAspects.[Тяжесть последствий], TempAspects.Приоритетность FROM TempAspects;";
Comm->Execute();

MP<TADODataSet>Zn(this);
Zn->Connection=Zast->ADOAspect;
Zn->CommandText="select * From Значимость Order by [Мин граница]";
Zn->Active=true;

MP<TADODataSet>Asp(this);
Asp->Connection=Zast->ADOAspect;
Asp->CommandText="select * From Аспекты";
Asp->Active=true;

for(Asp->First();!Asp->Eof;Asp->Next())
{
 int Z=Asp->FieldByName("Z")->Value;
 bool Znak=false;

 for(Zn->First();!Zn->Eof;Zn->Next())
 {
  int Min=Zn->FieldByName("Мин граница")->Value;
  int Max=Zn->FieldByName("Макс граница")->Value;

  if(Z>=Min & Z<=Max)
  {
   Znak=Zn->FieldByName("Критерий")->Value;
   break;
  }
 }

 Asp->Edit();
 Asp->FieldByName("Значимость")->Value=Znak;
 Asp->Post();
}

MoveAspects->Active=false;
MoveAspects->Active=true;
ChangeCPodr();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка объединения аспектов (главспец)","Ошибка "+IntToStr(GetLastError()));

}
}
//----------------------------------------------
void __fastcall TMAsp::FormClose(TObject *Sender, TCloseAction &Action)
{
//Zast->BlockMK(true);
Zast->MClient->BlockServer("CloseMAsp");


}
//----------------------------------------------------
void __fastcall TMAsp::RadioGroup1Click(TObject *Sender)
{
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
 CText="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";

ComboBox4->Text="";
ComboBox5->Text="";
ComboBox6->Text="";
ComboBox7->Text="";
ComboBox1->ItemIndex=-1;
ComboBox2->Visible=false;
ComboBox3->Visible=false;
MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
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
 ComboBox5->Visible=false;
 ComboBox6->Visible=false;
 ComboBox7->Visible=false;
 Button1->Visible=true;
 Button2->Visible=false;
 Button3->Visible=false;
 Button4->Visible=false;
ComboBox2->Visible=false;
ComboBox3->Visible=false;

ComboBox4->DroppedDown=true;
ComboBox4->AutoComplete=false;
ComboBox4->SetFocus();

 break;
 }
 case 3:
 {
  ComboBox1->Visible=false;
 ComboBox4->Visible=false;
 ComboBox5->Visible=true;
 ComboBox6->Visible=false;
 ComboBox7->Visible=false;
 Button1->Visible=false;
 Button2->Visible=true;
 Button3->Visible=false;
 Button4->Visible=false;
ComboBox2->Visible=false;
ComboBox3->Visible=false;

ComboBox5->DroppedDown=true;
ComboBox5->AutoComplete=false;
ComboBox5->SetFocus();
 break;
 }
 case 4:
 {
  ComboBox1->Visible=false;
 ComboBox4->Visible=false;
 ComboBox5->Visible=false;
 ComboBox6->Visible=true;
 ComboBox7->Visible=false;
 Button1->Visible=false;
 Button2->Visible=false;
 Button3->Visible=true;
 Button4->Visible=false;
ComboBox2->Visible=false;
ComboBox3->Visible=false;

ComboBox6->DroppedDown=true;
ComboBox6->AutoComplete=false;
ComboBox6->SetFocus();
 break;
 }
 case 5:
 {
 ComboBox1->Visible=false;
 ComboBox4->Visible=false;
 ComboBox5->Visible=false;
 ComboBox6->Visible=false;
 ComboBox7->Visible=true;
 Button1->Visible=false;
 Button2->Visible=false;
 Button3->Visible=false;
 Button4->Visible=true;
 ComboBox2->Visible=false;
ComboBox3->Visible=false;

ComboBox7->DroppedDown=true;
ComboBox7->AutoComplete=false;
ComboBox7->SetFocus();
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

void __fastcall TMAsp::Button1Click(TObject *Sender)
{

InputDocs->IForm=2;
InputDocs->Mode=1;
InputDocs->ShowModal();

}
//---------------------------------------------------------------------------

void __fastcall TMAsp::Button2Click(TObject *Sender)
{

InputDocs->IForm=2;
InputDocs->Mode=2;
InputDocs->ShowModal();

}
//---------------------------------------------------------------------------

void __fastcall TMAsp::Button3Click(TObject *Sender)
{

InputDocs->IForm=2;
InputDocs->Mode=3;
InputDocs->ShowModal();

}
//---------------------------------------------------------------------------

void __fastcall TMAsp::Button4Click(TObject *Sender)
{

InputDocs->IForm=2;
InputDocs->Mode=4;
InputDocs->ShowModal();

}
//---------------------------------------------------------------------------
void TMAsp::InpTer()
{

String S="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";
CText=S+" Where [Вид территории]="+IntToStr(Index)+" Order By Аспекты.[Номер аспекта]";


MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
}
//------------
void TMAsp::InpDeyat()
{
String S="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";
CText=S+" Where Деятельность="+IntToStr(Index)+" Order By Аспекты.[Номер аспекта]";

MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
}
//------------
void TMAsp::InpAsp()
{
String S="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";
CText=S+" Where Аспект="+IntToStr(Index)+" Order By Аспекты.[Номер аспекта]";

MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
}
//------------
void TMAsp::InpVozd()
{
String S="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";
CText=S+" Where Воздействие="+IntToStr(Index)+" Order By Аспекты.[Номер аспекта]";

MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
}
//------------


void __fastcall TMAsp::ComboBox1Click(TObject *Sender)
{
Podr->First();
Podr->MoveBy(ComboBox1->ItemIndex);

Index=Podr->FieldByName("Номер подразделения")->Value;
String S="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";
CText=S+" Where Подразделение="+IntToStr(Index)+" Order By Аспекты.[Номер аспекта]";

MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::ComboBox2Change(TObject *Sender)
{
String S="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";

if (ComboBox2->ItemIndex==0)
{

CText=S+" Where Подразделение<>0 AND Деятельность<>0 AND Аспект<>0 AND Воздействие<>0 Order By Аспекты.[Номер аспекта]";

}
else
{
CText=S+" Where Подразделение=0 OR Деятельность=0 OR Аспект=0 OR Воздействие=0 Order By Аспекты.[Номер аспекта]";

}
MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::ComboBox3Change(TObject *Sender)
{
String S="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";

if (ComboBox3->ItemIndex==0)
{
CText=S+" Where Значимость=True ORDER BY Аспекты.[Номер аспекта]";



}
else
{
CText=S+" Where Значимость=False ORDER BY Аспекты.[Номер аспекта]";



}

MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::Edit1DblClick(TObject *Sender)
{
Button1->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::Edit2DblClick(TObject *Sender)
{
Button2->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::Edit3DblClick(TObject *Sender)
{
Button3->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::Edit4DblClick(TObject *Sender)
{
Button4->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::Timer1Timer(TObject *Sender)
{
Timer1->Enabled=false;
this->Close();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
  Application->HelpJump("IDH_ДВИЖЕНИЕ_АСПЕКТОВ");
}
}
//---------------------------------------------------------------------------





void __fastcall TMAsp::ComboBox4DropDown(TObject *Sender)
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
void __fastcall TMAsp::ComboBox4Change(TObject *Sender)
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

void __fastcall TMAsp::ComboBox4Select(TObject *Sender)
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
MAsp->Index=Tab->FieldByName("Номер территории")->AsInteger;
InpTer();
}
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::ComboBox4KeyPress(TObject *Sender, char &Key)
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
MAsp->Index=Tab->FieldByName("Номер территории")->AsInteger;
InpTer();
}
}

}
//---------------------------------------------------------------------------

void __fastcall TMAsp::ComboBox5Change(TObject *Sender)
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

void __fastcall TMAsp::ComboBox5DropDown(TObject *Sender)
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

void __fastcall TMAsp::ComboBox5KeyPress(TObject *Sender, char &Key)
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
MAsp->Index=Tab->FieldByName("Номер деятельности")->AsInteger;
InpDeyat();
}
}
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::ComboBox5Select(TObject *Sender)
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
MAsp->Index=Tab->FieldByName("Номер деятельности")->AsInteger;
InpDeyat();
}
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::ComboBox6Change(TObject *Sender)
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

void __fastcall TMAsp::ComboBox6DropDown(TObject *Sender)
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

void __fastcall TMAsp::ComboBox6KeyPress(TObject *Sender, char &Key)
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
MAsp->Index=Tab->FieldByName("Номер аспекта")->AsInteger;
InpAsp();
}
}
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::ComboBox6Select(TObject *Sender)
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
MAsp->Index=Tab->FieldByName("Номер аспекта")->AsInteger;
InpAsp();
}
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::ComboBox7Change(TObject *Sender)
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

void __fastcall TMAsp::ComboBox7DropDown(TObject *Sender)
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

void __fastcall TMAsp::ComboBox7KeyPress(TObject *Sender, char &Key)
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
MAsp->Index=Tab->FieldByName("Номер воздействия")->AsInteger;
InpVozd();
}
}
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::ComboBox7Select(TObject *Sender)
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
MAsp->Index=Tab->FieldByName("Номер воздействия")->AsInteger;
InpVozd();
}
}
//---------------------------------------------------------------------------

