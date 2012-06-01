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
//DBAspects->Connection=Zast->ADOAspect;
//DBAspects->CommandText="SELECT Аспекты.Подразделение, Подразделения.[Название подразделения], Деятельность.[Наименование деятельности], Аспект.[Наименование аспекта], Воздействия.[Наименование воздействия], Ситуации.[Название ситуации], Аспекты.Z FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация ORDER BY Подразделения.[Название подразделения], Аспекты.[Номер аспекта];";
//DBAspects->Active=true;

}
//---------------------------------------------------------------------------
void __fastcall TMAsp::FormShow(TObject *Sender)
{
String R="Завершено";
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

Zast->MClient->Start();
Zast->MClient->RegForm(this);


Table* STab=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

STab->SetCommandText("SELECT Аспекты.[Номер аспекта], Аспекты.Ситуация, Аспекты.[Вид территории], Аспекты.[Деятельность], Аспекты.[Специальность], Аспекты.[Аспект], Аспекты.[Воздействие], Аспекты.[Z], Аспекты.[Значимость]  FROM Аспекты order by [Номер аспекта]");
STab->Active(true);

int SN=STab->RecordCount();

MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
Tab->CommandText="SELECT Аспекты.[Номер аспекта], Аспекты.Ситуация, Аспекты.[Вид территории], Аспекты.[Деятельность], Аспекты.[Специальность], Аспекты.[Аспект], Аспекты.[Воздействие], Аспекты.[Z], Аспекты.[Значимость]  FROM Аспекты order by [Номер аспекта]";
Tab->Active=true;

int CN=Tab->RecordCount;

if(CN!=SN)
{
if( Application->MessageBoxA("Количество аспектов на сервере не совпадает с количеством аспектов в локальной базе данных\rОбновить список аспектов?","Обновление аспектов",MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON1)==IDYES)
{
Zast->MClient->WriteDiaryEvent("NetAspects","Начало обновления аспектов (главспец)","По количеству");
R=Documents->LoadAspects();
}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects","Отказ от обновления аспектов (главспец)","По количеству");
}
}
else
{
if(Zast->MClient->VerifyTable(STab, Tab)!=0)
{
if( Application->MessageBoxA("Содержание аспектов на сервере не совпадает с содержанием аспектов в локальной базе данных\rОбновить список аспектов?","Обновление аспектов",MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON1)==IDYES)
{
Zast->MClient->WriteDiaryEvent("NetAspects","Начало обновления аспектов (главспец)","По содержанию");
R=Documents->LoadAspects();
}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects","Отказ от обновления аспектов (главспец)","По содержанию");
}
}
}

Zast->MClient->DeleteTable(this, STab);

if(R!="Завершено")
{
 ShowMessage(R);
 Cont=false;
 Timer1->Enabled=true;
}



Zast->MClient->Stop();

MoveAspects->Active=false;
MoveAspects->CommandText="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация; ";
MoveAspects->Connection=Zast->ADOAspect;
MoveAspects->Active=true;



ChangeCPodr();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::BitBtn1Click(TObject *Sender)
{
//DataSetFirst1->Execute();
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
//DataSetPrior1->Execute();
MoveAspects->Prior();
ChangeCPodr();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::BitBtn3Click(TObject *Sender)
{
//DataSetNext1->Execute();
MoveAspects->Next();
ChangeCPodr();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::BitBtn4Click(TObject *Sender)
{
//DataSetLast1->Execute();
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
//Запись аспектов
Zast->MClient->Start();
FProg->Label1->Caption="Сохранение аспектов";
ShowMessage(SaveAspects());
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::BitBtn6Click(TObject *Sender)
{
//Чтение аспектов
if(Application->MessageBoxA("Вы действительно хотите прочесть список аспектов с сервера?\rВсе несохраненные изменения по перемещению аспектов будут удалены!","Чтение списка аспектов",MB_YESNO+MB_ICONWARNING+MB_DEFBUTTON2+MB_APPLMODAL)==IDYES)
{
Zast->MClient->WriteDiaryEvent("NetAspects","Начало загрузки аспектов (главспец)","");
Zast->MClient->Start();
FProg->Label1->Caption="Загрузка аспектов";
ShowMessage(Documents->LoadAspects());
Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки аспектов (главспец)","");
Zast->MClient->Stop();
}
MoveAspects->Active=false;
MoveAspects->Active=true;
ChangeCPodr();
}
//---------------------------------------------------------------------------
String TMAsp::LoadAspects()
{
String Res="Ошибка передачи";

Zast->MClient->WriteDiaryEvent("NetAspects","Начало загрузки аспектов (главспец)","");

try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * From TempAspects";
Comm->Execute();

MP<TADODataSet>CAsp(this);
CAsp->Connection=Zast->ADOAspect;
CAsp->CommandText="SELECT TempAspects.[Номер аспекта], TempAspects.Подразделение, TempAspects.Ситуация, TempAspects.[Вид территории], TempAspects.Деятельность, TempAspects.Специальность, TempAspects.Аспект, TempAspects.Воздействие, TempAspects.Z, TempAspects.Значимость, TempAspects.[Проявление воздействия], TempAspects.[Тяжесть последствий], TempAspects.Приоритетность FROM TempAspects ORDER BY TempAspects.[Номер аспекта];";
CAsp->Active=true;

Table* SAsp=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

SAsp->SetCommandText("SELECT Аспекты.[Номер аспекта], Аспекты.Подразделение, Аспекты.Ситуация, Аспекты.[Вид территории], Аспекты.Деятельность, Аспекты.Специальность, Аспекты.Аспект, Аспекты.Воздействие, Аспекты.Z, Аспекты.Значимость, Аспекты.[Проявление воздействия], Аспекты.[Тяжесть последствий], Аспекты.Приоритетность FROM Аспекты ORDER BY Аспекты.[Номер аспекта];");
SAsp->Active(true);

Zast->MClient->LoadTable(SAsp, CAsp, FProg->Label1, FProg->Progress);

if(Zast->MClient->VerifyTable(SAsp, CAsp)==0)
{

MergeAspects();

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки аспектов (главспец)","");
Res="Завершено";
}
Zast->MClient->DeleteTable(this, SAsp);

if(Res!="Завершено")
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой загрузки аспектов (главспец)","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка загрузки аспектов (главспец)","Ошибка "+IntToStr(GetLastError()));
}
return Res;
}
//-----------------------------------------------------------------------------
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
if(Cont==true)
{
DataSetRefresh1->Execute();
DataSetPost1->Execute();

//Проверка движения аспектов
Zast->MClient->Start();
MP<TADODataSet>CAsp(this);
CAsp->Connection=Zast->ADOAspect;
CAsp->CommandText="SELECT Аспекты.[Номер аспекта], Подразделения.ServerNum FROM Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение ORDER BY Аспекты.[Номер аспекта];";
CAsp->Active=true;

Table* SAsp=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

SAsp->SetCommandText("SELECT Аспекты.[Номер аспекта], Аспекты.Подразделение FROM Аспекты order by Аспекты.[номер аспекта]; ");
SAsp->Active(true);

if(Zast->MClient->VerifyTable(SAsp, CAsp)!=0)
{
if(Application->MessageBoxA("Движение аспектов не сохранено.\rСохранить?","Проверка",MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON1)==IDYES)
{
SaveAspects();
}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects","Отказ от сохранения движения аспектов аспектов (главспец)","");
}
}
Zast->MClient->DeleteTable(this, SAsp);
Zast->MClient->Stop();
}
}
//---------------------------------------------------------------------------
String TMAsp::SaveAspects()
{
String Ret="Ошибка передачи";
Zast->MClient->WriteDiaryEvent("NetAspects","Начало сохранения аспектов (главспец)","");

try
{
Zast->MClient->SetCommandText("Аспекты","Delete * From TempAspects");
Zast->MClient->CommandExec("Аспекты");

MP<TADODataSet>CAsp(this);
CAsp->Connection=Zast->ADOAspect;
CAsp->CommandText="SELECT Аспекты.[Номер аспекта], Подразделения.ServerNum FROM Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение Order by Аспекты.[Номер аспекта]; ";
CAsp->Active=true;

Table* SAsp=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

SAsp->SetCommandText("SELECT TempAspects.[Номер аспекта], TempAspects.Подразделение From TempAspects order by [Номер аспекта];");
SAsp->Active(true);

Zast->MClient->LoadTable(CAsp, SAsp, FProg->Label1, FProg->Progress);

if(Zast->MClient->VerifyTable(SAsp, CAsp)==0)
{
Zast->MClient->MergeAspectsMainSpec();
Zast->MClient->WriteDiaryEvent("NetAspects","Конец сохранения аспектов (главспец)","");
Ret="Завершено";
}
Zast->MClient->DeleteTable(this, SAsp);

if(Ret!="Завершено")
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой сохранения аспектов (главспец)","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка сохранения аспектов (главспец)","Ошибка "+IntToStr(GetLastError()));
}
return Ret;
}
//----------------------------------------------------
void __fastcall TMAsp::RadioGroup1Click(TObject *Sender)
{
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
 CText="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";

//  Button5->Enabled=true;
Edit1->Text="";
Edit2->Text="";
Edit3->Text="";
Edit4->Text="";
ComboBox1->ItemIndex=-1;
ComboBox2->Visible=false;
ComboBox3->Visible=false;
//Filtr1="select * from Аспекты";
//Filtr2="SELECT Аспекты.* FROM Аспекты WHERE (((Значимость)=True))";
MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
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
Edit1->Text=InputDocs->TextBr;
Index=InputDocs->NumBr;
String S="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";
CText=S+" Where [Вид территории]="+IntToStr(Index)+" Order By Аспекты.[Номер аспекта]";


MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
/*
Form1->Aspects->Active=false;
Form1->Aspects->CommandText=CText;
Form1->Aspects->Active=true;
 Button5->Enabled=true;
 */
}
//------------
void TMAsp::InpDeyat()
{
Edit2->Text=InputDocs->TextBr;
Index=InputDocs->NumBr;
String S="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";
CText=S+" Where Деятельность="+IntToStr(Index)+" Order By Аспекты.[Номер аспекта]";

MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();

/*
Form1->Aspects->Active=false;
Form1->Aspects->CommandText=CText;
Form1->Aspects->Active=true;
 Button5->Enabled=true;
 */
}
//------------
void TMAsp::InpAsp()
{
Edit3->Text=InputDocs->TextBr;
Index=InputDocs->NumBr;
String S="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";
CText=S+" Where Аспект="+IntToStr(Index)+" Order By Аспекты.[Номер аспекта]";

MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
/*
Form1->Aspects->Active=false;
Form1->Aspects->CommandText=CText;
Form1->Aspects->Active=true;
 Button5->Enabled=true;
 */
}
//------------
void TMAsp::InpVozd()
{
Edit4->Text=InputDocs->TextBr;
Index=InputDocs->NumBr;
String S="SELECT Аспекты.[Номер аспекта], Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Территория.[Номер территории], Территория.[Наименование территории], Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Аспекты.Значимость, Аспекты.Z, Ситуации.[Номер ситуации], Ситуации.[Название ситуации], Аспекты.Подразделение FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Территория.[Номер территории] = Аспекты.[Вид территории]) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация";
CText=S+" Where Воздействие="+IntToStr(Index)+" Order By Аспекты.[Номер аспекта]";

MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
/*
Form1->Aspects->Active=false;
Form1->Aspects->CommandText=CText;
Form1->Aspects->Active=true;
 Button5->Enabled=true;
 */
}
//------------


void __fastcall TMAsp::ComboBox1Click(TObject *Sender)
{
//ComboBox1->ItemIndex=CPodr->ItemIndex;
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

