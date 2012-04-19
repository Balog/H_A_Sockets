//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "SetZn.h"
#include "Zastavka.h"
//#include "MainForm.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TSZn *SZn;
//---------------------------------------------------------------------------
__fastcall TSZn::TSZn(TComponent* Owner)
        : TForm(Owner)
{
C=false;
Ins=false;
}
//---------------------------------------------------------------------------
void TSZn::Crit(TDataSet *DataSet)
{
int NR=DataSet->RecNo;

while(!DataSet->Eof)
{
 DataSet->Edit();
 DataSet->FieldByName("Критерий1")->Value="Да";
 DataSet->FieldByName("Критерий")->Value=true;
 DataSet->Post();
 DataSet->Next();
}
DataSet->First();
DataSet->MoveBy(NR-1);
}
//----------------------------------------------------

void __fastcall TSZn::ADODataSet1BeforePost(TDataSet *DataSet)
{

if(C==false)
{
AnsiString A=DataSet->FieldByName("Критерий1")->Value;
if(A=="Да")
{
 DataSet->FieldByName("Критерий")->Value=true;
 Kriteriy=true;
}
else
{
 DataSet->FieldByName("Критерий")->Value=false;
 Kriteriy=false;
}

}

}
//---------------------------
void TSZn::Crit1(TDataSet *DataSet)
{

int NR=DataSet->RecNo;

while(!DataSet->Bof)
{
 DataSet->Edit();
 DataSet->FieldByName("Критерий1")->Value="Нет";
 DataSet->FieldByName("Критерий")->Value=false;
 DataSet->Post();
 DataSet->Prior();
}
DataSet->First();
DataSet->MoveBy(NR-1);
}

//---------------------------
void TSZn::Crit2(TDataSet *DataSet)
{
int NR=DataSet->RecNo;
ADODataSet1->Active=false;
ADODataSet1->CommandText="select * from Значимость Order By [Мин граница]";
ADODataSet1->Active=true;
DataSet->Last();



int Min=DataSet->FieldByName("Мин граница")->Value;
bool Cr=DataSet->FieldByName("Критерий")->Value;

DataSet->Edit();
DataSet->FieldByName("Макс граница")->Value=Min+1;
DataSet->Post();

DataSet->Prior();
for(int i=1;i<DataSet->RecordCount;i++)
{
DataSet->Edit();
DataSet->FieldByName("Макс граница")->Value=Min-1;
if(Cr==false)
{
DataSet->FieldByName("Критерий")->Value=false;
DataSet->FieldByName("Критерий1")->Value="Нет";

}
DataSet->Post();
Min=DataSet->FieldByName("Мин граница")->Value;
Cr=DataSet->FieldByName("Критерий")->Value;
DataSet->Prior();
}
DataSet->First();
DataSet->Edit();
DataSet->FieldByName("Мин граница")->Value=0;
DataSet->Post();
ADODataSet1->Active=false;
ADODataSet1->CommandText="select * from Значимость Order By [Мин граница]";
ADODataSet1->Active=true;
DataSet->First();

//bool K=false;
for(int i=0;i<DataSet->RecordCount;i++)
{
if(DataSet->FieldByName("Мин граница")->Value>DataSet->FieldByName("Макс граница")->Value)
{
 ShowMessage("В списке уже есть минмальная граница равная "+DataSet->FieldByName("Мин граница")->Value+" исправьте ее");
//K=true;

break;
}
ADODataSet1->Next();
}

DataSet->First();
DataSet->MoveBy(NR-1);



}
//---------------------------------------------------------------------------
void __fastcall TSZn::ADODataSet1AfterPost(TDataSet *DataSet)
{

if(C==false)
{

 /*
ADODataSet1->Active=false;
ADODataSet1->CommandText="select * from Значимость Order By [Мин граница]";
ADODataSet1->Active=true;
 */
bool B=Kriteriy;
if(B==true)
{
C=true;
Crit(DataSet);
C=false;
}
else
{
C=true;
Crit1(DataSet);
C=false;
}

C=true;
Crit2(DataSet);
C=false;
}
Button1->Enabled=true;
}

//---------------------------------------------------------------------------
void __fastcall TSZn::DataSource1UpdateData(TObject *Sender)
{
//

}
//---------------------------------------------------------------------------

void __fastcall TSZn::ADODataSet1BeforeEdit(TDataSet *DataSet)
{
//
Button1->Enabled=false;



}

//---------------------------------------------------------------------------




void __fastcall TSZn::FormShow(TObject *Sender)
{
ADODataSet1->Active=false;
ADODataSet1->CommandText="select * from Значимость Order By [Мин граница]";
ADODataSet1->Active=true;

}
//---------------------------------------------------------------------------

void __fastcall TSZn::Button1Click(TObject *Sender)
{
this->Close();
}
//---------------------------------------------------------------------------

void __fastcall TSZn::N1Click(TObject *Sender)
{
// Добавить
Ins=true;
ADODataSet1->Last();
int Max=ADODataSet1->FieldByName("Макс граница")->Value;
bool Cr=ADODataSet1->FieldByName("Критерий")->Value;
ADODataSet1->Append();
ADODataSet1->FieldByName("Мин граница")->Value=Max;
ADODataSet1->FieldByName("Макс граница")->Value=Max+1;
ADODataSet1->FieldByName("Критерий")->Value=Cr;
ADODataSet1->FieldByName("Наименование значимости")->Value="Новая значимость";
ADODataSet1->FieldByName("Необходимая мера")->Value="Новая необходимая мера";

if (Cr==true)
{
ADODataSet1->FieldByName("Критерий1")->Value="Да";
}
else
{
ADODataSet1->FieldByName("Критерий1")->Value="Нет";
}

ADODataSet1->Post();
ADODataSet1->Last();
Ins=false;
}
//---------------------------------------------------------------------------

void __fastcall TSZn::N2Click(TObject *Sender)
{
// Удалить

ADODataSet1->Delete();
ADODataSet2->Active=false;
ADODataSet2->CommandText="Select * From Значимость Order By [Мин граница]";
ADODataSet2->Active=true;
ADODataSet2->Last();

int Min=ADODataSet2->FieldByName("Мин граница")->Value;

ADODataSet2->Prior();
for(int i=1;i<ADODataSet2->RecordCount;i++)
{
ADODataSet2->Edit();
// int Min=ADODataSet1->FieldByName("Мин граница")->Value;


ADODataSet2->FieldByName("Макс граница")->Value=Min-1;
Min=ADODataSet2->FieldByName("Мин граница")->Value;
ADODataSet2->Post();
ADODataSet2->Prior();
}
ADODataSet1->Active=false;
ADODataSet1->Active=true;
}
//---------------------------------------------------------------------------

void __fastcall TSZn::FormClose(TObject *Sender, TCloseAction &Action)
{
/*
Zast->CopyZn();


Form1->InitCombo();
Form1->Calc();
*/
}
//---------------------------------------------------------------------------

void __fastcall TSZn::FormCloseQuery(TObject *Sender, bool &CanClose)
{

CanClose=Button1->Enabled;
if(!CanClose)
{
 ShowMessage("Перейдите на другую строку, сохраните результат!");
}

}
//---------------------------------------------------------------------------



void __fastcall TSZn::PopupMenu1Popup(TObject *Sender)
{
if(ADODataSet1->RecordCount>2)
{
  PopupMenu1->Items->Items[0]->Enabled=true;
  PopupMenu1->Items->Items[1]->Enabled=true;
}
else
{
  PopupMenu1->Items->Items[0]->Enabled=true;
  PopupMenu1->Items->Items[1]->Enabled=false;
}
}
//---------------------------------------------------------------------------

void __fastcall TSZn::FormActivate(TObject *Sender)
{
this->Position=poDesktopCenter;        
}
//---------------------------------------------------------------------------

void __fastcall TSZn::ADODataSet1BeforeInsert(TDataSet *DataSet)
{
//
// Добавить


}
//---------------------------------------------------------------------------

void __fastcall TSZn::ADODataSet1AfterInsert(TDataSet *DataSet)
{
if(Ins==false)
{
DataSet->Prior();
DataSet->Next();
}
}
//---------------------------------------------------------------------------




