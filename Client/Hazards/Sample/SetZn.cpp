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
 DataSet->FieldByName("��������1")->Value="��";
 DataSet->FieldByName("��������")->Value=true;
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
AnsiString A=DataSet->FieldByName("��������1")->Value;
if(A=="��")
{
 DataSet->FieldByName("��������")->Value=true;
 Kriteriy=true;
}
else
{
 DataSet->FieldByName("��������")->Value=false;
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
 DataSet->FieldByName("��������1")->Value="���";
 DataSet->FieldByName("��������")->Value=false;
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
ADODataSet1->CommandText="select * from ���������� Order By [��� �������]";
ADODataSet1->Active=true;
DataSet->Last();



int Min=DataSet->FieldByName("��� �������")->Value;
bool Cr=DataSet->FieldByName("��������")->Value;

DataSet->Edit();
DataSet->FieldByName("���� �������")->Value=Min+1;
DataSet->Post();

DataSet->Prior();
for(int i=1;i<DataSet->RecordCount;i++)
{
DataSet->Edit();
DataSet->FieldByName("���� �������")->Value=Min-1;
if(Cr==false)
{
DataSet->FieldByName("��������")->Value=false;
DataSet->FieldByName("��������1")->Value="���";

}
DataSet->Post();
Min=DataSet->FieldByName("��� �������")->Value;
Cr=DataSet->FieldByName("��������")->Value;
DataSet->Prior();
}
DataSet->First();
DataSet->Edit();
DataSet->FieldByName("��� �������")->Value=0;
DataSet->Post();
ADODataSet1->Active=false;
ADODataSet1->CommandText="select * from ���������� Order By [��� �������]";
ADODataSet1->Active=true;
DataSet->First();

//bool K=false;
for(int i=0;i<DataSet->RecordCount;i++)
{
if(DataSet->FieldByName("��� �������")->Value>DataSet->FieldByName("���� �������")->Value)
{
 ShowMessage("� ������ ��� ���� ���������� ������� ������ "+DataSet->FieldByName("��� �������")->Value+" ��������� ��");
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
ADODataSet1->CommandText="select * from ���������� Order By [��� �������]";
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
ADODataSet1->CommandText="select * from ���������� Order By [��� �������]";
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
// ��������
Ins=true;
ADODataSet1->Last();
int Max=ADODataSet1->FieldByName("���� �������")->Value;
bool Cr=ADODataSet1->FieldByName("��������")->Value;
ADODataSet1->Append();
ADODataSet1->FieldByName("��� �������")->Value=Max;
ADODataSet1->FieldByName("���� �������")->Value=Max+1;
ADODataSet1->FieldByName("��������")->Value=Cr;
ADODataSet1->FieldByName("������������ ����������")->Value="����� ����������";
ADODataSet1->FieldByName("����������� ����")->Value="����� ����������� ����";

if (Cr==true)
{
ADODataSet1->FieldByName("��������1")->Value="��";
}
else
{
ADODataSet1->FieldByName("��������1")->Value="���";
}

ADODataSet1->Post();
ADODataSet1->Last();
Ins=false;
}
//---------------------------------------------------------------------------

void __fastcall TSZn::N2Click(TObject *Sender)
{
// �������

ADODataSet1->Delete();
ADODataSet2->Active=false;
ADODataSet2->CommandText="Select * From ���������� Order By [��� �������]";
ADODataSet2->Active=true;
ADODataSet2->Last();

int Min=ADODataSet2->FieldByName("��� �������")->Value;

ADODataSet2->Prior();
for(int i=1;i<ADODataSet2->RecordCount;i++)
{
ADODataSet2->Edit();
// int Min=ADODataSet1->FieldByName("��� �������")->Value;


ADODataSet2->FieldByName("���� �������")->Value=Min-1;
Min=ADODataSet2->FieldByName("��� �������")->Value;
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
 ShowMessage("��������� �� ������ ������, ��������� ���������!");
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
// ��������


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




