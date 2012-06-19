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
Podr->Active=false;
Podr->CommandText="Select * From �������������";
Podr->Connection=Zast->ADOAspect;
Podr->Active=true;
ComboBox1->Clear();
Podr->First();
for(int i=0;i<Podr->RecordCount;i++)
{
 AnsiString T=Podr->FieldByName("�������� �������������")->Value;
 ComboBox1->Items->Add(T);
 Podr->Next();
}

String Text="";
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
Tab->CommandText="Select [����� ����������], [������������ ����������] From ���������� Order by [����� ����������]";//"Select [����� ����������], [������������ ����������] From ���������� Where [������������ ����������] Like '%�%' Order by [����� ����������]";
Tab->Active=true;

ComboBox4->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox4->Items->Add(Tab->FieldByName("������������ ����������")->AsString);
}

Tab->Active=false;
String CT="Select [����� ������������], [������������ ������������] From ������������ Where [������������ ������������] Like '%"+Text+"%' AND �����=true Order by [����� ������������]";
Tab->CommandText=CT;
Tab->Active=true;
ComboBox5->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox5->Items->Add(Tab->FieldByName("������������ ������������")->AsString);
}

Tab->Active=false;
CT="Select [����� �������], [������������ �������] From ������ Where [������������ �������] Like '%"+Text+"%' AND �����=true Order by [����� �������]";
Tab->CommandText=CT;
Tab->Active=true;
ComboBox6->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox6->Items->Add(Tab->FieldByName("������������ �������")->AsString);
}

Tab->Active=false;
CT="Select [����� �����������], [������������ �����������] From ����������� Where [������������ �����������] Like '%"+Text+"%' AND �����=true Order by [����� �����������]";
Tab->CommandText=CT;
Tab->Active=true;
ComboBox7->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox7->Items->Add(Tab->FieldByName("������������ �����������")->AsString);
}

MoveAspects->Active=false;
CT="SELECT �������.[����� �������], �������������.[����� �������������], �������������.[�������� �������������], ����������.[����� ����������], ����������.[������������ ����������], ������������.[����� ������������], ������������.[������������ ������������], ������.[����� �������], ������.[������������ �������], �����������.[����� �����������], �����������.[������������ �����������], �������.����������, �������.Z, ��������.[����� ��������], ��������.[�������� ��������], �������.������������� ";
CT=CT+" FROM �������� INNER JOIN (����������� INNER JOIN (������ INNER JOIN (������������ INNER JOIN (���������� INNER JOIN (������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) ON ����������.[����� ����������] = �������.[��� ����������]) ON ������������.[����� ������������] = �������.������������) ON ������.[����� �������] = �������.������) ON �����������.[����� �����������] = �������.�����������) ON ��������.[����� ��������] = �������.�������� ";
CT=CT+" Order by �������.[����� �������]";
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
Podr->CommandText="Select * From ������������� order by [����� �������������]";
Podr->Active=true;

CPodr->Clear();
for(Podr->First();!Podr->Eof;Podr->Next())
{
 CPodr->Items->Add(Podr->FieldByName("�������� �������������")->Value);
}

int N=MoveAspects->FieldByName("�������������")->AsInteger;

Podr->Locate("����� �������������", N, SO);
int Num=Podr->RecNo;

CPodr->ItemIndex=Num-1;
}
else
{
CPodr->Clear();
CPodr->Items->Add("��� ������ ��������");
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
Podr->CommandText="Select * From ������������� order by [����� �������������]";
Podr->Active=true;

Podr->First();
Podr->MoveBy(CPodr->ItemIndex);
int N=Podr->FieldByName("����� �������������")->Value;

MoveAspects->Edit();
MoveAspects->FieldByName("�������������")->Value=N;
MoveAspects->Post();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::BitBtn5Click(TObject *Sender)
{
//Zast->SaveAspectsMSpec0->Execute();
Zast->MClient->BlockServer("SaveAspectsMSpec0");
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::BitBtn6Click(TObject *Sender)
{
if(Application->MessageBoxA("�� ������������� ������ �������� ������ �������� � �������?\r��� ������������� ��������� �� ����������� �������� ����� �������!","������ ������ ��������",MB_YESNO+MB_ICONWARNING+MB_DEFBUTTON2+MB_APPLMODAL)==IDYES)
{
Zast->MClient->BlockServer("ReadAspectsMSpec");
}

}
//---------------------------------------------------------------------------
void TMAsp::MergeAspects()
{
Zast->MClient->WriteDiaryEvent("NetAspects","������ ����������� �������� (��������)","");

try
{

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * From �������";
Comm->Execute();

Comm->CommandText="INSERT INTO ������� ( ServerNum, �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, Z, ����������, [���������� �����������], [������� �����������], �������������� ) SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.�������������� FROM TempAspects;";
Comm->Execute();

MP<TADODataSet>Zn(this);
Zn->Connection=Zast->ADOAspect;
Zn->CommandText="select * From ���������� Order by [��� �������]";
Zn->Active=true;

MP<TADODataSet>Asp(this);
Asp->Connection=Zast->ADOAspect;
Asp->CommandText="select * From �������";
Asp->Active=true;

for(Asp->First();!Asp->Eof;Asp->Next())
{
 int Z=Asp->FieldByName("Z")->Value;
 bool Znak=false;

 for(Zn->First();!Zn->Eof;Zn->Next())
 {
  int Min=Zn->FieldByName("��� �������")->Value;
  int Max=Zn->FieldByName("���� �������")->Value;

  if(Z>=Min & Z<=Max)
  {
   Znak=Zn->FieldByName("��������")->Value;
   break;
  }
 }

 Asp->Edit();
 Asp->FieldByName("����������")->Value=Znak;
 Asp->Post();
}

MoveAspects->Active=false;
MoveAspects->Active=true;
ChangeCPodr();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ����������� �������� (��������)","������ "+IntToStr(GetLastError()));

}
}
//----------------------------------------------
void __fastcall TMAsp::FormClose(TObject *Sender, TCloseAction &Action)
{


 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("CompareMSpecAspects2");
 String ServerSQL="SELECT �������.[����� �������], �������.������������� FROM ������� order by �������.[����� �������];";
 String ClientSQL="SELECT TempAspects.[����� �������], TempAspects.������������� FROM TempAspects;";
Zast->MClient->ReadTable("�������",ServerSQL, "�������", ClientSQL);
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
 CText="SELECT �������.[����� �������], �������������.[����� �������������], �������������.[�������� �������������], ����������.[����� ����������], ����������.[������������ ����������], ������������.[����� ������������], ������������.[������������ ������������], ������.[����� �������], ������.[������������ �������], �����������.[����� �����������], �����������.[������������ �����������], �������.����������, �������.Z, ��������.[����� ��������], ��������.[�������� ��������], �������.������������� FROM �������� INNER JOIN (����������� INNER JOIN (������ INNER JOIN (������������ INNER JOIN (���������� INNER JOIN (������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) ON ����������.[����� ����������] = �������.[��� ����������]) ON ������������.[����� ������������] = �������.������������) ON ������.[����� �������] = �������.������) ON �����������.[����� �����������] = �������.�����������) ON ��������.[����� ��������] = �������.��������";

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

String S="SELECT �������.[����� �������], �������������.[����� �������������], �������������.[�������� �������������], ����������.[����� ����������], ����������.[������������ ����������], ������������.[����� ������������], ������������.[������������ ������������], ������.[����� �������], ������.[������������ �������], �����������.[����� �����������], �����������.[������������ �����������], �������.����������, �������.Z, ��������.[����� ��������], ��������.[�������� ��������], �������.������������� FROM �������� INNER JOIN (����������� INNER JOIN (������ INNER JOIN (������������ INNER JOIN (���������� INNER JOIN (������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) ON ����������.[����� ����������] = �������.[��� ����������]) ON ������������.[����� ������������] = �������.������������) ON ������.[����� �������] = �������.������) ON �����������.[����� �����������] = �������.�����������) ON ��������.[����� ��������] = �������.��������";
CText=S+" Where [��� ����������]="+IntToStr(Index)+" Order By �������.[����� �������]";


MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
}
//------------
void TMAsp::InpDeyat()
{
String S="SELECT �������.[����� �������], �������������.[����� �������������], �������������.[�������� �������������], ����������.[����� ����������], ����������.[������������ ����������], ������������.[����� ������������], ������������.[������������ ������������], ������.[����� �������], ������.[������������ �������], �����������.[����� �����������], �����������.[������������ �����������], �������.����������, �������.Z, ��������.[����� ��������], ��������.[�������� ��������], �������.������������� FROM �������� INNER JOIN (����������� INNER JOIN (������ INNER JOIN (������������ INNER JOIN (���������� INNER JOIN (������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) ON ����������.[����� ����������] = �������.[��� ����������]) ON ������������.[����� ������������] = �������.������������) ON ������.[����� �������] = �������.������) ON �����������.[����� �����������] = �������.�����������) ON ��������.[����� ��������] = �������.��������";
CText=S+" Where ������������="+IntToStr(Index)+" Order By �������.[����� �������]";

MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
}
//------------
void TMAsp::InpAsp()
{
String S="SELECT �������.[����� �������], �������������.[����� �������������], �������������.[�������� �������������], ����������.[����� ����������], ����������.[������������ ����������], ������������.[����� ������������], ������������.[������������ ������������], ������.[����� �������], ������.[������������ �������], �����������.[����� �����������], �����������.[������������ �����������], �������.����������, �������.Z, ��������.[����� ��������], ��������.[�������� ��������], �������.������������� FROM �������� INNER JOIN (����������� INNER JOIN (������ INNER JOIN (������������ INNER JOIN (���������� INNER JOIN (������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) ON ����������.[����� ����������] = �������.[��� ����������]) ON ������������.[����� ������������] = �������.������������) ON ������.[����� �������] = �������.������) ON �����������.[����� �����������] = �������.�����������) ON ��������.[����� ��������] = �������.��������";
CText=S+" Where ������="+IntToStr(Index)+" Order By �������.[����� �������]";

MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
}
//------------
void TMAsp::InpVozd()
{
String S="SELECT �������.[����� �������], �������������.[����� �������������], �������������.[�������� �������������], ����������.[����� ����������], ����������.[������������ ����������], ������������.[����� ������������], ������������.[������������ ������������], ������.[����� �������], ������.[������������ �������], �����������.[����� �����������], �����������.[������������ �����������], �������.����������, �������.Z, ��������.[����� ��������], ��������.[�������� ��������], �������.������������� FROM �������� INNER JOIN (����������� INNER JOIN (������ INNER JOIN (������������ INNER JOIN (���������� INNER JOIN (������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) ON ����������.[����� ����������] = �������.[��� ����������]) ON ������������.[����� ������������] = �������.������������) ON ������.[����� �������] = �������.������) ON �����������.[����� �����������] = �������.�����������) ON ��������.[����� ��������] = �������.��������";
CText=S+" Where �����������="+IntToStr(Index)+" Order By �������.[����� �������]";

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

Index=Podr->FieldByName("����� �������������")->Value;
String S="SELECT �������.[����� �������], �������������.[����� �������������], �������������.[�������� �������������], ����������.[����� ����������], ����������.[������������ ����������], ������������.[����� ������������], ������������.[������������ ������������], ������.[����� �������], ������.[������������ �������], �����������.[����� �����������], �����������.[������������ �����������], �������.����������, �������.Z, ��������.[����� ��������], ��������.[�������� ��������], �������.������������� FROM �������� INNER JOIN (����������� INNER JOIN (������ INNER JOIN (������������ INNER JOIN (���������� INNER JOIN (������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) ON ����������.[����� ����������] = �������.[��� ����������]) ON ������������.[����� ������������] = �������.������������) ON ������.[����� �������] = �������.������) ON �����������.[����� �����������] = �������.�����������) ON ��������.[����� ��������] = �������.��������";
CText=S+" Where �������������="+IntToStr(Index)+" Order By �������.[����� �������]";

MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::ComboBox2Change(TObject *Sender)
{
String S="SELECT �������.[����� �������], �������������.[����� �������������], �������������.[�������� �������������], ����������.[����� ����������], ����������.[������������ ����������], ������������.[����� ������������], ������������.[������������ ������������], ������.[����� �������], ������.[������������ �������], �����������.[����� �����������], �����������.[������������ �����������], �������.����������, �������.Z, ��������.[����� ��������], ��������.[�������� ��������], �������.������������� FROM �������� INNER JOIN (����������� INNER JOIN (������ INNER JOIN (������������ INNER JOIN (���������� INNER JOIN (������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) ON ����������.[����� ����������] = �������.[��� ����������]) ON ������������.[����� ������������] = �������.������������) ON ������.[����� �������] = �������.������) ON �����������.[����� �����������] = �������.�����������) ON ��������.[����� ��������] = �������.��������";

if (ComboBox2->ItemIndex==0)
{

CText=S+" Where �������������<>0 AND ������������<>0 AND ������<>0 AND �����������<>0 Order By �������.[����� �������]";

}
else
{
CText=S+" Where �������������=0 OR ������������=0 OR ������=0 OR �����������=0 Order By �������.[����� �������]";

}
MoveAspects->Active=false;
MoveAspects->CommandText=CText;
MoveAspects->Active=true;
ChangeCPodr();
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::ComboBox3Change(TObject *Sender)
{
String S="SELECT �������.[����� �������], �������������.[����� �������������], �������������.[�������� �������������], ����������.[����� ����������], ����������.[������������ ����������], ������������.[����� ������������], ������������.[������������ ������������], ������.[����� �������], ������.[������������ �������], �����������.[����� �����������], �����������.[������������ �����������], �������.����������, �������.Z, ��������.[����� ��������], ��������.[�������� ��������], �������.������������� FROM �������� INNER JOIN (����������� INNER JOIN (������ INNER JOIN (������������ INNER JOIN (���������� INNER JOIN (������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) ON ����������.[����� ����������] = �������.[��� ����������]) ON ������������.[����� ������������] = �������.������������) ON ������.[����� �������] = �������.������) ON �����������.[����� �����������] = �������.�����������) ON ��������.[����� ��������] = �������.��������";

if (ComboBox3->ItemIndex==0)
{
CText=S+" Where ����������=True ORDER BY �������.[����� �������]";



}
else
{
CText=S+" Where ����������=False ORDER BY �������.[����� �������]";



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
  Application->HelpJump("IDH_��������_��������");
}
}
//---------------------------------------------------------------------------





void __fastcall TMAsp::ComboBox4DropDown(TObject *Sender)
{

String Text=ComboBox4->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [����� ����������], [������������ ����������] From ���������� Where [������������ ����������] Like '%"+Text+"%' AND �����=true Order by [����� ����������]";
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
ComboBox4->Items->Add(Tab->FieldByName("������������ ����������")->AsString);
}

}
//---------------------------------------------------------------------------
void __fastcall TMAsp::ComboBox4Change(TObject *Sender)
{
String Text=ComboBox4->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [����� ����������], [������������ ����������] From ���������� Where [������������ ����������] Like '%"+Text+"%' AND �����=true Order by [����� ����������]";
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
ComboBox4->Items->Add(Tab->FieldByName("������������ ����������")->AsString);
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
String CT="Select [����� ����������], [������������ ����������] From ���������� Where [������������ ����������] Like '%"+Text+"%' AND �����=true Order by [����� ����������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox4->ItemIndex);
ComboBox4->Text=Tab->FieldByName("������������ ����������")->AsString;
if(!ComboBox4->DroppedDown & ComboBox4->ItemIndex!=-1)
{
MAsp->Index=Tab->FieldByName("����� ����������")->AsInteger;
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
String CT="Select [����� ����������], [������������ ����������] From ���������� Where [������������ ����������] Like '%"+Text+"%' AND �����=true Order by [����� ����������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox4->ItemIndex);
ComboBox4->Text=Tab->FieldByName("������������ ����������")->AsString;
if(ComboBox4->DroppedDown & ComboBox4->ItemIndex!=-1)
{
MAsp->Index=Tab->FieldByName("����� ����������")->AsInteger;
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
String CT="Select [����� ������������], [������������ ������������] From ������������ Where [������������ ������������] Like '%"+Text+"%' AND �����=true Order by [����� ������������]";
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
ComboBox5->Items->Add(Tab->FieldByName("������������ ������������")->AsString);
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
String CT="Select [����� ������������], [������������ ������������] From ������������ Where [������������ ������������] Like '%"+Text+"%' AND �����=true Order by [����� ������������]";
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
ComboBox5->Items->Add(Tab->FieldByName("������������ ������������")->AsString);
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
String CT="Select [����� ������������], [������������ ������������] From ������������ Where [������������ ������������] Like '%"+Text+"%' AND �����=true Order by [����� ������������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox5->ItemIndex);
ComboBox5->Text=Tab->FieldByName("������������ ������������")->AsString;
if(ComboBox5->DroppedDown & ComboBox5->ItemIndex!=-1)
{
MAsp->Index=Tab->FieldByName("����� ������������")->AsInteger;
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
String CT="Select [����� ������������], [������������ ������������] From ������������ Where [������������ ������������] Like '%"+Text+"%' AND �����=true Order by [����� ������������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox5->ItemIndex);
ComboBox5->Text=Tab->FieldByName("������������ ������������")->AsString;
if(!ComboBox5->DroppedDown & ComboBox5->ItemIndex!=-1)
{
MAsp->Index=Tab->FieldByName("����� ������������")->AsInteger;
InpDeyat();
}
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::ComboBox6Change(TObject *Sender)
{
String Text=ComboBox6->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [����� �������], [������������ �������] From ������ Where [������������ �������] Like '%"+Text+"%' AND �����=true Order by [����� �������]";
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
ComboBox6->Items->Add(Tab->FieldByName("������������ �������")->AsString);
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
String CT="Select [����� �������], [������������ �������] From ������ Where [������������ �������] Like '%"+Text+"%' AND �����=true Order by [����� �������]";
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
ComboBox6->Items->Add(Tab->FieldByName("������������ �������")->AsString);
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
String CT="Select [����� �������], [������������ �������] From ������ Where [������������ �������] Like '%"+Text+"%' AND �����=true Order by [����� �������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox6->ItemIndex);
ComboBox6->Text=Tab->FieldByName("������������ �������")->AsString;
if(ComboBox6->DroppedDown & ComboBox6->ItemIndex!=-1)
{
MAsp->Index=Tab->FieldByName("����� �������")->AsInteger;
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
String CT="Select [����� �������], [������������ �������] From ������ Where [������������ �������] Like '%"+Text+"%' AND �����=true Order by [����� �������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox6->ItemIndex);
ComboBox6->Text=Tab->FieldByName("������������ �������")->AsString;
if(!ComboBox6->DroppedDown & ComboBox6->ItemIndex!=-1)
{
MAsp->Index=Tab->FieldByName("����� �������")->AsInteger;
InpAsp();
}
}
//---------------------------------------------------------------------------

void __fastcall TMAsp::ComboBox7Change(TObject *Sender)
{
String Text=ComboBox7->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
String CT="Select [����� �����������], [������������ �����������] From ����������� Where [������������ �����������] Like '%"+Text+"%' AND �����=true Order by [����� �����������]";
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
ComboBox7->Items->Add(Tab->FieldByName("������������ �����������")->AsString);
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
String CT="Select [����� �����������], [������������ �����������] From ����������� Where [������������ �����������] Like '%"+Text+"%' AND �����=true Order by [����� �����������]";
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
ComboBox7->Items->Add(Tab->FieldByName("������������ �����������")->AsString);
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
String CT="Select [����� �����������], [������������ �����������] From ����������� Where [������������ �����������] Like '%"+Text+"%' AND �����=true Order by [����� �����������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox7->ItemIndex);
ComboBox7->Text=Tab->FieldByName("������������ �����������")->AsString;
if(ComboBox7->DroppedDown & ComboBox7->ItemIndex!=-1)
{
MAsp->Index=Tab->FieldByName("����� �����������")->AsInteger;
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
String CT="Select [����� �����������], [������������ �����������] From ����������� Where [������������ �����������] Like '%"+Text+"%' AND �����=true Order by [����� �����������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox7->ItemIndex);
ComboBox7->Text=Tab->FieldByName("������������ �����������")->AsString;
if(!ComboBox7->DroppedDown & ComboBox7->ItemIndex!=-1)
{
MAsp->Index=Tab->FieldByName("����� �����������")->AsInteger;
InpVozd();
}
}
//---------------------------------------------------------------------------

