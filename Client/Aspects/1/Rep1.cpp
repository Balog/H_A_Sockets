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
Podr->Connection=RepBase;
Podr->CommandText=PodrComText;
Podr->Active=true;


 CPodrazdel->Clear();
for(Podr->First();!Podr->Eof;Podr->Next())
{
 CPodrazdel->Items->Add(Podr->FieldByName("�������� �������������")->Value);

}
 CPodrazdel->ItemIndex=0;
Date1->Date=Now();
Date2->Date=Now();
}
//---------------------------------------------------------------------------
void __fastcall TReport1::Button1Click(TObject *Sender)
{
Podr->First();
Podr->MoveBy(CPodrazdel->ItemIndex);
int NumPodr=Podr->FieldByName("ServerNum")->Value;

//������� TempAspects
MP<TADOCommand>Comm(this);
Comm->Connection=RepBase;
Comm->CommandText="Delete * From TempAspects";
Comm->Execute();


if(Role==4)
{
 //����������������
 //����� ��������� ������
String S="INSERT INTO TempAspects ( [����� �������], �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, G, O, R, S, T, L, N, Z, ����������, [���������� �����������], [������� �����������], ��������������, [������������� �����������], [������������ �����������], [���������� � ��������], [������������ ���������� � ��������], �����������, [���� ��������], [������ ��������], [����� ��������], ServerNum ) ";
S=S+" SELECT TOP 2 �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������, �������.[������������� �����������], �������.[������������ �����������], �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.�����������, �������.[���� ��������], �������.[������ ��������], �������.[����� ��������], �������.ServerNum ";
S=S+" FROM �������;";
Comm->CommandText=S;
Comm->Execute();

CreateRep();
}
else
{
 //�������� ��� ������������
 //����������� ������ � �������
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ContReport");
 String ServerSQL="SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������,  �������.[������������� �����������],  �������.[������������ �����������],  �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.[���� ��������], �������.[������ ��������], �������.[����� ��������] FROM ������� Where �������������="+IntToStr(NumPodr)+" ;";
 String ClientSQL="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������],  TempAspects.[������������ �����������],  TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������],  TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;";
if(Role==2)
{
Zast->MClient->ReadTable("�������",ServerSQL, "�������", ClientSQL);
}
else
{
Zast->MClient->ReadTable("�������",ServerSQL, "�������_�", ClientSQL);
}
}
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
Podr->First();
Podr->MoveBy(CPodrazdel->ItemIndex);
int NumPodr=Podr->FieldByName("ServerNum")->Value;

if(NumRep==1)
{
Reports *R=new Reports();
R->Connect=RepBase;
R->Role=Role;
R->CreateReport1(NumPodr, Date1->Date, Date2->Date, Edit1->Text, Flt, FltName);
delete R;
}
else
{
Reports *R=new Reports();
R->Connect=RepBase;
R->Role=Role;
R->CreateReport2(NumPodr, Date1->Date, Date2->Date, Edit1->Text, Flt, FltName);
delete R;
}
}
//--------------------------------------------------------------------------
