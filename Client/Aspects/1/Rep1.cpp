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
/*
if(!Registration)
{
Registration=true;
 Zast->MClient->Start();
 Zast->MClient->RegForm(this);
 Zast->MClient->Stop();
}
*/
/*
MP<TADODataSet>TempPodr(this);
TempPodr->Connection=Zast->ADOAspect;
TempPodr->CommandText="Select Temp�������������.[ServerNum], Temp�������������.[�������� �������������] From TempDeyat Order by [ServerNum]";

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * From Temp�������������";
Comm->Execute();

Table* SPodr=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

SPodr->SetCommandText("Select �������������.[����� �������������], �������������.[�������� �������������] From ������������ Order by [����� �������������]");
SPodr->Active(true);
TempDeyat->Active=true;
Zast->MClient->LoadTable(SPodr, TempPodr);


if(Zast->MClient->VerifyTable(SPodr, TempPodr)==0)
{

}
Zast->MClient->DeleteTable(this, SPodr);
*/



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
Zast->MClient->ReadTable("�������",ServerSQL, ClientSQL);
}
/*
Podr->Active=true;
Podr->First();

Podr->MoveBy(CPodrazdel->ItemIndex);
int NumPodr=Podr->FieldByName("ServerNum")->Value;
Zast->MClient->Start();
if(NumRep==1)
{
//Form1->CreateReport1(NumPodr, Date1->Date, Date2->Date, Edit1->Text);
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
Zast->MClient->Stop();
this->Close();
*/
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
//Form1->CreateReport1(NumPodr, Date1->Date, Date2->Date, Edit1->Text);
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