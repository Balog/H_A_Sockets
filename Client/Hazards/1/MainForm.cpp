//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "MainForm.h"
#include "Zastavka.h"
#include "About.h"
#include "MasterPointer.h"
#include "InputFiltr.h"
#include "Progress.h"
#include "Rep1.h"
#include "About.h"
#include "InpDocs.h"
#include "Main.h"
#include "inifiles.hpp"
#include "F_Vvedenie.h"
#include "Metod.h"
#include "Winuser.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
Quit=false;
Aspects->Connection=Zast->ADOUsrAspect;
Podrazdel->Connection=Zast->ADOUsrAspect;
Situaciya->Connection=Zast->ADOUsrAspect;
Territoriya->Connection=Zast->ADOUsrAspect;
Deyatelnost->Connection=Zast->ADOUsrAspect;
Aspect->Connection=Zast->ADOUsrAspect;
Vozdeystvie->Connection=Zast->ADOUsrAspect;
Posledstvie->Connection=Zast->ADOUsrAspect;
Tiagest->Connection=Zast->ADOUsrAspect;
Prioritet->Connection=Zast->ADOUsrAspect;
TempAspects->Connection=Zast->ADOUsrAspect;
Znachimost->Connection=Zast->ADOUsrAspect;
Report->Connection=Zast->ADOUsrAspect;

Filtr1="";
Filtr2="����������=True";
Registered=false;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormShow(TObject *Sender)
{
Prog->Hide();
//******************************************
//       �������� ����-������� ��������
//     ���� ��� ������ �� ����������������
//         � ���������� �����������
//******************************************
 MP<TADOCommand>Comm(this);
 Comm->Connection=Zast->ADOUsrAspect;
if(!Demo)
{
 Comm->CommandText="DELETE �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+") AND ((�������.Demo)=True));";
 Comm->Execute();
}
else
{
//******************************************
// �������� ��� ������� ���������������� ��� Demo
//******************************************
 Comm->CommandText="UPDATE Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login SET �������.Demo = True WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+")); ";
 Comm->Execute();
}
//****************************************

N3->Enabled=!Demo;

String Text="";
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOAspect;
Tab->CommandText="Select [����� ����������], [������������ ����������] From ���������� Order by [����� ����������]";//"Select [����� ����������], [������������ ����������] From ���������� Where [������������ ����������] Like '%�%' Order by [����� ����������]";
Tab->Active=true;

ComboBox1->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox1->Items->Add(Tab->FieldByName("������������ ����������")->AsString);
}

Tab->Active=false;
String CT="Select [����� ������������], [������������ ������������] From ������������ Where [������������ ������������] Like '%"+Text+"%' AND �����=true Order by [����� ������������]";
Tab->CommandText=CT;
Tab->Active=true;

ComboBox2->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox2->Items->Add(Tab->FieldByName("������������ ������������")->AsString);
}

Tab->Active=false;
CT="Select [����� �������], [������������ �������] From ������ Where [������������ �������] Like '%"+Text+"%' AND �����=true Order by [����� �������]";
Tab->CommandText=CT;
Tab->Active=true;

ComboBox3->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox3->Items->Add(Tab->FieldByName("������������ �������")->AsString);
}

Tab->Active=false;
CT="Select [����� �����������], [������������ �����������] From ����������� Where [������������ �����������] Like '%"+Text+"%' AND �����=true Order by [����� �����������]";
Tab->CommandText=CT;
Tab->Active=true;

ComboBox4->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox4->Items->Add(Tab->FieldByName("������������ �����������")->AsString);
}

String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Hazards.ini");

int Num=Ini->ReadInteger(IntToStr(NumLogin),"CurrentRecord",1);
Filter->CText=Ini->ReadString(IntToStr(NumLogin),"Filter","");
if(Filter->CText=="")
{
 Filter->SetDefFiltr();
}
LFiltr->Caption=Ini->ReadString(IntToStr(NumLogin),"NameFilter","��������");
Filter->NumFiltr=Ini->ReadInteger(IntToStr(NumLogin),"NumFilter", 0);

Form1->LFiltr->Caption=Ini->ReadString(IntToStr(NumLogin),"NameFilter","��������");


InputDocs->TextBr=Ini->ReadString(IntToStr(NumLogin),"TextFilter","");

Initialize(Num);
if(Aspects->RecordCount!=0)
{
CPodrazdel->SetFocus();
}
CountInvalid();
}
//---------------------------------------------------------------------------
void TForm1::Initialize()
{
 Initialize(1);
}
//---------------------------------------------------------------------------
void TForm1::Initialize(int NumRec)
{
SetAspects(Login, NumRec);
SetCombo();

Aspects->Active=true;
if(Podrazdel->RecordCount!=0)
{

}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects","��� ������������ �� ��������� �������������",Login);
ShowMessage("��� ������������ "+Login+" �� ��������� �������������!\r���������� ������ ���������.");
Zast->Close();
}
 InitCombo();
}
//----------------------------------------------------------------------------
void TForm1::SetAspects(String Login, int NumRec)
{
//������ �������, ������������� � ������������� ������������� ��� �������� � ��������� ����.
//���� ������������� ��� ����������� �� ������� ������� �������������.
//���� �������� ������������� ������������� - ������� �������� �� �����

//����� �����, ������� ������������� ������������� � ������� �������������
//����������� �� ��������� ���� ��� ������� ��� ������������� ������� �������
//������ ������� ���������� ������������� ��������

//������ ������� ���������� ������ ������������� ��� ���������� ������ �������������

Aspects->Active=false;
Aspects->CommandText=Filter->CText;
//"SELECT  �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+")) ORDER BY �������.[����� �������];";
Aspects->Active=true;




Podrazdel->Active=false;
Podrazdel->CommandText="SELECT �������������.* FROM Logins INNER JOIN (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+")) ORDER BY �������������.[�������� �������������];";
Podrazdel->Active=true;

Situaciya->Active=false;
Situaciya->CommandText="select * from �������� where [����� ��������]<>0";
Situaciya->Active=true;

if(NumRec==0)
{
BitBtn1->Click();
}
if(NumRec>Aspects->RecordCount)
{
BitBtn4->Click();
}
if(NumRec>0 & NumRec<=Aspects->RecordCount)
{
 Aspects->First();
 Aspects->MoveBy(NumRec-1);
}

}
//---------------------------------------------------------
void TForm1::SetCombo()
{
Podrazdel->Active=false;
Situaciya->Active=false;
Territoriya->Active=false;
Deyatelnost->Active=false;
Aspect->Active=false;
Vozdeystvie->Active=false;
Posledstvie->Active=false;
Tiagest->Active=false;
Prioritet->Active=false;

Podrazdel->Active=true;
Situaciya->Active=true;
Territoriya->Active=true;
Deyatelnost->Active=true;
Aspect->Active=true;
Vozdeystvie->Active=true;
Posledstvie->Active=true;
Tiagest->Active=true;
Prioritet->Active=true;

CPodrazdel->Clear();
for(Podrazdel->First();!Podrazdel->Eof;Podrazdel->Next())
{
 CPodrazdel->Items->Add(Podrazdel->FieldByName("�������� �������������")->AsString);
}

CSit->Clear();
for(Situaciya->First();!Situaciya->Eof;Situaciya->Next())
{
 CSit->Items->Add(Situaciya->FieldByName("�������� ��������")->AsString);
}

CProyav->Clear();
for(Posledstvie->First();!Posledstvie->Eof;Posledstvie->Next())
{
 CProyav->Items->Add(Posledstvie->FieldByName("������������ ����������")->AsString);
}

CPosl->Clear();
for(Tiagest->First();!Tiagest->Eof;Tiagest->Next())
{
 CPosl->Items->Add(Tiagest->FieldByName("������������ �����������")->AsString);
}

CPrior->Clear();
for(Prioritet->First();!Prioritet->Eof;Prioritet->Next())
{
 CPrior->Items->Add(Prioritet->FieldByName("������������ ��������������")->AsString);
}
}
//---------------------------------------------------------
void TForm1::InitCombo()
{
Aspects->Active=true;
Podrazdel->Active=true;
Situaciya->Active=true;
Territoriya->Active=true;
Deyatelnost->Active=true;
Aspect->Active=true;
Vozdeystvie->Active=true;
Posledstvie->Active=true;
Tiagest->Active=true;
Prioritet->Active=true;
 int N1=0,N2=0,N3=0,N4=0,N5=0;

if(Aspects->RecordCount!=0)
{
EnabledForm(true);




 N1=Aspects->FieldByName("�������������")->Value;
 N2=Aspects->FieldByName("��������")->Value;
 N3=Aspects->FieldByName("���������� �����������")->Value;
 N4=Aspects->FieldByName("������� �����������")->Value;
 N5=Aspects->FieldByName("��������������")->Value;



if(Podrazdel->Locate("����� �������������",N1, SO))
{
CPodrazdel->ItemIndex=Podrazdel->RecNo-1;
}
else
{
ShowMessage("������ �������������");
}

if(N2==0)
{
CSit->Style=csDropDown;
CSit->Text="�������� ��������";
}
else
{
if(Situaciya->Locate("����� ��������",N2, SO))
{
CSit->Style=csDropDownList;
CSit->ItemIndex=Situaciya->RecNo-1;
}
else
{
ShowMessage("������ ��������");
}
}

if(Posledstvie->Locate("����� ����������",N3, SO))
{
CProyav->ItemIndex=Posledstvie->RecNo-1;
}
else
{
ShowMessage("������ ����������");
}

if(Tiagest->Locate("����� �����������",N4, SO))
{
CPosl->ItemIndex=Tiagest->RecNo-1;
}
else
{
ShowMessage("������ ����������");
}

if(Prioritet->Locate("����� ��������������",N5, SO))
{
CPrior->ItemIndex=Prioritet->RecNo-1;
}
else
{
ShowMessage("������ ��������������");
}



if (Aspects->RecordCount!=0)
{
DateTimePicker1->Date=Aspects->FieldByName("���� ��������")->Value;
DateTimePicker2->Date=Aspects->FieldByName("������ ��������")->Value;
DateTimePicker3->Date=Aspects->FieldByName("����� ��������")->Value;
}

//----������������� �� ����������-----------------------

int K;
if(Aspects->RecordCount!=0)
{
K=Aspects->FieldByName("��� ����������")->Value;
}
else
{
K=0;
}
Territoriya->Active=false;
Territoriya->CommandText="Select * From ���������� Where [����� ����������]="+IntToStr(K);
Territoriya->Active=true;
ComboBox1->Text=Territoriya->FieldByName("������������ ����������")->Value;


if(Aspects->RecordCount!=0)
{
K=Aspects->FieldByName("������������")->Value;
}
else
{
K=0;
}
Deyatelnost->Active=false;
Deyatelnost->CommandText="Select * From ������������ Where [����� ������������]="+IntToStr(K);
Deyatelnost->Active=true;
ComboBox2->Text=Deyatelnost->FieldByName("������������ ������������")->Value;

if(Aspects->RecordCount!=0)
{
K=Aspects->FieldByName("������")->Value;
}
else
{
K=0;
}
Aspect->Active=false;
Aspect->CommandText="Select * From ������ Where [����� �������]="+IntToStr(K);
Aspect->Active=true;
ComboBox3->Text=Aspect->FieldByName("������������ �������")->Value;

if(Aspects->RecordCount!=0)
{
K=Aspects->FieldByName("�����������")->Value;
}
else
{
K=0;
}
Vozdeystvie->Active=false;
Vozdeystvie->CommandText="Select * From ����������� Where [����� �����������]="+IntToStr(K);
Vozdeystvie->Active=true;
ComboBox4->Text=Vozdeystvie->FieldByName("������������ �����������")->Value;


//----------
if(Aspects->RecordCount!=0)
{
int N=Aspects->FieldByName("����� �������")->Value;

TempAspects->Active=false;
TempAspects->CommandText="Select * From ������� Where [����� �������]="+IntToStr(N);
TempAspects->Active=true;
if(TempAspects->RecordCount!=0)
{
if(TempAspects->FieldByName("G")->Value==1)
{

C1=true;

}
else
{

C1=false;

}

if(TempAspects->FieldByName("R")->Value==1)
{

C2=true;

}
else
{

C2=false;

}

if(TempAspects->FieldByName("T")->Value==1)
{

C3=true;

}
else
{
//
C3=false;

}

if(TempAspects->FieldByName("L")->Value==1)
{

C4=true;

}
else
{

C4=false;

}

if(TempAspects->FieldByName("O")->Value==1)
{

C5=true;

}
else
{

C5=false;

}

if(TempAspects->FieldByName("S")->Value==1)
{

C6=true;

}
else
{

C6=false;


}

if(TempAspects->FieldByName("N")->Value==1)
{

C7=true;

}
else
{

C7=false;

}

if(C1==1)
{
CheckBox1->Checked=true;
}
else
{
CheckBox1->Checked=false;
}

if(C2==1)
{
CheckBox2->Checked=true;
}
else
{
CheckBox2->Checked=false;
}

if(C3==1)
{
CheckBox3->Checked=true;
}
else
{
CheckBox3->Checked=false;
}

if(C4==1)
{
CheckBox4->Checked=true;
}
else
{
CheckBox4->Checked=false;
}

if(C5==1)
{
CheckBox5->Checked=true;
}
else
{
CheckBox5->Checked=false;
}

if(C6==1)
{
CheckBox6->Checked=true;
}
else
{
CheckBox6->Checked=false;
}

if(C7==1)
{
CheckBox7->Checked=true;
}
else
{
CheckBox7->Checked=false;
}
}
}



//----������� ������------------
if(Aspects->RecordCount==0 & Podrazdel->RecordCount!=0)
{
 NewRecord();
}
CountRecord->Text=IntToStr(Aspects->RecordCount);

NumRecord->Text=IntToStr(Aspects->RecNo);


Calc();
}
else
{
EnabledForm(false);
NumRecord->Text="0";
CountRecord->Text="0";
}

 Label1->Visible=IsNew();
}
//---------------------------------------------------------------------------
void TForm1::NewRecord()
{
 NewRecord(0, 0, 0, 0, 0, 0);
}
//---------------------------------------------------------------------------
void TForm1::NewRecord(int Podr, int Sit, int Terr, int Deyat, int Asp, int Vozd)
{
Podrazdel->Active=false;
Podrazdel->Active=true;
Situaciya->Active=false;
Situaciya->Active=true;
Posledstvie->Active=true;
Tiagest->Active=true;
Prioritet->Active=true;


if (Situaciya->RecordCount==0)
{
 Situaciya->Append();
 Situaciya->FieldByName("�������� ��������")->Value="����� ��������";
 Situaciya->Post();
}


Podrazdel->First();
Situaciya->First();
Posledstvie->First();
Tiagest->First();
Prioritet->First();

 Aspects->Append();
 Aspects->FieldByName("��� ����������")->Value=0;
 Aspects->FieldByName("������������")->Value=0;
 Aspects->FieldByName("������")->Value=0;
 Aspects->FieldByName("�����������")->Value=0;

 if(Podr==0)
 {
 int NP=Podrazdel->FieldByName("����� �������������")->Value;
 Aspects->FieldByName("�������������")->Value=NP;
 }
 else
 {
 Aspects->FieldByName("�������������")->Value=Podr;
 }

 if(Sit==0)
 {
 int NS=Situaciya->FieldByName("����� ��������")->Value;
 Aspects->FieldByName("��������")->Value=NS;
 }
 else
 {
  Aspects->FieldByName("��������")->Value=Sit;
 }

 Aspects->FieldByName("�������������")->Value=" ";

 Aspects->FieldByName("������� �����������")->Value=Tiagest->FieldByName("����� �����������")->Value;
 Aspects->FieldByName("���������� �����������")->Value=Posledstvie->FieldByName("����� ����������")->Value;
 Aspects->FieldByName("��������������")->Value=Prioritet->FieldByName("����� ��������������")->Value;
 Aspects->FieldByName("������������ �����������")->Value="";
 Aspects->FieldByName("���� ��������")->Value=Now();
 Aspects->FieldByName("������ ��������")->Value=Now();
 Aspects->FieldByName("����� ��������")->Value=Now();
 Aspects->FieldByName("Demo")->Value=Demo;

 if(Terr!=0)
 {
  Aspects->FieldByName("��� ����������")->Value=Terr;
 }

 if(Deyat!=0)
 {
  Aspects->FieldByName("������������")->Value=Deyat;
 }

 if(Asp!=0)
 {
  Aspects->FieldByName("������")->Value=Asp;
 }

 if(Vozd!=0)
 {
  Aspects->FieldByName("�����������")->Value=Vozd;
 }
 Aspects->Post();

 CheckBox1->Checked=false;
CheckBox2->Checked=false;
CheckBox3->Checked=false;
CheckBox4->Checked=false;
CheckBox5->Checked=false;
CheckBox6->Checked=false;
CheckBox7->Checked=false;

CProyav->ItemIndex=0;
CPosl->ItemIndex=0;
CPodrazdel->ItemIndex=0;
CSit->ItemIndex=0;
 Calc();
 InitCombo();

}
//----------------------------------------------------------------------------
void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
SavePosition();
if(!Demo)
{
switch (Application->MessageBoxA("�������� ��� ��������� �� ������?","����� �� ���������",MB_YESNOCANCEL	+MB_ICONQUESTION+MB_DEFBUTTON1))
{
 case IDYES:
 {
 Quit=true;
Zast->Stop=true;
Zast->MClient->Act.WaitCommand=0;

Zast->Saved=false;
N3->Click();

 Action=caNone;
 break;
 }
 case IDNO:
 {
Zast->MClient->WriteDiaryEvent("NetAspects","���������� ������ ����� �� ������","");
Zast->Stop=true;
Sleep(2000);
Zast->Close();
 Action=caNone;
 break;
 }
 case IDCANCEL:
 {
 Action=caNone;
 break;
 }
}

}
else
{
 Zast->Close();
}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::N8Click(TObject *Sender)
{
this->Close();
}
//---------------------------------------------------------------------------

void TForm1::SynhronTerr()
{

try
{
Territoriya->Active=false;
Territoriya->CommandText="Select * From ���������� Where �����=true";
Territoriya->Active=true;
Territoriya->First();
for(int i=0;i<Territoriya->RecordCount;i++)
{
int Num=Territoriya->FieldByName("����� ����������")->Value;
Branches->Active=false;
Branches->CommandText="Select * From �����_5 Where �����=true and [����� �����]="+IntToStr(Num);
Branches->Active=true;
if (Branches->RecordCount==0)
{
 TempAspects->Active=false;
 TempAspects->CommandText="Select �������.[����� �������], �������.[��� ����������] From ������� Where [��� ����������]="+IntToStr(Num);
 TempAspects->Active=true;
 TempAspects->First();
 for(int j=0;j<TempAspects->RecordCount;j++)
 {
  TempAspects->Edit();
  TempAspects->FieldByName("��� ����������")->Value=0;
  TempAspects->Post();
  TempAspects->Next();
 }
 Territoriya->Delete();
}
else
{
 AnsiString Text=Branches->FieldByName("��������")->Value;
 Territoriya->Edit();
 Territoriya->FieldByName("������������ ����������")->Value=Text;
 Territoriya->Post();
}
 Territoriya->Next();
}
Branches->Active=false;
Branches->CommandText="Select * From �����_5 Where �����=true";
Branches->Active=true;
Branches->First();
for(int i=0;i<Branches->RecordCount;i++)
{
int Num=Branches->FieldByName("����� �����")->Value;
 Territoriya->Active=false;
 Territoriya->CommandText="Select * From ���������� Where [����� ����������]="+IntToStr(Num);
 Territoriya->Active=true;
 if (Territoriya->RecordCount==0)
 {
  int Num=Branches->FieldByName("����� �����")->Value;
  AnsiString Text=Branches->FieldByName("��������")->Value;

 Territoriya->Active=false;
 Territoriya->CommandText="Select * From ����������";
 Territoriya->Active=true;

  Territoriya->Append();
  Territoriya->FieldByName("����� ����������")->Value=Num;
  Territoriya->FieldByName("������������ ����������")->Value=Text;
  Territoriya->FieldByName("�����")->Value=true;
  Territoriya->Post();
 }
Branches->Next();
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ������������� ����������"," ������ "+IntToStr(GetLastError()));
}
}
//---------------------------------------------------------------------------
void TForm1::SynhronDeyat()
{
try
{
Deyatelnost->Active=false;
Deyatelnost->CommandText="Select * From ������������ Where �����=true";
Deyatelnost->Active=true;
Deyatelnost->First();
for(int i=0;i<Deyatelnost->RecordCount;i++)
{
int Num=Deyatelnost->FieldByName("����� ������������")->Value;
Branches->Active=false;
Branches->CommandText="Select * From �����_6 Where �����=true and [����� �����]="+IntToStr(Num);
Branches->Active=true;

if (Branches->RecordCount==0)
{
 TempAspects->Active=false;
 TempAspects->CommandText="Select �������.[����� �������], �������.[������������] From ������� Where [��� ����������]="+IntToStr(Num);
 TempAspects->Active=true;
 TempAspects->First();
 for(int j=0;j<TempAspects->RecordCount;j++)
 {
  TempAspects->Edit();
  TempAspects->FieldByName("������������")->Value=0;
  TempAspects->Post();
  TempAspects->Next();
 }
Deyatelnost->Delete();
}
else
{
 AnsiString Text=Branches->FieldByName("��������")->Value;
 Deyatelnost->Edit();
 Deyatelnost->FieldByName("������������ ������������")->Value=Text;
 Deyatelnost->Post();
}
 Deyatelnost->Next();
}
Branches->Active=false;
Branches->CommandText="Select * From �����_6 Where �����=true";
Branches->Active=true;
Branches->First();
for(int i=0;i<Branches->RecordCount;i++)
{
int Num=Branches->FieldByName("����� �����")->Value;
 Deyatelnost->Active=false;
 Deyatelnost->CommandText="Select * From ������������ Where [����� ������������]="+IntToStr(Num);
 Deyatelnost->Active=true;
 if (Deyatelnost->RecordCount==0)
 {
  int Num=Branches->FieldByName("����� �����")->Value;
  AnsiString Text=Branches->FieldByName("��������")->Value;
  Deyatelnost->Append();
  Deyatelnost->FieldByName("����� ������������")->Value=Num;
  Deyatelnost->FieldByName("������������ ������������")->Value=Text;
  Deyatelnost->FieldByName("�����")->Value=true;
  Deyatelnost->Post();
 }
Branches->Next();
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ������������� �������������","");
}
}
//---------------------------------------------------------------------------
void TForm1::SynhronAspect()
{
try
{
Aspect->Active=false;
Aspect->CommandText="Select * From ������ Where �����=true";
Aspect->Active=true;
Aspect->First();
for(int i=0;i<Aspect->RecordCount;i++)
{
int Num=Aspect->FieldByName("����� �������")->Value;
Branches->Active=false;
Branches->CommandText="Select * From �����_7 Where �����=true and [����� �����]="+IntToStr(Num);
Branches->Active=true;

if (Branches->RecordCount==0)
{
 TempAspects->Active=false;
 TempAspects->CommandText="Select * From ������� Where [������]="+IntToStr(Num);
 TempAspects->Active=true;
 TempAspects->First();
 for(int j=0;j<TempAspects->RecordCount;j++)
 {
  TempAspects->Edit();
  TempAspects->FieldByName("������")->Value=0;
  TempAspects->Post();
  TempAspects->Next();
 }
Aspect->Delete();
}
else
{
 AnsiString Text=Branches->FieldByName("��������")->Value;
Aspect->Edit();
Aspect->FieldByName("������������ �������")->Value=Text;
Aspect->Post();
}
Aspect->Next();
}
Branches->Active=false;
Branches->CommandText="Select * From �����_7 Where �����=true";
Branches->Active=true;
Branches->First();
for(int i=0;i<Branches->RecordCount;i++)
{
int Num=Branches->FieldByName("����� �����")->Value;
Aspect->Active=false;
Aspect->CommandText="Select * From ������ Where [����� �������]="+IntToStr(Num);
Aspect->Active=true;
 if (Aspect->RecordCount==0)
 {
  int Num=Branches->FieldByName("����� �����")->Value;
  AnsiString Text=Branches->FieldByName("��������")->Value;
Aspect->Append();
Aspect->FieldByName("����� �������")->Value=Num;
Aspect->FieldByName("������������ �������")->Value=Text;
Aspect->FieldByName("�����")->Value=true;
Aspect->Post();
 }
Branches->Next();
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ������������� �����������","");
}
}
//---------------------------------------------------------------------------
void TForm1::SynhronVozd()
{
try
{
Vozdeystvie->Active=false;
Vozdeystvie->CommandText="Select * From ����������� Where �����=true";
Vozdeystvie->Active=true;
Vozdeystvie->First();
for(int i=0;i<Vozdeystvie->RecordCount;i++)
{
int Num=Vozdeystvie->FieldByName("����� �����������")->Value;
Branches->Active=false;
Branches->CommandText="Select * From �����_3 Where �����=true and [����� �����]="+IntToStr(Num);
Branches->Active=true;

if (Branches->RecordCount==0)
{
 TempAspects->Active=false;
 TempAspects->CommandText="Select �������.[����� �������], �������.[������������] From ������� Where [��� ����������]="+IntToStr(Num);
 TempAspects->Active=true;
 TempAspects->First();
 for(int j=0;j<TempAspects->RecordCount;j++)
 {
  TempAspects->Edit();
  TempAspects->FieldByName("�����������")->Value=0;
  TempAspects->Post();
  TempAspects->Next();
 }
Vozdeystvie->Delete();
}
else
{
 AnsiString Text=Branches->FieldByName("��������")->Value;
Vozdeystvie->Edit();
Vozdeystvie->FieldByName("������������ �����������")->Value=Text;
Vozdeystvie->Post();
}
Vozdeystvie->Next();
}
Branches->Active=false;
Branches->CommandText="Select * From �����_3 Where �����=true";
Branches->Active=true;
Branches->First();
for(int i=0;i<Branches->RecordCount;i++)
{
int Num=Branches->FieldByName("����� �����")->Value;
Vozdeystvie->Active=false;
Vozdeystvie->CommandText="Select * From ����������� Where [����� �����������]="+IntToStr(Num);
Vozdeystvie->Active=true;
 if (Vozdeystvie->RecordCount==0)
 {
  int Num=Branches->FieldByName("����� �����")->Value;
  AnsiString Text=Branches->FieldByName("��������")->Value;
Vozdeystvie->Append();
Vozdeystvie->FieldByName("����� �����������")->Value=Num;
Vozdeystvie->FieldByName("������������ �����������")->Value=Text;
Vozdeystvie->FieldByName("�����")->Value=true;
Vozdeystvie->Post();
 }
Branches->Next();
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ������������� �����������","");
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPodrazdelClick(TObject *Sender)
{
int N;
Podrazdel->First();
Podrazdel->MoveBy(CPodrazdel->ItemIndex);
N=Podrazdel->FieldByName("����� �������������")->Value;
Aspects->Edit();
Aspects->FieldByName("�������������")->Value=N;
Aspects->Post();

  Label1->Visible=IsNew();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CSitClick(TObject *Sender)
{
int N;
Situaciya->First();
Situaciya->MoveBy(CSit->ItemIndex);
N=Situaciya->FieldByName("����� ��������")->Value;
Aspects->Edit();
Aspects->FieldByName("��������")->Value=N;
Aspects->Post();

  Label1->Visible=IsNew();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CProyavClick(TObject *Sender)
{
int N;
Posledstvie->First();
Posledstvie->MoveBy(CProyav->ItemIndex);
N=Posledstvie->FieldByName("����� ����������")->Value;
Aspects->Edit();
Aspects->FieldByName("���������� �����������")->Value=N;
Aspects->Post();
Aspects->UpdateBatch();
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPoslClick(TObject *Sender)
{
int N;
Tiagest->First();
Tiagest->MoveBy(CPosl->ItemIndex);
N=Tiagest->FieldByName("����� �����������")->Value;
Aspects->Edit();
Aspects->FieldByName("������� �����������")->Value=N;
Aspects->Post();
Aspects->UpdateBatch();
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPriorClick(TObject *Sender)
{
int N;
Prioritet->First();
Prioritet->MoveBy(CPrior->ItemIndex);
N=Prioritet->FieldByName("����� ��������������")->Value;
Aspects->Edit();
Aspects->FieldByName("��������������")->Value=N;
Aspects->Post();

  Label1->Visible=IsNew();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn5Click(TObject *Sender)
{
if(BitBtn5->Caption=="")
{
NewRecord();
}
else
{
int Podr=Podrazdel->FieldByName("����� �������������")->Value;
int Sit=Situaciya->FieldByName("����� ��������")->Value;
int Terr=Territoriya->FieldByName("����� ����������")->Value;;
int Deyat=Deyatelnost->FieldByName("����� ������������")->Value;
int Asp=Aspect->FieldByName("����� �������")->Value;
int Vozd=Vozdeystvie->FieldByName("����� �����������")->Value;
NewRecord(Podr, Sit, Terr, Deyat, Asp, Vozd);
}
C1=false;
C2=false;
C3=false;
C4=false;
C5=false;
C6=false;
C7=false;
InitCombo();
SavePosition();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn1Click(TObject *Sender)
{
Aspects->First();
InitCombo();
SavePosition();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn4Click(TObject *Sender)
{
Aspects->Last();
InitCombo();
SavePosition();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn2Click(TObject *Sender)
{

Aspects->Prior();
InitCombo();
SavePosition();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn3Click(TObject *Sender)
{

Aspects->Next();
InitCombo();
SavePosition();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn6Click(TObject *Sender)
{
{
Aspects->Delete();
InitCombo();
SavePosition();
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button1Click(TObject *Sender)
{

InputDocs->IForm=1;
InputDocs->Mode=1;
InputDocs->ShowModal();

}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button2Click(TObject *Sender)
{

InputDocs->IForm=1;
InputDocs->Mode=2;
InputDocs->ShowModal();

}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button3Click(TObject *Sender)
{

InputDocs->IForm=1;
InputDocs->Mode=3;
InputDocs->ShowModal();

}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button4Click(TObject *Sender)
{

InputDocs->IForm=1;
InputDocs->Mode=4;
InputDocs->ShowModal();

}
//---------------------------------------------------------------------------
void TForm1::InpTer()
{

ComboBox1->Text=InputDocs->TextBr;
Aspects->Edit();
Aspects->FieldByName("��� ����������")->Value=InputDocs->NumBr;
Aspects->Post();
  Label1->Visible=IsNew();
}
//-------------------------------------------
void TForm1::InpDeyat()
{

ComboBox2->Text=InputDocs->TextBr;
Aspects->Edit();
Aspects->FieldByName("������������")->Value=InputDocs->NumBr;
Aspects->Post();
  Label1->Visible=IsNew();
}
//-------------------------------------------
void TForm1::InpAsp()
{

ComboBox3->Text=InputDocs->TextBr;
Aspects->Edit();
Aspects->FieldByName("������")->Value=InputDocs->NumBr;
Aspects->Post();
  Label1->Visible=IsNew();
}
//-------------------------------------------
void TForm1::InpVozd()
{

ComboBox4->Text=InputDocs->TextBr;
Aspects->Edit();
Aspects->FieldByName("�����������")->Value=InputDocs->NumBr;
Aspects->Post();
  Label1->Visible=IsNew();
}
//-------------------------------------------
void TForm1::InpMeropr()
{

DBMemo31->Clear();
DBMemo31->Lines->Add(InputDocs->TextBr);
Aspects->Edit();
Aspects->FieldByName("������������ �����������")->Value=InputDocs->TextBr;
Aspects->Post();
  Label1->Visible=IsNew();
}
//-------------------------------------------
void TForm1::Calc()
{

/*
int G,O,R,S,T,L,N;
int Z,C,F,P;
Aspects->Active=true;
if (Aspects->RecordCount!=0)
{
 if(C1==true)
 {
  G=2;
 }
 else
 {
  G=0;
 }
 if(C5==true)
 {
  O=1;
 }
 else
 {
  O=0;
 }
 if(C2==true)
 {
  R=1;
 }
 else
 {
  R=0;
 }
 if(C6==true)
 {
  S=2;
 }
 else
 {
  S=0;
 }
 if(C3==true)
 {
  T=2;
 }
 else
 {
  T=0;
 }
 if(C4==true)
 {
  L=1;
 }
 else
 {
  L=0;
 }
 if(C7==true)
 {
  N=1;
 }
 else
 {
  N=0;
 }
 Tiagest->Active=false;
 Tiagest->CommandText="Select * From [������� �����������] ORDER By [����� �����������]";
 Tiagest->Active=true;
 Tiagest->First();
 Tiagest->MoveBy(CPosl->ItemIndex);
 P=Tiagest->FieldByName("����")->Value;

 Posledstvie->Active=false;
 Posledstvie->CommandText="Select * From [���������� �����������] Order by [����� ����������]";
 Posledstvie->Active=true;
 Posledstvie->First();
 Posledstvie->MoveBy(CProyav->ItemIndex);
 F=Posledstvie->FieldByName("����")->Value;

 C=G+O+R+S+T+L+N;
 if(C==0)
 {
  C=1;
 }
 Z=C*F*P;

int N=Aspects->FieldByName("����� �������")->Value;

TempAspects->Active=false;
TempAspects->CommandText="Select * From ������� Where [����� �������]="+IntToStr(N);;
TempAspects->Active=true;
TempAspects->Edit();
TempAspects->FieldByName("Z")->Value=Z;
if(C1==true)
{
TempAspects->FieldByName("G")->Value=1;
}
else
{
TempAspects->FieldByName("G")->Value=0;
}

if(C2==true)
{
TempAspects->FieldByName("R")->Value=1;
}
else
{
TempAspects->FieldByName("R")->Value=0;
}

if(C3==true)
{
TempAspects->FieldByName("T")->Value=1;
}
else
{
TempAspects->FieldByName("T")->Value=0;
}

if(C4==true)
{
TempAspects->FieldByName("L")->Value=1;
}
else
{
TempAspects->FieldByName("L")->Value=0;
}

if(C5==true)
{
TempAspects->FieldByName("O")->Value=1;
}
else
{
TempAspects->FieldByName("O")->Value=0;
}

if(C6==true)
{
TempAspects->FieldByName("S")->Value=1;
}
else
{
TempAspects->FieldByName("S")->Value=0;
}

if(C7==true)
{
TempAspects->FieldByName("N")->Value=1;
}
else
{
TempAspects->FieldByName("N")->Value=0;
}
TempAspects->FieldByName("Z")->Value=Z;
TempAspects->Post();
Edit3->Text=IntToStr(Z);



//-----------������-���---------------------


Znachimost->Active=false;
Znachimost->CommandText="Select * From ���������� Order By [���� �������]";
Znachimost->Active=true;
AnsiString NZ;
AnsiString M;
int NumZn=-1;
bool Zn, ZZ=false;
for(int i=0;i<Znachimost->RecordCount;i++)
{
bool BB=Znachimost->FieldByName("��������")->Value;
if(BB==true & ZZ==false)
{
ZZ=true;
Zmax->Text=Znachimost->FieldByName("��� �������")->Value;
}
if (Z>=Znachimost->FieldByName("��� �������")->Value & Z<Znachimost->FieldByName("���� �������")->Value)
{
NumZn=Znachimost->FieldByName("����� ����������")->Value;
NZ=Znachimost->FieldByName("������������ ����������")->Value;
M=Znachimost->FieldByName("����������� ����")->Value;
Zn=Znachimost->FieldByName("��������")->Value;

}
Znachimost->Next();
}
if(ZZ==false)
{
Znachimost->Last();
Zmax->Text="---";
}

if(NumZn==-1)
{
Znachimost->Last();
NumZn=Znachimost->FieldByName("����� ����������")->Value;
NZ=Znachimost->FieldByName("������������ ����������")->Value;
M=Znachimost->FieldByName("����������� ����")->Value;
Zn=Znachimost->FieldByName("��������")->Value;

}

if (Zn==true)
{
 CPrior->Enabled=true;
 Edit6->Font->Color=clRed;
 Edit6->Font->Style=TFontStyles()<< fsBold;
 Edit3->Font->Color=clRed;
 Edit3->Font->Style=TFontStyles()<< fsBold;
 Edit6->Text="��������";

}
else
{
 CPrior->Enabled=false;
 CPrior->ItemIndex=1;
  Edit6->Font->Color=clBlack;
 Edit6->Font->Style= TFontStyles();
 Edit3->Font->Color=clBlack;
 Edit3->Font->Style=TFontStyles();
 Edit6->Text="����������";

}
int N1=Aspects->FieldByName("����� �������")->Value;

TempAspects->Active=false;
TempAspects->CommandText="Select * From ������� Where [����� �������]="+IntToStr(N1);
TempAspects->Active=true;


TempAspects->Edit();
TempAspects->FieldByName("����������")->Value=Zn;
TempAspects->Post();


StatusBar1->Panels->Items[0]->Text=NZ;
StatusBar1->Panels->Items[1]->Text=IntToStr(Z);
StatusBar1->Panels->Items[2]->Text=M;

}
Aspects->UpdateBatch();
*/
double G,O,R,S,T,L,N;
double Z,C,F,P;
double B, UVZ, VVR, Tp, Pr;
Aspects->Active=true;
if (Aspects->RecordCount!=0)
{
 if(C1==true)
 {
  G=0.2;
 }
 else
 {
  G=0;
 }
 if(C2==true)
 {
  R=0.1;
 }
 else
 {
  R=0;
 }
  if(C3==true)
 {
  T=0.2;
 }
 else
 {
  T=0;
 }
 if(C4==true)
 {
  L=0.2;
 }
 else
 {
  L=0;
 }
  if(C5==true)
 {
  O=0.2;
 }
 else
 {
  O=0;
 }
 if(C6==true)
 {
  S=0.1;
 }
 else
 {
  S=0;
 }


 if(C7==true)
 {
  N=0.1;
 }
 else
 {
  N=0;
 }

 if(C1 | C2 | C3 | C4 | C5 | C6)
 {
 CheckBox7->Enabled=false;
 }
 else
 {
 CheckBox7->Enabled=true;
 }

 if(C7)
 {
 CheckBox1->Enabled=false;
 CheckBox2->Enabled=false;
 CheckBox3->Enabled=false;
 CheckBox4->Enabled=false;
 CheckBox5->Enabled=false;
 CheckBox6->Enabled=false;
}
else
{
 CheckBox1->Enabled=true;
 CheckBox2->Enabled=true;
 CheckBox3->Enabled=true;
 CheckBox4->Enabled=true;
 CheckBox5->Enabled=true;
 CheckBox6->Enabled=true;
}

 Posledstvie->Active=false;
 Posledstvie->CommandText="Select * From [���������� �����������] Order by [����]";
 Posledstvie->Active=true;
 Posledstvie->First();
 Posledstvie->MoveBy(CProyav->ItemIndex);
 F=Posledstvie->FieldByName("����")->Value;


 Tiagest->Active=false;
 Tiagest->CommandText="Select * From [������� �����������] ORDER By [����]";
 Tiagest->Active=true;
 Tiagest->First();
 Tiagest->MoveBy(CPosl->ItemIndex);
 UVZ=Tiagest->FieldByName("����")->Value;

Prioritet->Active=false;
Prioritet->CommandText="Select * From ��� Order by [����� ��������������]  DESC";
Prioritet->Active=true;
Prioritet->First();
Prioritet->MoveBy(CPrior->ItemIndex);
VVR=Prioritet->FieldByName("����")->Value;



 C=G+O+R+S+T+L+N;
 if (C==0)
 {
 C=0.1;
 }
 B=C*F;
 Tp=UVZ/VVR;
 Pr=Tp*B;
Edit8->Text=B;
int s1=Tp*100;
double s2=s1;
double s3=s2/100;
if(Tp-s3<0.01)
{
if(Tp-s3>=0.005)
{
s3=s3+0.01;
}
}

Edit7->Text=FloatToStr(((double)((int)(s3*100)))/100);
/*
 C=G+O+R+S+T+L+N;
 if(C==0)
 {
  C=1;
 }
 Z=C*F*P;
*/

int NumRisk=Aspects->FieldByName("����� �������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������� Where [����� �������]="+IntToStr(NumRisk);
TempAspects->Active=true;
TempAspects->Edit();
TempAspects->FieldByName("Z")->Value=Pr;

TempAspects->FieldByName("G")->Value=CheckBox1->Checked;

TempAspects->FieldByName("R")->Value=CheckBox2->Checked;

TempAspects->FieldByName("T")->Value=CheckBox3->Checked;

TempAspects->FieldByName("L")->Value=CheckBox4->Checked;

TempAspects->FieldByName("O")->Value=CheckBox5->Checked;

TempAspects->FieldByName("S")->Value=CheckBox6->Checked;

TempAspects->FieldByName("N")->Value=CheckBox7->Checked;


TempAspects->Post();
//DBEdit7->Text=Z;
int p1=Pr*100;
double p2=p1;
double p3=p2/100;
if(Pr-p3<0.01)
{
if(Pr-p3>=0.005)
{
p3=p3+0.01;
}
}

Edit3->Text=FloatToStr(p3);



//-----------������-���---------------------


Znachimost->Active=false;
Znachimost->CommandText="Select * From ���������� Order By [��� �������]";
Znachimost->Active=true;
AnsiString NZ;
AnsiString M;
int NumZn=-1;
bool Zn, ZZ=false;
for(int i=0;i<Znachimost->RecordCount;i++)
{
bool BB=Znachimost->FieldByName("��������")->Value;
if(BB==true & ZZ==false)
{
ZZ=true;
Zmax->Text=Znachimost->FieldByName("��� �������")->Value;
}
double K1=Znachimost->FieldByName("��� �������")->AsFloat;
double K2=Znachimost->FieldByName("���� �������")->AsFloat;
if (p3>=K1 & p3<K2)
{
NumZn=Znachimost->FieldByName("����� ����������")->Value;
NZ=Znachimost->FieldByName("������������ ����������")->Value;
M=Znachimost->FieldByName("����������� ����")->Value;
Zn=Znachimost->FieldByName("��������")->Value;

}
Znachimost->Next();
}
if(ZZ==false)
{
Znachimost->Last();
Zmax->Text="---";
}

if(NumZn==-1)
{
Znachimost->Last();
NumZn=Znachimost->FieldByName("����� ����������")->Value;
NZ=Znachimost->FieldByName("������������ ����������")->Value;
M=Znachimost->FieldByName("����������� ����")->Value;
Zn=Znachimost->FieldByName("��������")->Value;

}
Edit10->Text=NZ;
if (Zn==true)
{
 //CPrior->Enabled=true;
 Edit6->Font->Color=clRed;
 Edit6->Font->Style=TFontStyles()<< fsBold;
 Edit3->Font->Color=clRed;
 Edit3->Font->Style=TFontStyles()<< fsBold;
 Edit6->Text="��������";

}
else
{
 //CPrior->Enabled=false;
 //CPrior->ItemIndex=1;
  Edit6->Font->Color=clBlack;
 Edit6->Font->Style= TFontStyles();
 Edit3->Font->Color=clBlack;
 Edit3->Font->Style=TFontStyles();
 Edit6->Text="����������";

}



TempAspects->Active=false;
TempAspects->CommandText="Select * From ������� Where [����� �������]="+DBEdit1->Text;
TempAspects->Active=true;


TempAspects->Edit();
TempAspects->FieldByName("����������")->Value=Zn;
TempAspects->FieldByName("������������ ����������")->Value=NZ;
TempAspects->Post();
//DBCheckBox1->Checked=Zn;

CBPrior->ItemIndex=Aspects->FieldByName("��������������")->AsInteger;

StatusBar1->Panels->Items[0]->Text=NZ;
StatusBar1->Panels->Items[1]->Text=FloatToStr(Pr);
StatusBar1->Panels->Items[2]->Text=M;

}
Aspects->UpdateBatch();

 Label1->Visible=IsNew();

CountInvalid();
}
//---------------------------------------------------------------
void __fastcall TForm1::DBCheckBox1Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBCheckBox2Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBCheckBox3Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBCheckBox4Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBCheckBox5Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBCheckBox6Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBCheckBox7Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPoslChange(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------


void __fastcall TForm1::CheckBox1Click(TObject *Sender)
{
C1=CheckBox1->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox2Click(TObject *Sender)
{
C2=CheckBox2->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox3Click(TObject *Sender)
{
C3=CheckBox3->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox4Click(TObject *Sender)
{
C4=CheckBox4->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox5Click(TObject *Sender)
{
C5=CheckBox5->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox6Click(TObject *Sender)
{
C6=CheckBox6->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox7Click(TObject *Sender)
{
C7=CheckBox7->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker1Change(TObject *Sender)
{
Aspects->Edit();
Aspects->FieldByName("���� ��������")->Value=(int)DateTimePicker1->Date;
Aspects->Post();

Aspects->UpdateBatch();
  Label1->Visible=IsNew();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker2Change(TObject *Sender)
{
Aspects->Edit();
Aspects->FieldByName("������ ��������")->Value=(int)DateTimePicker2->Date;
Aspects->Post();

  Label1->Visible=IsNew();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker3Change(TObject *Sender)
{
Aspects->Edit();
Aspects->FieldByName("����� ��������")->Value=(int)DateTimePicker3->Date;
Aspects->Post();

  Label1->Visible=IsNew();
}
//---------------------------------------------------------------------------


void __fastcall TForm1::Button6Click(TObject *Sender)
{

InputDocs->IForm=1;

InputDocs->Mode=5;
InputDocs->ShowModal();

}
//---------------------------------------------------------------------------




void __fastcall TForm1::Button5Click(TObject *Sender)
{
Filter->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn7Click(TObject *Sender)
{

Report1->Flt=Filtr1;
Report1->FltName=LFiltr->Caption;
Report1->PodrComText=Podrazdel->CommandText;

Report1->NumRep=1;
Report1->RepBase=Zast->ADOUsrAspect;
Report1->Role=Role;
Report1->NumLogin=NumLogin;
//ShowMessage(Filter->CText);
int NWhere=Filter->CText.LowerCase().Pos("where")+5;
String SR=Filter->CText.SubString(NWhere, Filter->CText.Length());
int NOrder=SR.LowerCase().Pos("order")-1;
Report1->Flt=SR.SubString(0, NOrder);
Report1->FltName=LFiltr->Caption;

if(this->Role<4)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ContStartReports");
 String ServerSQL="SELECT �������������.[����� �������������], �������������.[�������� �������������] FROM (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� GROUP BY ObslOtdel.Login, �������������.[����� �������������], �������������.[�������� �������������] HAVING (((ObslOtdel.Login)="+IntToStr(NumLogin)+")) ORDER BY �������������.[�������� �������������];";
 String ClientSQL="Select [ServerNum], [�������� �������������] From Temp�������������";
 Zast->MClient->ReadTable("���������",ServerSQL, "���������_�", ClientSQL);
}
else
{
Report1->Podr->Active=false;

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From Temp�������������";
Comm->Execute();


Comm->CommandText="INSERT INTO Temp������������� ( [����� �������������], [�������� �������������], ServerNum ) SELECT �������������.[����� �������������], �������������.[�������� �������������], �������������.ServerNum FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login GROUP BY Logins.Role, �������������.[����� �������������], �������������.[�������� �������������], �������������.ServerNum, �������.������������� HAVING (((Logins.Role)=4));";
Comm->Execute();
Zast->ContStartReports->Execute();
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBMemo1Exit(TObject *Sender)
{
Aspects->UpdateBatch();
  Label1->Visible=IsNew();
}
//----------------------------------

void __fastcall TForm1::N9Click(TObject *Sender)
{
Zast->BlockMK(true);
try
{
Zast->MClient->BlockServer("PrepareReadAspectsUsr");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//------------------------------------------------------------------------
void TForm1::PrepareMergeAspects()
{
Zast->MClient->WriteDiaryEvent("NetAspects","���������� � �������� ��������","");
try
{

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="Select * From TempAspects order by [����� �������]";
Temp->Active=true;

MP<TADODataSet>Memo(this);
Memo->Connection=Zast->ADOUsrAspect;


for(Temp->First();!Temp->Eof;Temp->Next())
{
 int TempNum=Temp->FieldByName("����� �������")->Value;

MP<TMemo>M(this);
M->Visible=false;
M->Parent=this;
M->Width=1009;

TStrings* TT=M->Lines;
TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo1 Where Num="+IntToStr(TempNum);
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("������������� �����������")->Assign(TT);
Temp->Post();


TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo2 Where Num="+IntToStr(TempNum);
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("������������ �����������")->Assign(TT);
Temp->Post();


TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo3 Where Num="+IntToStr(TempNum);
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("���������� � ��������")->Assign(TT);
Temp->Post();

TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo4 Where Num="+IntToStr(TempNum);
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("������������ ���������� � ��������")->Assign(TT);
Temp->Post();
}
Zast->MClient->WriteDiaryEvent("NetAspects","���������� � �������� �������� ���������","");

}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ���������� � �������� ��������"," ������ "+IntToStr(GetLastError()));

}
}
//------------------------------------------------------------------------
void  TForm1::MergeAspects(int NumLogin)
{
Zast->MClient->WriteDiaryEvent("NetAspects","���������� ��������","");
try
{
MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOUsrAspect;
Podr->CommandText="Select * From �������������";
Podr->Active=true;

MP<TADODataSet>TempAsp(this);
TempAsp->Connection=Zast->ADOUsrAspect;
TempAsp->CommandText="Select * From TempAspects";
TempAsp->Active=true;

for(TempAsp->First();!TempAsp->Eof;TempAsp->Next())
{
 int N=TempAsp->FieldByName("�������������")->Value;

 if(Podr->Locate("ServerNum",N,SO))
 {
  int Num=Podr->FieldByName("����� �������������")->Value;

  TempAsp->Edit();
  TempAsp->FieldByName("�������������")->Value=Num;
  TempAsp->Post();
 }
 else
 {
 Zast->MClient->WriteDiaryEvent("NetAspects ������","���� ���������� ��������","");
  ShowMessage("������ ����������� ��������");
 }
}

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
String ST="INSERT INTO ������� ( �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, G, O, R, S, T, L, N, Z, ����������, [���������� �����������], [������� �����������], ��������������, [������������� �����������], [������������ �����������], [���������� � ��������], [������������ ���������� � ��������], �����������, [���� ��������], [������ ��������], [����� ��������], ServerNum ) ";
ST=ST+" SELECT TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������], TempAspects.[������������ �����������], TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������], TempAspects.�����������, TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������], TempAspects.[����� �������] ";
ST=ST+" FROM TempAspects; ";
Comm->CommandText=ST;
Comm->Execute();


}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ���������� ��������"," ������ "+IntToStr(GetLastError()));

}
}
//----------------------------------------------------------------------



void __fastcall TForm1::FormActivate(TObject *Sender)
{
this->Width=1024;
this->Height=742;

}
//---------------------------------------------------------------------------




//---------------------------------------------------------------------------
void __fastcall TForm1::BitBtn9Click(TObject *Sender)
{
Report1->Flt=Filtr2;
Report1->FltName=LFiltr->Caption;
Report1->PodrComText=Podrazdel->CommandText;

Report1->NumRep=2;
Report1->RepBase=Zast->ADOUsrAspect;
Report1->Role=Role;
Report1->NumLogin=NumLogin;


int NWhere=Filter->CText.LowerCase().Pos("where")+5;
String SR=Filter->CText.SubString(NWhere, Filter->CText.Length());
int NOrder=SR.LowerCase().Pos("order")-1;
Report1->Flt=SR.SubString(0, NOrder);
Report1->FltName=LFiltr->Caption;

if(this->Role<4)
{
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ContStartReports");
 String ServerSQL="SELECT �������������.[����� �������������], �������������.[�������� �������������] FROM (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� GROUP BY ObslOtdel.Login, �������������.[����� �������������], �������������.[�������� �������������] HAVING (((ObslOtdel.Login)="+IntToStr(NumLogin)+")) ORDER BY �������������.[�������� �������������];";
 String ClientSQL="Select [ServerNum], [�������� �������������] From Temp�������������";
 Zast->MClient->ReadTable("���������",ServerSQL, "���������_�", ClientSQL);
}
else
{
Report1->Podr->Active=false;

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From Temp�������������";
Comm->Execute();


Comm->CommandText="INSERT INTO Temp������������� ( [����� �������������], [�������� �������������], ServerNum ) SELECT �������������.[����� �������������], �������������.[�������� �������������], �������������.ServerNum FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login GROUP BY Logins.Role, �������������.[����� �������������], �������������.[�������� �������������], �������������.ServerNum, �������.������������� HAVING (((Logins.Role)=4));";
Comm->Execute();
Zast->ContStartReports->Execute();
}

/*
Report1->Flt=Filtr2;
Report1->FltName=LFiltr->Caption;
Report1->PodrComText=Podrazdel->CommandText;

Report1->NumRep=2;
Report1->RepBase=Zast->ADOUsrAspect;
Report1->Role=Role;
Report1->NumLogin=NumLogin;

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("ContStartReports");
 String ServerSQL="SELECT �������������.[����� �������������], �������������.[�������� �������������] FROM (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� GROUP BY ObslOtdel.Login, �������������.[����� �������������], �������������.[�������� �������������] HAVING (((ObslOtdel.Login)="+IntToStr(NumLogin)+")) ORDER BY �������������.[�������� �������������];";
 String ClientSQL="Select [ServerNum], [�������� �������������] From Temp�������������";
 Zast->MClient->ReadTable("�������",ServerSQL, "�������_�", ClientSQL);
*/


}

//---------------------------------------------------------------------------

void __fastcall TForm1::DBMemo2Exit(TObject *Sender)
{
Aspects->UpdateBatch();
  Label1->Visible=IsNew();        
}
//---------------------------------------------------------------------------


void __fastcall TForm1::DBMemo31Exit(TObject *Sender)
{
Aspects->UpdateBatch();
  Label1->Visible=IsNew();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBMemo4Exit(TObject *Sender)
{
Aspects->UpdateBatch();
  Label1->Visible=IsNew();      
}
//---------------------------------------------------------------------------
void __fastcall TForm1::MetodikaN19Click(TObject *Sender)
{

Initialize();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N14Click(TObject *Sender)
{

Initialize();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N15Click(TObject *Sender)
{

Initialize();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N16Click(TObject *Sender)
{

Initialize();
}
//---------------------------------------------------------------------------


void __fastcall TForm1::CProyavKeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPoslKeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPriorKeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPodrazdelKeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CSitKeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker1KeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker2KeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker3KeyPress(TObject *Sender, char &Key)
{
Key=0;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N1Click(TObject *Sender)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_�����_������������");
}
//---------------------------------------------------------------------------
void __fastcall TForm1::Edit1DblClick(TObject *Sender)
{
Button1->Click();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Edit2DblClick(TObject *Sender)
{
Button2->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Edit4DblClick(TObject *Sender)
{
Button3->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Edit5DblClick(TObject *Sender)
{
Button4->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N10Click(TObject *Sender)
{
Zast->BlockMK(true);
try
{
ReadSprav();
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------
void  TForm1::ReadSprav()
{
Prog->SignComplete=true;
Prog->Show();
Prog->PB->Min=0;
Prog->PB->Position=0;
Prog->PB->Max=9;

Documents->ReadWrite.clear();
Str_RW S;
S.NameAction="ReadMetodika";
S.Text="������ ��������...";
S.Num=1;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadPodrazd";
S.Text="������ �������������...";
S.Num=2;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadCrit";
S.Text="������ ���������...";
S.Num=3;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadSit";
S.Text="������ ��������...";
S.Num=4;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadVozd1";
S.Text="������ ������ �����������...";
S.Num=5;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadMeropr1";
S.Text="������ ������ �����������...";
S.Num=6;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadTerr1";
S.Text="������ ������ ����������...";
S.Num=7;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadDeyat1";
S.Text="������ ������ ����� ������������...";
S.Num=8;
Documents->ReadWrite.push_back(S);

S.NameAction="ReadAspect1";
S.Text="������ ������ ������������� ��������...";
S.Num=9;
Documents->ReadWrite.push_back(S);

Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------
void __fastcall TForm1::N3Click(TObject *Sender)
{
Zast->BlockMK(true);
try
{
Zast->MClient->BlockServer("PrepWriteAspUsr_ADM");
//Zast->MClient->BlockServer("PrepWriteAspUsr");
}
catch(...)
{
 Zast->BlockMK(false);
}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::N4Click(TObject *Sender)
{

this->Hide();
Vvedenie->Show();

}
//---------------------------------------------------------------------------


void __fastcall TForm1::N5Click(TObject *Sender)
{

this->Hide();
Metodika->Show();

}
//---------------------------------------------------------------------------

void __fastcall TForm1::WMSysCommand(TMessage & Msg)
{
  switch (Msg.WParam)
  {
case SC_MINIMIZE:
{
Application->Minimize();
break;
}

default:
DefWindowProc(Handle,Msg.Msg,Msg.WParam,Msg.LParam);

 }

}
//-------------------------------------------------------------------
void __fastcall TForm1::N7Click(TObject *Sender)
{
FAbout->ShowModal();
}
//---------------------------------------------------------------------------
void TForm1::EnabledForm(bool B)
{
Label2->Enabled=B;
CPodrazdel->Enabled=B;
Label3->Enabled=B;
CSit->Enabled=B;
Label4->Enabled=B;
ComboBox1->Enabled=B;
if (!B) ComboBox1->Text="";
Button1->Enabled=B;
Label5->Enabled=B;
ComboBox2->Enabled=B;
if (!B) ComboBox2->Text="";
Button2->Enabled=B;
Label6->Enabled=B;
DBEdit2->Enabled=B;
if (!B) DBEdit2->Text="";
Label7->Enabled=B;
ComboBox3->Enabled=B;
if (!B) ComboBox3->Text="";
Button3->Enabled=B;
Label8->Enabled=B;
ComboBox4->Enabled=B;
if (!B) ComboBox4->Text="";
Button4->Enabled=B;
Button5->Enabled=B;
LFiltr->Enabled=B;
Label9->Enabled=B;
Label10->Enabled=B;
DateTimePicker1->Enabled=B;
CheckBox1->Enabled=B;
CheckBox2->Enabled=B;
Label12->Enabled=B;
CheckBox3->Enabled=B;
CheckBox4->Enabled=B;
Label14->Enabled=B;
CheckBox5->Enabled=B;
Label11->Enabled=B;
CheckBox6->Enabled=B;
Label13->Enabled=B;
CheckBox7->Enabled=B;
Label15->Enabled=B;
Label16->Enabled=B;
CProyav->Enabled=B;
Label17->Enabled=B;
CPosl->Enabled=B;
Label18->Enabled=B;
Edit6->Enabled=B;
if (!B) Edit6->Text="";
Label19->Enabled=B;
Edit3->Enabled=B;
if (!B) Edit3->Text="";
Label26->Enabled=B;
Zmax->Enabled=B;
if (!B) Zmax->Text="";
//Label20->Enabled=B;
Label21->Enabled=B;
CPrior->Enabled=B;
Label22->Enabled=B;
Label24->Enabled=B;
Label23->Enabled=B;
Label36->Enabled=B;
Label37->Enabled=B;
DBMemo1->Enabled=B;
Label25->Enabled=B;
DBMemo31->Enabled=B;
Button6->Enabled=B;
DBMemo2->Enabled=B;
DBMemo4->Enabled=B;
Label28->Enabled=B;
DateTimePicker2->Enabled=B;
Label29->Enabled=B;
DateTimePicker3->Enabled=B;
Label32->Enabled=B;
//Label33->Enabled=B;
BitBtn7->Enabled=B;
Label34->Enabled=B;
//Label35->Enabled=B;
BitBtn9->Enabled=B;

BitBtn1->Enabled=B;
BitBtn2->Enabled=B;
BitBtn3->Enabled=B;
BitBtn4->Enabled=B;
BitBtn6->Enabled=B;

StatusBar1->Visible=B;
}
//--------------------------------------------------------------------------
void __fastcall TForm1::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_����_��������������_�_��������_��������");
}
}
//---------------------------------------------------------------------------


void __fastcall TForm1::NumRecordKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==13)
{
Initialize(StrToInt(NumRecord->Text));
SavePosition();
}
}
//---------------------------------------------------------------------------
void TForm1::SavePosition()
{
String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Hazards.ini");
Ini->WriteInteger(IntToStr(NumLogin),"CurrentRecord",Aspects->RecNo);
Ini->WriteString(IntToStr(NumLogin),"Filter",Filter->CText);
Ini->WriteString(IntToStr(NumLogin),"NameFilter",LFiltr->Caption);
Ini->WriteInteger(IntToStr(NumLogin),"NumFilter", Filter->RadioGroup1->ItemIndex);
switch (Filter->RadioGroup1->ItemIndex)
{
 case 0:
 {
Ini->WriteString(IntToStr(NumLogin),"TextFilter","");
 break;
 }
 case 1:
 {
Ini->WriteString(IntToStr(NumLogin),"TextFilter",Filter->ComboBox1->Text);
 break;
 }
 case 2:
 {
Ini->WriteString(IntToStr(NumLogin),"TextFilter",Filter->ComboBox4->Text);
 break;
 }
 case 3:
 {
Ini->WriteString(IntToStr(NumLogin),"TextFilter",Filter->ComboBox5->Text);
 break;
 }
 case 4:
 {
Ini->WriteString(IntToStr(NumLogin),"TextFilter",Filter->ComboBox6->Text);
 break;
 }
 case 5:
 {
Ini->WriteString(IntToStr(NumLogin),"TextFilter",Filter->ComboBox7->Text);
 break;
 }
 case 6:
 {
Ini->WriteString(IntToStr(NumLogin),"TextFilter",Filter->ComboBox2->Text);
 break;
 }
 case 7:
 {
Ini->WriteString(IntToStr(NumLogin),"TextFilter",Filter->ComboBox3->Text);
 break;
 }

}
}
//---------------------------------------------------------------------------
//-----------------------------------------------------------------------------
void __fastcall TForm1::NumRecordKeyPress(TObject *Sender, char &Key)
{
Set <char, '0', '9'> Dig;
Dig << '0' << '1' << '2' << '3' << '4' << '5'
    << '6' << '7' << '8' << '9';
if ( (! Dig.Contains(Key)) && (Key != '\b'))
   { Key = 0; Beep();}
}
//---------------------------------------------------------------------------


void __fastcall TForm1::ComboBox1Change(TObject *Sender)
{
ComboBox1->AutoComplete=false;

String Text=ComboBox1->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� ����������], [������������ ����������] From ���������� Where [������������ ����������] Like '%"+Text+"%' AND �����=true Order by [����� ����������]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox1->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox1->DropDownCount=30;
}



ComboBox1->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox1->Items->Add(Tab->FieldByName("������������ ����������")->AsString);
}
ComboBox1->DroppedDown=true;



ComboBox1->Text=Text;
 keybd_event(35,0,0,0);

 keybd_event(35,0,KEYEVENTF_KEYUP,0);


}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox1DropDown(TObject *Sender)
{
String Text=ComboBox1->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� ����������], [������������ ����������] From ���������� Where [������������ ����������] Like '%"+Text+"%' AND �����=true Order by [����� ����������]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox1->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox1->DropDownCount=30;
}

ComboBox1->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox1->Items->Add(Tab->FieldByName("������������ ����������")->AsString);
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox1KeyPress(TObject *Sender, char &Key)
{
if(Key==13)
{
String Text=ComboBox1->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� ����������], [������������ ����������] From ���������� Where [������������ ����������] Like '%"+Text+"%' AND �����=true Order by [����� ����������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox1->ItemIndex);
ComboBox1->Text=Tab->FieldByName("������������ ����������")->AsString;
if(ComboBox1->DroppedDown & ComboBox1->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("����� ����������")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("������������ ����������")->AsString;

InpTer();
}
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox1Select(TObject *Sender)
{
String Text=ComboBox1->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� ����������], [������������ ����������] From ���������� Where [������������ ����������] Like '%"+Text+"%' AND �����=true Order by [����� ����������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox1->ItemIndex);
ComboBox1->Text=Tab->FieldByName("������������ ����������")->AsString;
if(!ComboBox1->DroppedDown & ComboBox1->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("����� ����������")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("������������ ����������")->AsString;

InpTer();
}
CountInvalid();
}
//---------------------------------------------------------------------------



void __fastcall TForm1::ComboBox2Change(TObject *Sender)
{
ComboBox2->AutoComplete=false;

String Text=ComboBox2->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� ������������], [������������ ������������] From ������������ Where [������������ ������������] Like '%"+Text+"%' AND �����=true Order by [����� ������������]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox2->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox2->DropDownCount=30;
}



ComboBox2->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox2->Items->Add(Tab->FieldByName("������������ ������������")->AsString);
}
ComboBox2->DroppedDown=true;



ComboBox2->Text=Text;
 keybd_event(35,0,0,0);

 keybd_event(35,0,KEYEVENTF_KEYUP,0);


}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox2DropDown(TObject *Sender)
{
String Text=ComboBox2->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� ������������], [������������ ������������] From ������������ Where [������������ ������������] Like '%"+Text+"%' AND �����=true Order by [����� ������������]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox2->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox2->DropDownCount=30;
}

ComboBox2->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox2->Items->Add(Tab->FieldByName("������������ ������������")->AsString);
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox2KeyPress(TObject *Sender, char &Key)
{
if(Key==13)
{
String Text=ComboBox2->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� ������������], [������������ ������������] From ������������ Where [������������ ������������] Like '%"+Text+"%' AND �����=true Order by [����� ������������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox2->ItemIndex);
ComboBox2->Text=Tab->FieldByName("������������ ������������")->AsString;
if(ComboBox2->DroppedDown & ComboBox2->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("����� ������������")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("������������ ������������")->AsString;

InpDeyat();
}
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox2Select(TObject *Sender)
{
String Text=ComboBox2->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� ������������], [������������ ������������] From ������������ Where [������������ ������������] Like '%"+Text+"%' AND �����=true Order by [����� ������������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox1->ItemIndex);
ComboBox2->Text=Tab->FieldByName("������������ ������������")->AsString;
if(!ComboBox2->DroppedDown & ComboBox2->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("����� ������������")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("������������ ������������")->AsString;

InpDeyat();
}
CountInvalid();         
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox3Change(TObject *Sender)
{
ComboBox3->AutoComplete=false;

String Text=ComboBox3->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� �������], [������������ �������] From ������ Where [������������ �������] Like '%"+Text+"%' AND �����=true Order by [����� �������]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox3->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox3->DropDownCount=30;
}



ComboBox3->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox3->Items->Add(Tab->FieldByName("������������ �������")->AsString);
}
ComboBox3->DroppedDown=true;



ComboBox3->Text=Text;
 keybd_event(35,0,0,0);

 keybd_event(35,0,KEYEVENTF_KEYUP,0);


}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox3DropDown(TObject *Sender)
{
String Text=ComboBox3->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� �������], [������������ �������] From ������ Where [������������ �������] Like '%"+Text+"%' AND �����=true Order by [����� �������]";
Tab->CommandText=CT;
Tab->Active=true;

if(Tab->RecordCount<30)
{
ComboBox3->DropDownCount=Tab->RecordCount;
}
else
{
ComboBox3->DropDownCount=30;
}

ComboBox3->Clear();
for(Tab->First();!Tab->Eof;Tab->Next())
{
ComboBox3->Items->Add(Tab->FieldByName("������������ �������")->AsString);
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox3KeyPress(TObject *Sender, char &Key)
{
if(Key==13)
{
String Text=ComboBox3->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� �������], [������������ �������] From ������ Where [������������ �������] Like '%"+Text+"%' AND �����=true Order by [����� �������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox3->ItemIndex);
ComboBox3->Text=Tab->FieldByName("������������ �������")->AsString;
if(ComboBox3->DroppedDown & ComboBox3->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("����� �������")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("������������ �������")->AsString;

InpAsp();
}
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox3Select(TObject *Sender)
{
String Text=ComboBox3->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� �������], [������������ �������] From ������ Where [������������ �������] Like '%"+Text+"%' AND �����=true Order by [����� �������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox1->ItemIndex);
ComboBox3->Text=Tab->FieldByName("������������ �������")->AsString;
if(!ComboBox3->DroppedDown & ComboBox3->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("����� �������")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("������������ �������")->AsString;

InpAsp();
}
 CountInvalid();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox4Change(TObject *Sender)
{
ComboBox4->AutoComplete=false;

String Text=ComboBox4->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� �����������], [������������ �����������] From ����������� Where [������������ �����������] Like '%"+Text+"%' AND �����=true Order by [����� �����������]";
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
ComboBox4->Items->Add(Tab->FieldByName("������������ �����������")->AsString);
}
ComboBox4->DroppedDown=true;



ComboBox4->Text=Text;
 keybd_event(35,0,0,0);

 keybd_event(35,0,KEYEVENTF_KEYUP,0);


}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox4DropDown(TObject *Sender)
{
String Text=ComboBox4->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� �����������], [������������ �����������] From ����������� Where [������������ �����������] Like '%"+Text+"%' AND �����=true Order by [����� �����������]";
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
ComboBox4->Items->Add(Tab->FieldByName("������������ �����������")->AsString);
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox4KeyPress(TObject *Sender, char &Key)
{
if(Key==13)
{
String Text=ComboBox4->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� �����������], [������������ �����������] From ����������� Where [������������ �����������] Like '%"+Text+"%' AND �����=true Order by [����� �����������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox4->ItemIndex);
ComboBox4->Text=Tab->FieldByName("������������ �����������")->AsString;
if(ComboBox4->DroppedDown & ComboBox4->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("����� �����������")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("������������ �����������")->AsString;

InpVozd();
}
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox4Select(TObject *Sender)
{
String Text=ComboBox4->Text;
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
String CT="Select [����� �����������], [������������ �����������] From ����������� Where [������������ �����������] Like '%"+Text+"%' AND �����=true Order by [����� �����������]";
Tab->CommandText=CT;
Tab->Active=true;

Tab->First();
Tab->MoveBy(ComboBox4->ItemIndex);
ComboBox4->Text=Tab->FieldByName("������������ �����������")->AsString;
if(!ComboBox4->DroppedDown & ComboBox4->ItemIndex!=-1)
{
InputDocs->NumBr=Tab->FieldByName("����� �����������")->AsInteger;
InputDocs->TextBr=Tab->FieldByName("������������ �����������")->AsString;

InpVozd();
}
 CountInvalid();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N11Click(TObject *Sender)
{
BitBtn5->Hint="������ �����";
BitBtn5->Caption="";
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N12Click(TObject *Sender)
{
BitBtn5->Hint="����� - ����� �������";
BitBtn5->Caption="C";
}
//---------------------------------------------------------------------------

void __fastcall TForm1::AspQClick(TObject *Sender)
{
Zast->MClient->BlockServer("AspQ1");
}
//---------------------------------------------------------------------------
bool TForm1::IsNew()
{
DataSetRefresh2->Execute();
bool New=false;
MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="SELECT CompareAspects.*, �������������.[����� �������������] FROM ������������� INNER JOIN CompareAspects ON �������������.ServerNum = CompareAspects.�������������; ";
Temp->Active=true;

MP<TADODataSet>Asp(this);
Asp->Connection=Zast->ADOUsrAspect;
Asp->CommandText="Select * from ������� where [����� �������]="+IntToStr(Aspects->FieldByName("����� �������")->AsInteger);
Asp->Active=true;

int SNum=Asp->FieldByName("ServerNum")->AsInteger;
if(Zast->Role!=4)
{
if(Asp->RecordCount!=0)
{
if(Temp->Locate("����� �������", SNum, SO))
{
 //������
 if(Temp->FieldByName("����� �������������")->Value!=Asp->FieldByName("�������������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("��������")->Value!=Asp->FieldByName("��������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("��� ����������")->Value!=Asp->FieldByName("��� ����������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("������������")->Value!=Asp->FieldByName("������������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("�������������")->Value!=Asp->FieldByName("�������������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("������")->Value!=Asp->FieldByName("������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("�����������")->Value!=Asp->FieldByName("�����������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("G")->Value!=Asp->FieldByName("G")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("O")->Value!=Asp->FieldByName("O")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("R")->Value!=Asp->FieldByName("R")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("S")->Value!=Asp->FieldByName("S")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("T")->Value!=Asp->FieldByName("T")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("L")->Value!=Asp->FieldByName("L")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("N")->Value!=Asp->FieldByName("N")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("Z")->Value!=Asp->FieldByName("Z")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("����������")->Value!=Asp->FieldByName("����������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("���������� �����������")->Value!=Asp->FieldByName("���������� �����������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("������� �����������")->Value!=Asp->FieldByName("������� �����������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("��������������")->Value!=Asp->FieldByName("��������������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("������������� �����������")->Value!=Asp->FieldByName("������������� �����������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("������������ �����������")->Value!=Asp->FieldByName("������������ �����������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("���������� � ��������")->Value!=Asp->FieldByName("���������� � ��������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("������������ ���������� � ��������")->Value!=Asp->FieldByName("������������ ���������� � ��������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("���� ��������")->Value!=Asp->FieldByName("���� ��������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("������ ��������")->Value!=Asp->FieldByName("������ ��������")->Value)
 {
  New=true;
 }
 if(Temp->FieldByName("����� ��������")->Value!=Asp->FieldByName("����� ��������")->Value)
 {
  New=true;
 }

}
else
{
 //�� ������
 New=true;
}
}
}
return New;
}
//---------------------------------------------------------------------------
void TForm1::CountInvalid()
{
DataSetRefresh2->Execute();
String CText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(Form1->NumLogin)+")) AND  ([��� ����������]=0 OR ������������=0 OR ������=0 OR �����������=0 OR ��������=0) Order By [����� �������]";

MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
Tab->CommandText=CText;
Tab->Active=true;

int Col=Tab->RecordCount;

if(Col!=0)
{
Label27->Visible=true;
Label27->Caption="���������� �������� - "+IntToStr(Col)+" ��";
}
else
{
Label27->Visible=false;
}
}



