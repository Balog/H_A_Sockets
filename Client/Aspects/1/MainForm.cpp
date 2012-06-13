//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "MainForm.h"
#include "Zastavka.h"
#include "About.h"
#include "MasterPointer.h"
#include "InputFiltr.h"
//#include "LoadKeyFile.h"
//#include "Svod.h"
//#include "Rep1.h"
//#include "File_operations.cpp"

#include "Progress.h"
//#include "F_Vvedenie.h"
//#include "Metod.h"
#include "Rep1.h"
#include "About.h"
#include "InpDocs.h"
#include "Main.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
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
/*
CurrentRecord=0;
Aspects->Active=false;
Aspects->CommandText="select * from ������� ORDER BY [����� �������]";
Aspects->Active=true;
Filtr1="select * from �������";
Filtr2="SELECT * From ������� WHERE ����������=True";
*/

Registered=false;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormShow(TObject *Sender)
{

//ShowMessage("������������");
//Pass->Hide();
//Pass->Close();
//Zast->MClient->Start();

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
  //Comm->CommandText="DELETE Logins.AdmNum, �������.Demo, �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.AdmNum)="+IntToStr(NumLogin)+") AND ((�������.Demo)=True)); ";

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
Initialize();
//N9->Enabled=!Demo;

/*
if(!Registered)
{
Zast->MClient->Start();
Zast->MClient->RegForm(this);
Registered=true;

MP<TADODataSet>Logins(this);
Logins->Connection=Zast->ADOUsrAspect;
Logins->CommandText="Select * From Logins Where ServerNum="+IntToStr(NumLogin);
Logins->Active=true;

Login=Logins->FieldByName("Login")->AsString;
if(Demo)
{
this->Caption="���������������� �����!   ������������: "+Login;
Zast->MClient->WriteDiaryEvent("NetAspects","���������",Login);

}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects","������������",Login);
this->Caption="������������: "+Login;
}

Initialize();
Zast->MClient->Stop();
}


Zast->MClient->WriteDiaryEvent("NetAspects","������ ������ ������������ ��������","");
*/
}
//---------------------------------------------------------------------------
void TForm1::Initialize()
{
//ShowMessage("������ ���������� �������");
SetAspects(Login);
//ShowMessage("���������� ������������");
LoadSpravs();
//ShowMessage("������");

SetCombo();

Aspects->Active=true;
if(Podrazdel->RecordCount!=0)
{
/*
if(Aspects->RecordCount==0)
{
 NewRecord();

}
*/
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
void TForm1::SetAspects(String Login)
{
//������ �������, ������������� � ������������� ������������� ��� �������� � ��������� ����.
//���� ������������� ��� ����������� �� ������� ������� �������������.
//���� �������� ������������� ������������� - ������� �������� �� �����

//����� �����, ������� ������������� ������������� � ������� �������������
//����������� �� ��������� ���� ��� ������� ��� ������������� ������� �������
//������ ������� ���������� ������������� ��������

//������ ������� ���������� ������ ������������� ��� ���������� ������ �������������

Aspects->Active=false;
Aspects->CommandText="SELECT  �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+")) ORDER BY �������.[����� �������];";
Aspects->Active=true;

Podrazdel->Active=false;
Podrazdel->CommandText="SELECT �������������.* FROM Logins INNER JOIN (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+")) ORDER BY �������������.[�������� �������������];";
Podrazdel->Active=true;

Situaciya->Active=false;
Situaciya->CommandText="select * from �������� where [����� ��������]<>0";
Situaciya->Active=true;
}
//---------------------------------------------------------
bool TForm1::LoadSpravs()
{
/*
bool Ret=true;
//�������� ���������� ���� ������������ � ������������� ������.
//����� ������������� ������ ������� ��������� �������������� ������� �������� (������������� �������� ������)
Zast->MClient->WriteDiaryEvent("NetAspects","������ �������� ������������� �������� ������������ (��������������)","");
try
{
FProg->Progress->Position=0;
FProg->Progress->Min=0;
FProg->Progress->Max=7;
FProg->Visible=true;


FProg->Progress->Position++;
FProg->Label1->Caption="������ ����������� ��������...";
FProg->Label1->Repaint();
if(VerifySit())
{
Ret=Ret & LoadSit();
}

FProg->Progress->Position++;
FProg->Label1->Caption="������ ����������� ����������...";
FProg->Label1->Repaint();
if(VerifyTerr())
{
Ret=Ret & LoadTerr();
}

FProg->Progress->Position++;
FProg->Label1->Caption="������ ����������� ����� �������������...";
FProg->Label1->Repaint();
if(VerifyDeyat())
{
Ret=Ret & LoadDeyat();
}

FProg->Progress->Position++;
FProg->Label1->Caption="������ ����������� ��������...";
FProg->Label1->Repaint();
if(VerifyAsp())
{
Ret=Ret & LoadAsp();
}

FProg->Progress->Position++;
FProg->Label1->Caption="������ ����������� ����� �����������...";
FProg->Label1->Repaint();
if(VerifyVozd())
{
Ret=Ret & LoadVozd();
}

FProg->Progress->Position++;
FProg->Label1->Caption="������ ����������� ��������� ������...";
FProg->Label1->Repaint();
if(VerifyCryt())
{
Ret=Ret & LoadCryt();
}

FProg->Progress->Position++;
FProg->Label1->Caption="������ ����������� �����������...";
FProg->Label1->Repaint();
VerifyMeropr();

FProg->Visible=false;

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","���� �������� ������������� �������� ������������ (��������������)","");
ShowMessage("�� ����� ���������� ������������ ��������� ������");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� ������������� �������� ������������ (��������������)"," ������ "+IntToStr(GetLastError()));
}
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ������������� �������� ������������ (��������������)","");
return Ret;
*/
}
//---------------------------------------------------------
void TForm1::SetCombo()
{
Aspects->Active=false;
Podrazdel->Active=false;
Situaciya->Active=false;
Territoriya->Active=false;
Deyatelnost->Active=false;
Aspect->Active=false;
Vozdeystvie->Active=false;
Posledstvie->Active=false;
Tiagest->Active=false;
Prioritet->Active=false;

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
//Zmax->Text=Zast->Zmax;
 int N1=0,N2=0,N3=0,N4=0,N5=0;
// int Num1,Num2,Num3,Num4,Num5;

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
Edit1->Text=Territoriya->FieldByName("������������ ����������")->Value;

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
Edit2->Text=Deyatelnost->FieldByName("������������ ������������")->Value;

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
Edit4->Text=Aspect->FieldByName("������������ �������")->Value;

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
Edit5->Text=Vozdeystvie->FieldByName("������������ �����������")->Value;


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
}
//---------------------------------------------------------------------------
void TForm1::NewRecord()
{
Podrazdel->Active=false;
Podrazdel->Active=true;
Situaciya->Active=false;
Situaciya->Active=true;
Posledstvie->Active=true;
Tiagest->Active=true;
Prioritet->Active=true;



/*
if (Podrazdel->RecordCount==0)
{
 Podrazdel->Append();
 Podrazdel->FieldByName("�������� �������������")->Value="����� �������������";
 Podrazdel->Post();
}
*/
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
 //Aspects->FieldByName("����������")->Value=1;

 int NP=Podrazdel->FieldByName("����� �������������")->Value;
 Aspects->FieldByName("�������������")->Value=NP;
 int NS=Situaciya->FieldByName("����� ��������")->Value;
 Aspects->FieldByName("��������")->Value=NS;
 Aspects->FieldByName("�������������")->Value=" ";
// Aspects->FieldByName("�������������")->Value=Podrazdel->FieldByName("����� �������������")->Value;
// Aspects->FieldByName("��������")->Value=Situaciya->FieldByName("����� ��������")->Value;
 Aspects->FieldByName("������� �����������")->Value=Tiagest->FieldByName("����� �����������")->Value;
 Aspects->FieldByName("���������� �����������")->Value=Posledstvie->FieldByName("����� ����������")->Value;
 Aspects->FieldByName("��������������")->Value=Prioritet->FieldByName("����� ��������������")->Value;
 Aspects->FieldByName("������������ �����������")->Value="";
 Aspects->FieldByName("���� ��������")->Value=Now();
 Aspects->FieldByName("������ ��������")->Value=Now();
 Aspects->FieldByName("����� ��������")->Value=Now();
 Aspects->FieldByName("Demo")->Value=Demo;

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
//---------------------------------------------------------------------------
void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
switch (Application->MessageBoxA("�������� ��� ��������� �� ������?","����� �� ���������",MB_YESNOCANCEL	+MB_ICONQUESTION+MB_DEFBUTTON1))
{
 case IDYES:
 {
Zast->Stop=true;
Zast->MClient->Act.WaitCommand=0;

Zast->Saved=false;
//Zast->PrepareSaveLogins->Execute();


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
//---------------------------------------------------------------------------
void __fastcall TForm1::N8Click(TObject *Sender)
{
//Zast->Close();
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
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn5Click(TObject *Sender)
{


NewRecord();
C1=false;
C2=false;
C3=false;
C4=false;
C5=false;
C6=false;
C7=false;
InitCombo();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn1Click(TObject *Sender)
{
Aspects->First();
InitCombo();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn4Click(TObject *Sender)
{
Aspects->Last();
InitCombo();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn2Click(TObject *Sender)
{

Aspects->Prior();
InitCombo();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn3Click(TObject *Sender)
{

Aspects->Next();
InitCombo();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn6Click(TObject *Sender)
{
//if(Aspects->RecordCount>1)
{
Aspects->Delete();
InitCombo();
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

Edit1->Text=InputDocs->TextBr;
Aspects->Edit();
Aspects->FieldByName("��� ����������")->Value=InputDocs->NumBr;
Aspects->Post();

}
//-------------------------------------------
void TForm1::InpDeyat()
{

Edit2->Text=InputDocs->TextBr;
Aspects->Edit();
Aspects->FieldByName("������������")->Value=InputDocs->NumBr;
Aspects->Post();

}
//-------------------------------------------
void TForm1::InpAsp()
{

Edit4->Text=InputDocs->TextBr;
Aspects->Edit();
Aspects->FieldByName("������")->Value=InputDocs->NumBr;
Aspects->Post();

}
//-------------------------------------------
void TForm1::InpVozd()
{

Edit5->Text=InputDocs->TextBr;
Aspects->Edit();
Aspects->FieldByName("�����������")->Value=InputDocs->NumBr;
Aspects->Post();

}
//-------------------------------------------
void TForm1::InpMeropr()
{

DBMemo31->Clear();
DBMemo31->Lines->Add(InputDocs->TextBr);
Aspects->Edit();
Aspects->FieldByName("������������ �����������")->Value=InputDocs->TextBr;
Aspects->Post();

}
//-------------------------------------------
void TForm1::Calc()
{

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
//DBEdit7->Text=Z;
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
//DBCheckBox1->Checked=Zn;


StatusBar1->Panels->Items[0]->Text=NZ;
StatusBar1->Panels->Items[1]->Text=IntToStr(Z);
StatusBar1->Panels->Items[2]->Text=M;

}
Aspects->UpdateBatch();
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
//DBEdit9->Text=DateTimePicker1->Date;
Aspects->Edit();
Aspects->FieldByName("���� ��������")->Value=DateTimePicker1->Date;
Aspects->Post();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker2Change(TObject *Sender)
{
//DBEdit10->Text=DateTimePicker2->Date;
Aspects->Edit();
Aspects->FieldByName("������ ��������")->Value=DateTimePicker2->Date;
Aspects->Post();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker3Change(TObject *Sender)
{
//DBEdit11->Text=DateTimePicker3->Date;
Aspects->Edit();
Aspects->FieldByName("����� ��������")->Value=DateTimePicker3->Date;
Aspects->Post();
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
Report1->ShowModal();

}
//---------------------------------------------------------------------------

void __fastcall TForm1::FormHide(TObject *Sender)
{
/*
CurrentRecord=Aspects->RecNo;
Aspects->UpdateBatch();
*/
}
//---------------------------------------------------------------------------

//--------------------------------------------------
void __fastcall TForm1::DBMemo1Exit(TObject *Sender)
{
Aspects->UpdateBatch();
}
//----------------------------------

void __fastcall TForm1::N9Click(TObject *Sender)
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="DELETE Logins.AdmNum, �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+"));";
Comm->Execute();

 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("MergeAspectsUser");
 String ServerSQL="SELECT �������.[����� �������],     �������.�������������,     �������.��������,     �������.[��� ����������],     �������.������������,     �������.�������������,     �������.������,     �������.�����������,     �������.G,     �������.O,     �������.R,     �������.S,     �������.T,     �������.L,     �������.N,     �������.Z,     �������.����������,     �������.[���������� �����������],     �������.[������� �����������],     �������.��������������,     �������.[������������� �����������],     �������.[������������ �����������],     �������.[���������� � ��������],     �������.[������������ ���������� � ��������],     �������.�����������,      �������.[���� ��������],     �������.[������ ��������],     �������.[����� ��������] FROM (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) INNER JOIN ������� ON �������������.[����� �������������] = �������.������������� WHERE (((ObslOtdel.Login)="+IntToStr(NumLogin)+"));";
 String ClientSQL="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������], TempAspects.[������������ �����������], TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������], TempAspects.�����������,  TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;";
Zast->MClient->ReadTable("�������", ServerSQL, "�������_�", ClientSQL);

/*
//������ ��������
try
{
Zast->MClient->Start();

Zast->MClient->WriteDiaryEvent("NetAspects","���������� �������� (������)","");

if(LoadAspects(NumLogin))
{
SetAspects(Login);
//ShowMessage("���������� ������������");
//LoadSpravs();
//ShowMessage("������");

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
Zast->MClient->WriteDiaryEvent("NetAspects","���������� �������� (������) ���������","");
 ShowMessage("���������");
}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","���� ���������� �������� (������)","");
 ShowMessage("������ �������� ��������");
}
Zast->MClient->Stop();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ���������� �������� (������)"," ������ "+IntToStr(GetLastError()));
}
*/
}
//---------------------------------------------------------------------------
bool TForm1::LoadAspects(int NumLogin)
{
/*
bool Ret=true;
Zast->MClient->WriteDiaryEvent("NetAspects","������ �������� ��������","");
try
{

MP<TADODataSet>LPodr(this);
LPodr->Connection=Zast->ADOUsrAspect;
LPodr->CommandText="SELECT �������������.[����� �������������], �������������.[�������� �������������], �������������.ServerNum FROM Logins INNER JOIN (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+"));";
LPodr->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="DELETE Logins.AdmNum, �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+"));";
Comm->Execute();

MP<TADODataSet>LTemp(this);
LTemp->Connection=Zast->ADOUsrAspect;
LTemp->CommandText="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������,   TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;";


Table* RTemp=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

RTemp->SetCommandText("SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������,  TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;");
//Remote->SetCommandText("Select ��������.[����� ��������], ��������.[�������� ��������] From �������� Order by [����� ��������]");




for(LPodr->First();!LPodr->Eof;LPodr->Next())
{
Comm->CommandText="Delete * From TempAspects";
Comm->Execute();

int PodrLoad=LPodr->FieldByName("ServerNum")->Value;
Zast->MClient->WriteDiaryEvent("NetAspects","�������� �������� �������������","�������������: "+LPodr->FieldByName("�������� �������������")->AsString);

Zast->MClient->PrepareLoadAspects(PodrLoad ,DBMemo1->Width, DBMemo31->Width, DBMemo2->Width, DBMemo4->Width);
LTemp->Active=false;
LTemp->Active=true;

RTemp->Active(false);
RTemp->Active(true);
FProg->Label1->Caption="�������� �������� "+LPodr->FieldByName("�������� �������������")->AsString;
Zast->MClient->LoadTable(RTemp, LTemp, FProg->Label1, FProg->Progress);

if(Zast->MClient->VerifyTable(LTemp, RTemp)==0)
{
MP<TADODataSet>Memo1(this);
Memo1->Connection=Zast->ADOUsrAspect;
Memo1->CommandText="SELECT Memo1.Num, Memo1.NumStr, Memo1.Text FROM Memo1 ORDER BY Memo1.Num, Memo1.NumStr;";
Memo1->Active=true;

Table* LMemo1=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);
LMemo1->SetCommandText("SELECT Memo1.Num, Memo1.NumStr, Memo1.Text FROM Memo1 ORDER BY Memo1.Num, Memo1.NumStr;");
LMemo1->Active(true);

Zast->MClient->LoadTable(LMemo1, Memo1);
Memo1->Active=false;
Memo1->Active=true;
if(Zast->MClient->VerifyTable(LMemo1, Memo1)==0)
{

MP<TADODataSet>Memo2(this);
Memo2->Connection=Zast->ADOUsrAspect;
Memo2->CommandText="SELECT Memo2.Num, Memo2.NumStr, Memo2.Text FROM Memo2 ORDER BY Memo2.Num, Memo2.NumStr;";
Memo2->Active=true;

Table* LMemo2=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);
LMemo2->SetCommandText("SELECT Memo2.Num, Memo2.NumStr, Memo2.Text FROM Memo2 ORDER BY Memo2.Num, Memo2.NumStr;");
LMemo2->Active(true);

Zast->MClient->LoadTable(LMemo2, Memo2);

if(Zast->MClient->VerifyTable(LMemo2, Memo2)==0)
{

MP<TADODataSet>Memo3(this);
Memo3->Connection=Zast->ADOUsrAspect;
Memo3->CommandText="SELECT Memo3.Num, Memo3.NumStr, Memo3.Text FROM Memo3 ORDER BY Memo3.Num, Memo3.NumStr;";
Memo3->Active=true;

Table* LMemo3=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);
LMemo3->SetCommandText("SELECT Memo3.Num, Memo3.NumStr, Memo3.Text FROM Memo3 ORDER BY Memo3.Num, Memo3.NumStr;");
LMemo3->Active(true);

Zast->MClient->LoadTable(LMemo3, Memo3);

if(Zast->MClient->VerifyTable(LMemo3, Memo3)==0)
{

MP<TADODataSet>Memo4(this);
Memo4->Connection=Zast->ADOUsrAspect;
Memo4->CommandText="SELECT Memo4.Num, Memo4.NumStr, Memo4.Text FROM Memo4 ORDER BY Memo4.Num, Memo4.NumStr;";
Memo4->Active=true;

Table* LMemo4=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);
LMemo4->SetCommandText("SELECT Memo4.Num, Memo4.NumStr, Memo4.Text FROM Memo4 ORDER BY Memo4.Num, Memo4.NumStr;");
LMemo4->Active(true);

Zast->MClient->LoadTable(LMemo4, Memo4);

if(Zast->MClient->VerifyTable(LMemo4, Memo4)==0)
{
PrepareMergeAspects();
MergeAspects(NumLogin);
Zast->MClient->WriteDiaryEvent("NetAspects","�������� �������� ���������","");

}
else
{
Ret=Ret & false;
break;
}
Zast->MClient->DeleteTable(this, LMemo4);

}
else
{
Ret=Ret & false;
break;
}
Zast->MClient->DeleteTable(this, LMemo3);

}
else
{
Ret=Ret & false;
break;
}
Zast->MClient->DeleteTable(this, LMemo2);

}
else
{
Ret=Ret & false;
break;
}
Zast->MClient->DeleteTable(this, LMemo1);
}
else
{
Ret=Ret & false;
break;
}
}
Zast->MClient->DeleteTable(this, RTemp);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","���� ���������� ��������","");

}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ���������� ������������ (������)"," ������ "+IntToStr(GetLastError()));

}
return Ret;
*/


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
/*
this->Top=0;
this->Left=0;
*/

this->Width=1024;
this->Height=742;

}
//---------------------------------------------------------------------------



void __fastcall TForm1::Cdjlysqhttcnhfcgtrnjd1Click(TObject *Sender)
{
/*
FSvod->Visible=true;
this->Visible=false;
*/
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
Report1->ShowModal();

}
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------

void __fastcall TForm1::DBMemo2Exit(TObject *Sender)
{
Aspects->UpdateBatch();        
}
//---------------------------------------------------------------------------


void __fastcall TForm1::DBMemo31Exit(TObject *Sender)
{
Aspects->UpdateBatch();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBMemo4Exit(TObject *Sender)
{
Aspects->UpdateBatch();        
}
//---------------------------------------------------------------------------











void __fastcall TForm1::MetodikaN19Click(TObject *Sender)
{
// CreateNew();
//NameForm(this, "������ ���������� ��������");
Initialize();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N14Click(TObject *Sender)
{
//Connect();
//NameForm(this, "������ ���������� ��������");
Initialize();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N15Click(TObject *Sender)
{
//Rename();
//NameForm(this, "������ ���������� ��������");
Initialize();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N16Click(TObject *Sender)
{
//SaveAs();
//NameForm(this, "������ ���������� ��������");
Initialize();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N17Click(TObject *Sender)
{
//FullPatch();
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

void __fastcall TForm1::N18Click(TObject *Sender)
{
/*
if(Zast->MEcolog!=1)
{
ShowMessage("������ ������ � ������� ��������������");
}
else
{
SZn->ShowModal();
}
*/
}
//---------------------------------------------------------------------------
bool TForm1::VerifySit()
{
/*
bool Ret=false;
try
{
MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOUsrAspect;
Local->CommandText="Select ��������.[����� ��������], ��������.[�������� ��������] From ��������  Where [����� ��������]<>0 Order by [����� ��������]";
Local->Active=true;

Table* Remote=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Remote->SetCommandText("Select ��������.[����� ��������], ��������.[�������� ��������] From �������� Order by [����� ��������]");
Remote->Active(true);

Ret=Zast->MClient->VerifyTable(Local, Remote)!=0;
Zast->MClient->DeleteTable(this, Remote);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� ��������"," ������ "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//--------------------------------------------
bool TForm1::VerifyTerr()
{
/*
bool Ret=false;
try
{
MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOUsrAspect;
Local->CommandText="Select ����������.[����� ����������], ����������.[������������ ����������] From ���������� Order by [����� ����������]";
Local->Active=true;

Table* Remote=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Remote->SetCommandText("Select ����������.[����� ����������], ����������.[������������ ����������] From ���������� Order by [����� ����������]");
Remote->Active(true);

Ret=Zast->MClient->VerifyTable(Local, Remote)!=0;
Zast->MClient->DeleteTable(this, Remote);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� ����������"," ������ "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//--------------------------------------------
bool TForm1::VerifyDeyat()
{
/*
bool Ret=false;
try
{
MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOUsrAspect;
Local->CommandText="Select ������������.[����� ������������], ������������.[������������ ������������] From ������������ Order by [����� ������������]";
Local->Active=true;

Table* Remote=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Remote->SetCommandText("Select ������������.[����� ������������], ������������.[������������ ������������] From ������������ Order by [����� ������������]");
Remote->Active(true);

Ret=Zast->MClient->VerifyTable(Local, Remote)!=0;
Zast->MClient->DeleteTable(this, Remote);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� �������������"," ������ "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//--------------------------------------------
bool TForm1::VerifyAsp()
{
/*
bool Ret=false;
try
{
MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOUsrAspect;
Local->CommandText="Select ������.[����� �������], ������.[������������ �������] From ������ Order by [����� �������]";
Local->Active=true;

Table* Remote=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Remote->SetCommandText("Select ������.[����� �������], ������.[������������ �������] From ������ Order by [����� �������]");
Remote->Active(true);

Ret=Zast->MClient->VerifyTable(Local, Remote)!=0;
Zast->MClient->DeleteTable(this, Remote);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� �����������"," ������ "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//-------------------------------------------
bool TForm1::VerifyVozd()
{
/*
bool Ret=false;
try
{

MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOUsrAspect;
Local->CommandText="Select �����������.[����� �����������], �����������.[������������ �����������] From ����������� Order by [����� �����������]";
Local->Active=true;

Table* Remote=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Remote->SetCommandText("Select �����������.[����� �����������], �����������.[������������ �����������] From ����������� Order by [����� �����������]");
Remote->Active(true);

Ret=Zast->MClient->VerifyTable(Local, Remote)!=0;
Zast->MClient->DeleteTable(this, Remote);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� �����������"," ������ "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//-------------------------------------------
bool TForm1::VerifyCryt()
{
/*
bool Ret=false;
try
{
MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOUsrAspect;
Local->CommandText="Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[��������], ����������.[��������1], ����������.[��� �������], ����������.[���� �������], ����������.[����������� ����] From ���������� Order by [����� ����������]";
Local->Active=true;

Table* Remote=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Remote->SetCommandText("Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[��������], ����������.[��������1], ����������.[��� �������], ����������.[���� �������], ����������.[����������� ����] From ���������� Order by [����� ����������]");
Remote->Active(true);

Ret=Zast->MClient->VerifyTable(Local, Remote)!=0;
Zast->MClient->DeleteTable(this, Remote);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� ���������"," ������ "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//--------------------------------------------
bool TForm1::LoadSit()
{
/*
bool Ret=false;
try
{
Zast->MClient->WriteDiaryEvent("NetAspects","������ �������� ��������","");
MP<TADODataSet>TempSit(this);
TempSit->Connection=Zast->ADOUsrAspect;
TempSit->CommandText="Select TempSit.[����� ��������], TempSit.[�������� ��������] From TempSit Order by [����� ��������]";

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempSit";
Comm->Execute();

Table* Sit=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Sit->SetCommandText("Select ��������.[����� ��������], ��������.[�������� ��������] From �������� Where �����=True  Order by [����� ��������]");
Sit->Active(true);
TempSit->Active=true;
Zast->MClient->LoadTable(Sit, TempSit);


if(Zast->MClient->VerifyTable(Sit, TempSit)==0)
{
MergeSit();
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ��������","");
Ret=true;
}
Zast->MClient->DeleteTable(this, Sit);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects","���� �������� ��������","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� ��������"," ������ "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//-------------------------------------------
void TForm1::MergeSit()
{
/*
try
{
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="UPDATE �������� SET ��������.Del = True WHERE (((��������.�����)=True));";
Comm->Execute();

Zast->MClient->WriteDiaryEvent("NetAspects","������ ����������� ��������","");
MP<TADODataSet>Sit(this);
Sit->Connection=Zast->ADOUsrAspect;
Sit->CommandText="Select * From ��������";
Sit->Active=true;

MP<TADODataSet>TSit(this);
TSit->Connection=Zast->ADOUsrAspect;
TSit->CommandText="Select * From TempSit";
TSit->Active=true;



for(Sit->First();!Sit->Eof;Sit->Next())
{
 int N=Sit->FieldByName("����� ��������")->Value;
 if(TSit->Locate("����� ��������", N, SO))
 {
  Sit->Edit();
  Sit->FieldByName("�������� ��������")->Value=TSit->FieldByName("�������� ��������")->Value;
  Sit->FieldByName("�����")->Value=true;
  Sit->FieldByName("Del")->Value=false;
  Sit->Post();

  TSit->Delete();
 }
}

Comm->CommandText="UPDATE �������� INNER JOIN ������� ON ��������.[����� ��������] = �������.�������� SET �������.�������� = 0 WHERE (((��������.Del)=True));";
Comm->Execute();

Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * from �������� Where Del=true";
Comm->Execute();

Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="INSERT INTO �������� ( [����� ��������], [�������� ��������], ����� ) SELECT TempSit.[����� ��������], TempSit.[�������� ��������], True AS ����� FROM TempSit;";
Comm->Execute();

Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * from TempSit";
Comm->Execute();


}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ����������� ��������"," ������ "+IntToStr(GetLastError()));
}
*/
}
//-----------------------------------------
bool TForm1::LoadTerr()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","������ �������� ����������","");
try
{
MP<TADODataSet>TempTerr(this);
TempTerr->Connection=Zast->ADOUsrAspect;
TempTerr->CommandText="Select TempTerr.[����� ����������], TempTerr.[������������ ����������], TempTerr.[�����] From TempTerr Order by [����� ����������]";

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempTerr";
Comm->Execute();

Table* Terr=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Terr->SetCommandText("Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[�����] From ���������� Order by [����� ����������]");
Terr->Active(true);
TempTerr->Active=true;
Zast->MClient->LoadTable(Terr, TempTerr);


if(Zast->MClient->VerifyTable(Terr, TempTerr)==0)
{
MergeTerr();
RefreshLocalReference("����_5","�����_5");

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ����������","");
Ret=true;
}
Zast->MClient->DeleteTable(this, Terr);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","���� �������� ����������","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� ����������"," ������ "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//--------------------------------------------
void TForm1::MergeTerr()
{
/*
Zast->MClient->WriteDiaryEvent("NetAspects","������ ����������� ����������","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="UPDATE ���������� SET ����������.Del = false;";
Comm->Execute();

MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
Tab->CommandText="select * from ����������";
Tab->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="Select * From TempTerr";
Temp->Active=true;

for(Tab->First();!Tab->Eof;Tab->Next())
{
 int Num=Tab->FieldByName("����� ����������")->Value;

 if(Temp->Locate("����� ����������",Num,SO))
 {
  Tab->Edit();
  Tab->FieldByName("������������ ����������")->Value=Temp->FieldByName("������������ ����������")->Value;
  Tab->FieldByName("�����")->Value=Temp->FieldByName("�����")->Value;
  Tab->Post();

  Temp->Delete();
 }
 else
 {
  Tab->Edit();
  Tab->FieldByName("Del")->Value=true;
  Tab->Post();
 }
}

Comm->CommandText="UPDATE ���������� INNER JOIN ������� ON ����������.[����� ����������] = �������.[��� ����������] SET �������.[��� ����������] = 0 WHERE (((����������.Del)=True));";
Comm->Execute();

Comm->CommandText="DELETE ����������.Del FROM ���������� WHERE (((����������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ���������� ( [����� ����������], [������������ ����������], [�����] ) SELECT TempTerr.[����� ����������], TempTerr.[������������ ����������], TempTerr.[�����] FROM TempTerr;";
Comm->Execute();

Comm->CommandText="DELETE TempTerr.* FROM TempTerr;";
Comm->Execute();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ����������� ����������"," ������ "+IntToStr(GetLastError()));
}
*/
}
//---------------------------------------------

bool TForm1::LoadDeyat()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","������ �������� �������������","");
try
{
MP<TADODataSet>TempDeyat(this);
TempDeyat->Connection=Zast->ADOUsrAspect;
TempDeyat->CommandText="Select TempDeyat.[����� ������������], TempDeyat.[������������ ������������], TempDeyat.[�����] From TempDeyat Order by [����� ������������]";

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempDeyat";
Comm->Execute();

Table* Deyat=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Deyat->SetCommandText("Select ������������.[����� ������������], ������������.[������������ ������������], ������������.[�����] From ������������ Order by [����� ������������]");
Deyat->Active(true);
TempDeyat->Active=true;
Zast->MClient->LoadTable(Deyat, TempDeyat);


if(Zast->MClient->VerifyTable(Deyat, TempDeyat)==0)
{
MergeDeyat();
RefreshLocalReference("����_6","�����_6");
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� �������������","");

Ret=true;
}
Zast->MClient->DeleteTable(this, Deyat);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","���� �������� �������������","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� �������������"," ������ "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//---------------------------------------------
void TForm1::MergeDeyat()
{
/*
Zast->MClient->WriteDiaryEvent("NetAspects","������ ����������� �������������","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="UPDATE ������������ SET ������������.Del = false;";
Comm->Execute();

MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
Tab->CommandText="select * from ������������";
Tab->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="Select * From TempDeyat";
Temp->Active=true;

for(Tab->First();!Tab->Eof;Tab->Next())
{
 int Num=Tab->FieldByName("����� ������������")->Value;

 if(Temp->Locate("����� ������������",Num,SO))
 {
  Tab->Edit();
  Tab->FieldByName("������������ ������������")->Value=Temp->FieldByName("������������ ������������")->Value;
  Tab->FieldByName("�����")->Value=Temp->FieldByName("�����")->Value;
  Tab->Post();

  Temp->Delete();
 }
 else
 {
  Tab->Edit();
  Tab->FieldByName("Del")->Value=true;
  Tab->Post();
 }
}

Comm->CommandText="UPDATE ������������ INNER JOIN ������� ON ������������.[����� ������������] = �������.[������������] SET �������.[������������] = 0 WHERE (((������������.Del)=True));";
Comm->Execute();

Comm->CommandText="DELETE ������������.Del FROM ������������ WHERE (((������������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ������������ ( [����� ������������], [������������ ������������], ����� ) SELECT TempDeyat.[����� ������������], TempDeyat.[������������ ������������], TempDeyat.����� FROM TempDeyat;";
Comm->Execute();

Comm->CommandText="DELETE TempDeyat.* FROM TempDeyat;";
Comm->Execute();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ����������� �������������"," ������ "+IntToStr(GetLastError()));
}
*/
}
//---------------------------------------------

bool TForm1::LoadAsp()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","������ �������� �����������","");
try
{
MP<TADODataSet>TempAsp(this);
TempAsp->Connection=Zast->ADOUsrAspect;
TempAsp->CommandText="Select TempAspect.[����� �������], TempAspect.[������������ �������], TempAspect.[�����] From TempAspect Order by [����� �������]";

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempAspect";
Comm->Execute();

Table* Asp=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Asp->SetCommandText("Select ������.[����� �������], ������.[������������ �������], ������.[�����] From ������ Order by [����� �������]");
Asp->Active(true);
TempAsp->Active=true;
Zast->MClient->LoadTable(Asp, TempAsp);


if(Zast->MClient->VerifyTable(Asp, TempAsp)==0)
{
MergeAsp();
RefreshLocalReference("����_7","�����_7");

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� �����������","");
Ret=true;
}
Zast->MClient->DeleteTable(this, Asp);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","���� �������� �����������","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� �����������"," ������ "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//---------------------------------------------
void TForm1::MergeAsp()
{
/*
Zast->MClient->WriteDiaryEvent("NetAspects","������ ����������� �����������","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="UPDATE ������ SET ������.Del = false;";
Comm->Execute();

MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
Tab->CommandText="select * from ������";
Tab->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="Select * From TempAspect";
Temp->Active=true;

for(Tab->First();!Tab->Eof;Tab->Next())
{
 int Num=Tab->FieldByName("����� �������")->Value;

 if(Temp->Locate("����� �������",Num,SO))
 {
  Tab->Edit();
  Tab->FieldByName("������������ �������")->Value=Temp->FieldByName("������������ �������")->Value;
  Tab->FieldByName("�����")->Value=Temp->FieldByName("�����")->Value;
  Tab->Post();

  Temp->Delete();
 }
 else
 {
  Tab->Edit();
  Tab->FieldByName("Del")->Value=true;
  Tab->Post();
 }
}

Comm->CommandText="UPDATE ������ INNER JOIN ������� ON ������.[����� �������] = �������.[������] SET �������.[������] = 0 WHERE (((������.Del)=True));";
Comm->Execute();

Comm->CommandText="DELETE ������.Del FROM ������ WHERE (((������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ������ ( [����� �������], [������������ �������], ����� ) SELECT TempAspect.[����� �������], TempAspect.[������������ �������], TempAspect.[�����] FROM TempAspect;";
Comm->Execute();

Comm->CommandText="DELETE TempAspect.* FROM TempAspect;";
Comm->Execute();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ����������� �����������"," ������ "+IntToStr(GetLastError()));

}
*/
}
//----------------------------------------------

bool TForm1::LoadVozd()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","������ �������� �����������","");
try
{
MP<TADODataSet>TempVozd(this);
TempVozd->Connection=Zast->ADOUsrAspect;
TempVozd->CommandText="Select TempVozd.[����� �����������], TempVozd.[������������ �����������], TempVozd.[�����] From TempVozd Order by [����� �����������]";

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempVozd";
Comm->Execute();

Table* Vozd=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Vozd->SetCommandText("Select �����������.[����� �����������], �����������.[������������ �����������], �����������.[�����] From ����������� Order by [����� �����������]");
Vozd->Active(true);
TempVozd->Active=true;
Zast->MClient->LoadTable(Vozd, TempVozd);


if(Zast->MClient->VerifyTable(Vozd, TempVozd)==0)
{
MergeVozd();

RefreshLocalReference("����_3","�����_3");
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� �����������","");
Ret=true;
}
Zast->MClient->DeleteTable(this, Vozd);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","���� �������� �����������","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� �����������"," ������ "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//---------------------------------------------
void TForm1::MergeVozd()
{
/*
Zast->MClient->WriteDiaryEvent("NetAspects","������ ����������� �����������","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="UPDATE ����������� SET �����������.Del = false;";
Comm->Execute();

MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
Tab->CommandText="select * from �����������";
Tab->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="Select * From TempVozd";
Temp->Active=true;

for(Tab->First();!Tab->Eof;Tab->Next())
{
 int Num=Tab->FieldByName("����� �����������")->Value;

 if(Temp->Locate("����� �����������",Num,SO))
 {
  Tab->Edit();
  Tab->FieldByName("������������ �����������")->Value=Temp->FieldByName("������������ �����������")->Value;
  Tab->FieldByName("�����")->Value=Temp->FieldByName("�����")->Value;
  Tab->Post();

  Temp->Delete();
 }
 else
 {
  Tab->Edit();
  Tab->FieldByName("Del")->Value=true;
  Tab->Post();
 }
}

Comm->CommandText="UPDATE ����������� INNER JOIN ������� ON �����������.[����� �����������] = �������.[�����������] SET �������.[�����������] = 0 WHERE (((�����������.Del)=True));";
Comm->Execute();

Comm->CommandText="DELETE �����������.Del FROM ����������� WHERE (((�����������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ����������� ( [����� �����������], [������������ �����������], ����� ) SELECT TempVozd.[����� �����������], TempVozd.[������������ �����������], TempVozd.[�����] FROM TempVozd;";
Comm->Execute();

Comm->CommandText="DELETE TempVozd.* FROM TempVozd;";
Comm->Execute();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ����������� �����������"," ������ "+IntToStr(GetLastError()));
}
*/
}
//---------------------------------------------
bool TForm1::LoadCryt()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","������ �������� ���������","");
try
{
MP<TADODataSet>TempZn(this);
TempZn->Connection=Zast->ADOUsrAspect;
TempZn->CommandText="Select TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.[��������], TempZn.[��������1], TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] From TempZn Order by [����� ����������]";

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempZn";
Comm->Execute();

Table* Zn=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

Zn->SetCommandText("Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[��������], ����������.[��������1], ����������.[��� �������], ����������.[���� �������], ����������.[����������� ����] From ���������� Order by [����� ����������]");
Zn->Active(true);

TempZn->Active=true;
Zast->MClient->LoadTable(Zn, TempZn);


if(Zast->MClient->VerifyTable(Zn, TempZn)==0)
{
MergeCryt();

Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� ���������","");
Ret=true;
}
Zast->MClient->DeleteTable(this, Zn);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","���� �������� ���������","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� ���������"," ������ "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//---------------------------------------------
void TForm1::MergeCryt()
{
/*
Zast->MClient->WriteDiaryEvent("NetAspects","������ ����������� ���������","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="DELETE * FROM ����������;";
Comm->Execute();

Comm->CommandText="INSERT INTO ���������� ( [����� ����������], [������������ ����������], ��������, ��������1, [��� �������], [���� �������], [����������� ����] ) SELECT TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.��������, TempZn.��������1, TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] FROM TempZn;";
Comm->Execute();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ����������� ���������"," ������ "+IntToStr(GetLastError()));
}
*/
}
//---------------------------------------------
bool  TForm1::VerifyMeropr()
{
/*
bool Ret=false;
try
{
MP<TADODataSet>LNode(this);
LNode->Connection=Zast->ADOConn;
LNode->CommandText="Select ����_4.[����� ����], ����_4.[��������], ����_4.[��������] From ����_4 Order by ����_4.[����� ����], ����_4.[��������]";

Table* SNode=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Reference")].ServerDB);

SNode->SetCommandText("Select ����_4.[����� ����], ����_4.[��������], ����_4.[��������] From ����_4 Order by ����_4.[����� ����], ����_4.[��������]");
SNode->Active(true);


MP<TADODataSet>LBranch(this);
LBranch->Connection=Zast->ADOConn;
LBranch->CommandText="Select ����_4.[����� ����], ����_4.[��������], ����_4.[��������] From ����_4 Order by ����_4.[����� ����], ����_4.[��������]";

Table* SBranch=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Reference")].ServerDB);

SBranch->SetCommandText("Select ����_4.[����� ����], ����_4.[��������], ����_4.[��������] From ����_4 Order by ����_4.[����� ����], ����_4.[��������]");
SBranch->Active(true);

if(Zast->MClient->VerifyTable(SNode, LNode)!=0 | Zast->MClient->VerifyTable(SBranch, LBranch)!=0)
{
RefreshLocalReference("����_4","�����_4");
Ret=true;
}
Zast->MClient->DeleteTable(this, SNode);
Zast->MClient->DeleteTable(this, SBranch);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� �����������"," ������ "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//----------------------------------------------
void TForm1::RefreshLocalReference(String NameNode, String NameBranch)
{
/*
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
Comm->CommandText="Delete * From TempNode";
Comm->Execute();

MP<TADODataSet>LNode(this);
LNode->Connection=Zast->ADOConn;
LNode->CommandText="Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by TempNode.[����� ����], TempNode.[��������]";
LNode->Active=true;

Table* SNode=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Reference")].ServerDB);

SNode->SetCommandText("Select "+NameNode+".[����� ����], "+NameNode+".[��������], "+NameNode+".[��������] From "+NameNode+" Order by "+NameNode+".[����� ����], "+NameNode+".[��������]");
SNode->Active(true);

Zast->MClient->LoadTable(SNode, LNode);


if(Zast->MClient->VerifyTable(SNode, LNode)==0)
{

Comm->CommandText="Delete * From TempBranch";
Comm->Execute();

MP<TADODataSet>LBranch(this);
LBranch->Connection=Zast->ADOConn;
LBranch->CommandText="Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by TempBranch.[����� �����], TempBranch.[����� ��������]";
LBranch->Active=true;

Table* SBranch=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Reference")].ServerDB);

SBranch->SetCommandText("Select "+NameBranch+".[����� �����], "+NameBranch+".[����� ��������], "+NameBranch+".[��������], "+NameBranch+".[�����] From "+NameBranch+" Order by "+NameBranch+".[����� �����], "+NameBranch+".[����� ��������]");
SBranch->Active(true);

Zast->MClient->LoadTable(SBranch, LBranch);


if(Zast->MClient->VerifyTable(SBranch, LBranch)==0)
{

Comm->CommandText="Delete * from "+NameNode;
Comm->Execute();

Comm->CommandText="INSERT INTO "+NameNode+" ( [����� ����], ��������, �������� ) SELECT TempNode.[����� ����], TempNode.��������, TempNode.�������� FROM TempNode;";
Comm->Execute();

Comm->CommandText="INSERT INTO "+NameBranch+" ( [����� �����], [����� ��������], ��������, ����� ) SELECT TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.��������, TempBranch.����� FROM TempBranch;";
Comm->Execute();

}
Zast->MClient->DeleteTable(this, SBranch);

}
Zast->MClient->DeleteTable(this, SNode);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ���������� ���������� ������"," ������ "+IntToStr(GetLastError()));
}
*/
}
//---------------------------------------------

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
Prog->SignComplete=true;
Prog->Show();
Prog->PB->Min=0;
Prog->PB->Position=0;
Prog->PB->Max=8;

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

Zast->ReadWriteDoc->Execute();
/*
Zast->MClient->Start();
Zast->MClient->WriteDiaryEvent("NetAspects","������ ���������� ������������ (������)","");
try
{

if(LoadSpravs())
{
SetCombo();
InitCombo();
Zast->MClient->WriteDiaryEvent("NetAspects","����� ���������� ������������ (������)","");
 ShowMessage("���������");
}
else
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","���� ���������� ������������ (������)","");

 ShowMessage("������ �������� ������������");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ���������� ������������ (������)"," ������ "+IntToStr(GetLastError()));

}
Zast->MClient->Stop();
*/
}
//---------------------------------------------------------------------------




void __fastcall TForm1::N3Click(TObject *Sender)
{
Zast->MClient->Act.ParamComm.clear();
Zast->MClient->Act.ParamComm.push_back("WriteAspectsUsr");
String ClientSQL="SELECT �������.[����� �������],     �������.�������������,     �������.��������,     �������.[��� ����������],     �������.������������,     �������.�������������,     �������.������,     �������.�����������,     �������.G,     �������.O,     �������.R,     �������.S,     �������.T,     �������.L,     �������.N,     �������.Z,     �������.����������,     �������.[���������� �����������],     �������.[������� �����������],     �������.��������������,     �������.[������������� �����������],     �������.[������������ �����������],     �������.[���������� � ��������],     �������.[������������ ���������� � ��������],     �������.�����������,      �������.[���� ��������],     �������.[������ ��������],     �������.[����� ��������] FROM ������� ORDER BY �������.[����� �������];";
String ServerSQL="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[������������� �����������], TempAspects.[������������ �����������], TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������], TempAspects.�����������,  TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;";
Zast->MClient->WriteTable("�������_�", ClientSQL, "�������", ServerSQL);


/*
Zast->MClient->Start();
Zast->MClient->WriteDiaryEvent("NetAspects","������ ������ �������� (������)","");
try
{
DataSetRefresh1->Execute();
DataSetPost1->Execute();
if(PrepareSaveAspects())
{

if(SaveAspects())
{
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ �������� (������)","");
 ShowMessage("���������");
}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","���� ������ �������� (������)","");

 ShowMessage("������ ������ ��������");
}
}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","���� ���������� � ������ �������� (������)","");

 ShowMessage("������ ���������� ������ ��������");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ������ �������� (������)"," ������ "+IntToStr(GetLastError()));

}
Zast->MClient->Stop();
*/
}
//---------------------------------------------------------------------------
bool  TForm1::PrepareSaveAspects()
{
/*
Zast->MClient->WriteDiaryEvent("NetAspects","������ ���������� � ������ ��������","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempAspects";
Comm->Execute();

String CT="INSERT INTO TempAspects ( [����� �������], �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, G, O, R, S, T, L, N, Z, ����������, [���������� �����������], [������� �����������], ��������������, [������������� �����������], [������������ �����������], [���������� � ��������], [������������ ���������� � ��������], �����������, [���� ��������], [������ ��������], [����� ��������], ServerNum ) ";
CT=CT+" SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������, �������.[������������� �����������], �������.[������������ �����������], �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.�����������, �������.[���� ��������], �������.[������ ��������], �������.[����� ��������], �������.[ServerNum] ";
CT=CT+" FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login ";
CT=CT+" WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+"));";

Comm->CommandText=CT;
Comm->Execute();

MP<TADODataSet>Temp(Form1);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="Select * From TempAspects";
Temp->Active=true;

for(Temp->First();!Temp->Eof;Temp->Next())
{
 int Num=Temp->FieldByName("����� �������")->Value;

MP<TDataSource>DS(this);
DS->DataSet=Temp;
DS->Enabled=true;

MP<TDBMemo>TDBM(this);
TDBM->Visible=false;
TDBM->Parent=Zast;
TDBM->Width=DBMemo1->Width;
TDBM->DataSource=DS;
TDBM->DataField="������������� �����������";

TStrings* TS=TDBM->Lines;
int N=TS->Count;

MP<TADODataSet>Memo1(this);
Memo1->Connection=Zast->ADOUsrAspect;
Memo1->CommandText="Select * From Memo1";
Memo1->Active=true;

for(int i=0;i<N;i++)
{
String S=TS->Strings[i];
Memo1->Insert();
Memo1->FieldByName("Num")->Value=Num;
Memo1->FieldByName("NumStr")->Value=i;
Memo1->FieldByName("Text")->Value=S;
Memo1->Post();
}

MP<TDBMemo>TDBM2(this);
TDBM2->Parent=Zast;
TDBM2->Visible=false;
TDBM2->Width=DBMemo31->Width;
TDBM2->DataSource=DS;
TDBM2->DataField="������������ �����������";

TStrings* TS2=TDBM2->Lines;
int N2=TS2->Count;

MP<TADODataSet>Memo2(this);
Memo2->Connection=Zast->ADOUsrAspect;
Memo2->CommandText="Select * From Memo2";
Memo2->Active=true;

for(int i=0;i<N2;i++)
{
String S=TS2->Strings[i];
Memo2->Insert();
Memo2->FieldByName("Num")->Value=Num;
Memo2->FieldByName("NumStr")->Value=i;
Memo2->FieldByName("Text")->Value=S;
Memo2->Post();
}

MP<TDBMemo>TDBM3(this);
TDBM3->Parent=Zast;
TDBM3->Visible=false;
TDBM3->Width=DBMemo2->Width;
TDBM3->DataSource=DS;
TDBM3->DataField="���������� � ��������";

TStrings* TS3=TDBM3->Lines;
int N3=TS3->Count;

MP<TADODataSet>Memo3(this);
Memo3->Connection=Zast->ADOUsrAspect;
Memo3->CommandText="Select * From Memo3";
Memo3->Active=true;

for(int i=0;i<N3;i++)
{
String S=TS3->Strings[i];
Memo3->Insert();
Memo3->FieldByName("Num")->Value=Num;
Memo3->FieldByName("NumStr")->Value=i;
Memo3->FieldByName("Text")->Value=S;
Memo3->Post();
}

MP<TDBMemo>TDBM4(this);
TDBM4->Parent=Zast;
TDBM4->Visible=false;
TDBM4->Width=DBMemo4->Width;
TDBM4->DataSource=DS;
TDBM4->DataField="������������ ���������� � ��������";

TStrings* TS4=TDBM4->Lines;
int N4=TS4->Count;

MP<TADODataSet>Memo4(this);
Memo4->Connection=Zast->ADOUsrAspect;
Memo4->CommandText="Select * From Memo4";
Memo4->Active=true;

for(int i=0;i<N4;i++)
{
String S=TS4->Strings[i];
Memo4->Insert();
Memo4->FieldByName("Num")->Value=Num;
Memo4->FieldByName("NumStr")->Value=i;
Memo4->FieldByName("Text")->Value=S;
Memo4->Post();
}
}
Zast->MClient->WriteDiaryEvent("NetAspects","����� ���������� � ������ ��������","");

return true;
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ���������� � ������ ��������"," ������ "+IntToStr(GetLastError()));

return false;
}
*/
}
//-----------------------------------------------------------
bool TForm1::SaveAspects()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","������ ������ ��������","");
try
{
Zast->MClient->SetCommandText("�������","Delete * From TempAspects");
Zast->MClient->CommandExec("�������");

MP<TADODataSet>LTemp(this);
LTemp->Connection=Zast->ADOUsrAspect;
LTemp->CommandText="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������], TempAspects.ServerNum FROM TempAspects ORDER BY TempAspects.[����� �������];";
LTemp->Active=true;

Table* STemp=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

STemp->SetCommandText("SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������], TempAspects.ServerNum FROM TempAspects ORDER BY TempAspects.[����� �������];");
STemp->Active(true);
FProg->Label1->Caption="���������� ��������";
Zast->MClient->LoadTable(LTemp, STemp, FProg->Label1, FProg->Progress);

if(Zast->MClient->VerifyTable(LTemp, STemp)==0)
{

MP<TADODataSet>LMemo1(this);
LMemo1->Connection=Zast->ADOUsrAspect;
LMemo1->CommandText="SELECT Memo1.Num, Memo1.NumStr, Memo1.Text FROM Memo1 ORDER BY Memo1.Num;";
LMemo1->Active=true;

Table* SMemo1=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

SMemo1->SetCommandText("SELECT Memo1.Num, Memo1.NumStr, Memo1.Text FROM Memo1 ORDER BY Memo1.Num;");
SMemo1->Active(true);

Zast->MClient->LoadTable(LMemo1, SMemo1);

if(Zast->MClient->VerifyTable(LMemo1, SMemo1)==0)
{

MP<TADODataSet>LMemo2(this);
LMemo2->Connection=Zast->ADOUsrAspect;
LMemo2->CommandText="SELECT Memo2.Num, Memo2.NumStr, Memo2.Text FROM Memo2 ORDER BY Memo2.Num;";
LMemo2->Active=true;

Table* SMemo2=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

SMemo2->SetCommandText("SELECT Memo2.Num, Memo2.NumStr, Memo2.Text FROM Memo2 ORDER BY Memo2.Num;");
SMemo2->Active(true);

Zast->MClient->LoadTable(LMemo2, SMemo2);

if(Zast->MClient->VerifyTable(LMemo2, SMemo2)==0)
{

MP<TADODataSet>LMemo3(this);
LMemo3->Connection=Zast->ADOUsrAspect;
LMemo3->CommandText="SELECT Memo3.Num, Memo3.NumStr, Memo3.Text FROM Memo3 ORDER BY Memo3.Num;";
LMemo3->Active=true;

Table* SMemo3=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

SMemo3->SetCommandText("SELECT Memo3.Num, Memo3.NumStr, Memo3.Text FROM Memo3 ORDER BY Memo3.Num;");
SMemo3->Active(true);

Zast->MClient->LoadTable(LMemo3, SMemo3);

if(Zast->MClient->VerifyTable(LMemo3, SMemo3)==0)
{

MP<TADODataSet>LMemo4(this);
LMemo4->Connection=Zast->ADOUsrAspect;
LMemo4->CommandText="SELECT Memo4.Num, Memo4.NumStr, Memo4.Text FROM Memo4 ORDER BY Memo4.Num;";
LMemo4->Active=true;

Table* SMemo4=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

SMemo4->SetCommandText("SELECT Memo4.Num, Memo4.NumStr, Memo4.Text FROM Memo4 ORDER BY Memo4.Num;");
SMemo4->Active(true);

Zast->MClient->LoadTable(LMemo4, SMemo4);

if(Zast->MClient->VerifyTable(LMemo4, SMemo4)==0)
{
Zast->MClient->SetCommandText("�������","Delete * From Temp�������������");
Zast->MClient->CommandExec("�������");

MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOUsrAspect;
Podr->CommandText="SELECT �������������.[����� �������������], �������������.[�������� �������������], �������������.ServerNum FROM �������������;";
Podr->Active=true;

Table* TempPodr=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

TempPodr->SetCommandText("SELECT Temp�������������.[����� �������������], Temp�������������.[�������� �������������], Temp�������������.ServerNum FROM Temp�������������;");
TempPodr->Active(true);

Zast->MClient->LoadTable(Podr, TempPodr);

if(Zast->MClient->VerifyTable(Podr, TempPodr)==0)
{
Ret=true;
try
{
Zast->MClient->PrepareMergeAspects(NumLogin, DBMemo1->Width, DBMemo31->Width, DBMemo2->Width, DBMemo4->Width);

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * from TempAspects";
Comm->Execute();

MP<TADODataSet>Asp(this);
Asp->Connection=Zast->ADOUsrAspect;
Asp->CommandText="SELECT �������.* FROM Logins INNER JOIN ((������������� INNER JOIN ������� ON �������������.[����� �������������] = �������.�������������) INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+"));";
Asp->Active=true;

MP<TADODataSet>LT(this);
LT->Connection=Zast->ADOUsrAspect;
LT->CommandText="SELECT TempAspects.[����� �������], TempAspects.ServerNum FROM TempAspects ORDER BY TempAspects.[����� �������];";
LT->Active=true;

Table* ST=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

ST->SetCommandText("SELECT TempAspects.[����� �������], TempAspects.ServerNum FROM TempAspects ORDER BY TempAspects.[����� �������];");
ST->Active(true);

Zast->MClient->LoadTable(ST, LT);

if(Zast->MClient->VerifyTable(LT, ST)==0)
{
for(LT->First();!LT->Eof;LT->Next())
{
 int N=LT->FieldByName("����� �������")->Value;

 if(Asp->Locate("����� �������",N, SO))
 {
  Asp->Edit();
  Asp->FieldByName("ServerNum")->Value=LT->FieldByName("ServerNum")->Value;
  Asp->Post();
 }
 else
 {
 ShowMessage("������ ���������");
 }
}


}
else
{
 ShowMessage("������ ��������");
}
Zast->MClient->DeleteTable(this, ST);
}
catch(...)
{
return false;
}
}
Zast->MClient->DeleteTable(this, TempPodr);
}
Zast->MClient->DeleteTable(this, SMemo4);

}
Zast->MClient->DeleteTable(this, SMemo3);

}
Zast->MClient->DeleteTable(this, SMemo2);

}
Zast->MClient->DeleteTable(this, SMemo1);
}
Zast->MClient->DeleteTable(this, STemp);
Zast->MClient->WriteDiaryEvent("NetAspects","����� ������ ��������","");
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ ������ ��������"," ������ "+IntToStr(GetLastError()));

}

return Ret;
*/
}
//------------------------------------------------------------
void __fastcall TForm1::N4Click(TObject *Sender)
{
/*
this->Hide();
Vvedenie->Show();
*/
}
//---------------------------------------------------------------------------


void __fastcall TForm1::N5Click(TObject *Sender)
{
/*
this->Hide();
Metodika->Show();
*/
}
//---------------------------------------------------------------------------

void __fastcall TForm1::WMSysCommand(TMessage & Msg)
{
  switch (Msg.WParam)
  {
/*
case SC_MINIMIZE:
{
 Shell_NotifyIcon(NIM_ADD, &NID);
  BCreate->Enabled = false;
 BDelete->Enabled = true;
 BHide->Enabled = true;
Form1->Visible=false;
//ShowMessage("SC_MINIMIZE");
                     break;
}
*/
case SC_MINIMIZE:
{

//  Shell_NotifyIcon(NIM_ADD, &NID);
//  BCreate->Enabled = false;
// BDelete->Enabled = true;
// BHide->Enabled = true;
//Form1->Visible=false;
//SendMessage(Handle,WM_SYSCOMMAND,SC_MINIMIZE,0);

//ShowMessage("SC_Close");
Application->Minimize();
break;
}

default:
DefWindowProc(Handle,Msg.Msg,Msg.WParam,Msg.LParam);

 }

}
//-------------------------------------------------------------------
bool TForm1::LoadPodr()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","������ �������� �������������","");
try
{
MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="SELECT Temp�������������.[����� �������������], Temp�������������.[�������� �������������] FROM Temp������������� order by [����� �������������];";
Temp->Active=true;

Table* SPodr=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

SPodr->SetCommandText("SELECT �������������.[����� �������������], �������������.[�������� �������������] FROM ������������� order by [����� �������������];");
SPodr->Active(true);

Zast->MClient->LoadTable(Temp, SPodr);

if(Zast->MClient->VerifyTable(Temp, SPodr)==0)
{
MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOUsrAspect;
Podr->CommandText="SELECT * FROM �������������;";
Podr->Active=true;

for(Podr->First();!Podr->Eof;Podr->Next())
{
 int N=Podr->FieldByName("ServerNum")->Value;
 if(Temp->Locate("����� �������������", N, SO))
 {
  Podr->Edit();
  Podr->FieldByName("�������� �������������")->Value=Temp->FieldByName("�������� �������������")->Value;
  Podr->FieldByName("ServerNum")->Value=Temp->FieldByName("����� �������������")->Value;
  Podr->Post();

  Temp->Delete();
 }
 else
 {
  Podr->Edit();
  Podr->FieldByName("Del")->Value=true;
  Podr->Post();
 }
}

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From ������������� Whete Del=true";
Comm->Execute();

Comm->CommandText="INSERT INTO ������������� ( ServerNum, [�������� �������������] ) SELECT Temp�������������.[����� �������������], Temp�������������.[�������� �������������] FROM Temp�������������;";
Comm->Execute();

Comm->CommandText="Delete * From Temp�������������";
Comm->Execute();
Zast->MClient->WriteDiaryEvent("NetAspects","����� �������� �������������","");
Ret=true;
}
Zast->MClient->DeleteTable(this, SPodr);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects","���� �������� �������������","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ������","������ �������� �������������"," ������ "+IntToStr(GetLastError()));
return false;
}
return Ret;
*/
}
//--------------------------------------------------------------------
void __fastcall TForm1::N7Click(TObject *Sender)
{
//FAbout->ShowModal();
}
//---------------------------------------------------------------------------
void TForm1::EnabledForm(bool B)
{
Label2->Enabled=B;
CPodrazdel->Enabled=B;
Label3->Enabled=B;
CSit->Enabled=B;
Label4->Enabled=B;
Edit1->Enabled=B;
if (!B) Edit1->Text="";
Button1->Enabled=B;
Label5->Enabled=B;
Edit2->Enabled=B;
if (!B) Edit2->Text="";
Button2->Enabled=B;
Label6->Enabled=B;
DBEdit2->Enabled=B;
if (!B) DBEdit2->Text="";
Label7->Enabled=B;
Edit4->Enabled=B;
if (!B) Edit4->Text="";
Button3->Enabled=B;
Label8->Enabled=B;
Edit5->Enabled=B;
if (!B) Edit5->Text="";
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
Label20->Enabled=B;
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
Label33->Enabled=B;
BitBtn7->Enabled=B;
Label34->Enabled=B;
Label35->Enabled=B;
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


