//---------------------------------------------------------------------------

#include <vcl.h>
#include "Reports.h"
#pragma hdrstop

#include "Reports.h"
#include "MasterPointer.h"
#include "Zastavka.h"
#include "IdGlobal.hpp"
#include "MainForm.h"
#include "Rep1.h"
//---------------------------------------------------------------------------

#pragma package(smart_init)
void Reports::CreateReport1(int Podr, TDateTime Date1, TDateTime Date2, AnsiString Isp, String Filtr1, String LFiltr)
{
AnsiString G1;
if(Role==4 | Role==3)
{
PrepareDemoReport();
String T1="";
if(Role==4)
{
T1="TOP 2";
}
else
{
T1="";
}
if (Filtr1=="")
{
G1="SELECT "+T1+" TempAspects.* FROM TempAspects INNER JOIN ������������� ON TempAspects.������������� = �������������.[����� �������������] WHERE (((�������������.ServerNum)="+IntToStr(Podr)+")) ORDER BY TempAspects.[����� �������];";
}
else
{
G1="SELECT "+T1+" TempAspects.* FROM TempAspects INNER JOIN ������������� ON TempAspects.������������� = �������������.[����� �������������] WHERE  "+Filtr1+" AND (((�������������.ServerNum)="+IntToStr(Podr)+")) ORDER BY TempAspects.[����� �������];";
}

}
else
{
PrepareReport();

if (Filtr1=="")
{
G1="select * from TempAspects Where �������������="+IntToStr(Podr)+" Order By [����� �������]";
}
else
{
G1="select * from TempAspects Where "+Filtr1+" AND �������������="+IntToStr(Podr)+" Order By [����� �������]";
}
}

MP<TADODataSet>Report(Zast);
Report->Connection=Connect;

MP<TADODataSet>TempAspects(Zast);
TempAspects->Connection=Connect;

Report->Active=false;


Report->CommandText=G1;
Report->Active=true;

AnsiString T2="�-001_1 ";
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������������� Where [ServerNum]="+IntToStr(Podr);
TempAspects->Active=true;
AnsiString NP=TempAspects->FieldByName("�������� �������������")->Value;
T2=T2+" �������� '"+NP+"'";
AnsiString PP1=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\�-001_1.xlt");
AnsiString PP2=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\"+T2+".xlt");
//CopyFile(P1.c_str() ,P2.c_str() , false);
CopyFileTo(PP1.c_str() ,PP2.c_str());

Variant App =Variant::CreateObject("Excel.Application");
App.OlePropertySet("Visible",false);
Variant Book;


Book=App.OlePropertyGet("Workbooks").OleFunction("Add", PP2.c_str());


Variant Sheet=App.OlePropertyGet("ActiveSheet");
Sheet.OlePropertySet("Name","�-001.1");

DeleteFile(PP2);


/*
int Start=17;
AnsiString Text;
int Num;
T="�������� ������ c "+Date1.DateString()+" �� "+Date2.DateString();
App.OlePropertyGet("Cells",9,1).OlePropertySet("Value",T.c_str());
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������������� Where [ServerNum]="+IntToStr(Podr);
TempAspects->Active=true;
NP=TempAspects->FieldByName("�������� �������������")->Value;
T=NP;
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value",T.c_str());
T="������ - "+LFiltr;
App.OlePropertyGet("Cells",12,1).OlePropertySet("Value",T.c_str());

Report->First();
int Number=0;
for(int i=0;i<Report->RecordCount;i++)
{
Number=i;
App.OlePropertyGet("Cells",Start+i,1).OlePropertySet("Value",i+1);
Num=Report->FieldByName("������������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������������ Where [����� ������������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ ������������")->AsString+" ";
App.OlePropertyGet("Cells",Start+i,2).OlePropertySet("Value",Text.c_str());

Variant V=Report->FieldByName("�������������")->Value;
if (V.IsNull()==true)
{
Text="";
}
else
{
Text=Report->FieldByName("�������������")->Value;
}
App.OlePropertyGet("Cells",Start+i,3).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������ Where [����� �������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ �������")->Value;
App.OlePropertyGet("Cells",Start+i,4).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("�����������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ����������� Where [����� �����������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ �����������")->Value;
App.OlePropertyGet("Cells",Start+i,5).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("��������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From �������� Where [����� ��������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("�������� ��������")->AsString;
App.OlePropertyGet("Cells",Start+i,6).OlePropertySet("Value",Text.c_str());

if(Report->FieldByName("G")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,7).OlePropertySet("Value",2);
}
else
{
App.OlePropertyGet("Cells",Start+i,7).OlePropertySet("Value",0);
}

if(Report->FieldByName("O")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,8).OlePropertySet("Value",1);
}
else
{
App.OlePropertyGet("Cells",Start+i,8).OlePropertySet("Value",0);
}

if(Report->FieldByName("R")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,9).OlePropertySet("Value",1);
}
else
{
App.OlePropertyGet("Cells",Start+i,9).OlePropertySet("Value",0);
}

if(Report->FieldByName("S")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,10).OlePropertySet("Value",2);
}
else
{
App.OlePropertyGet("Cells",Start+i,10).OlePropertySet("Value",0);
}

if(Report->FieldByName("T")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,11).OlePropertySet("Value",2);
}
else
{
App.OlePropertyGet("Cells",Start+i,11).OlePropertySet("Value",0);
}

if(Report->FieldByName("L")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,12).OlePropertySet("Value",1);
}
else
{
App.OlePropertyGet("Cells",Start+i,12).OlePropertySet("Value",0);
}

if(Report->FieldByName("N")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,13).OlePropertySet("Value",1);
}
else
{
App.OlePropertyGet("Cells",Start+i,13).OlePropertySet("Value",0);
}


Num=Report->FieldByName("������� �����������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From [������� �����������] Where [����� �����������]="+IntToStr(Num);
TempAspects->Active=true;
Num=TempAspects->FieldByName("����")->Value;
App.OlePropertyGet("Cells",Start+i,14).OlePropertySet("Value",Num);

Num=Report->FieldByName("���������� �����������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From [���������� �����������] Where [����� ����������]="+IntToStr(Num);
TempAspects->Active=true;
Num=TempAspects->FieldByName("����")->Value;
App.OlePropertyGet("Cells",Start+i,15).OlePropertySet("Value",Num);

App.OlePropertyGet("Cells",Start+i,16).OlePropertySet("Value",Report->FieldByName("Z")->Value);


bool Zn=Report->FieldByName("����������")->Value;
if (Zn==true)
{

Text="��������";
}
else
{

Text="����������";
}
App.OlePropertyGet("Cells",Start+i,17).OlePropertySet("Value",Text.c_str());



Report->Next();
}

App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+Number)).c_str()).OlePropertySet("VerticalAlignment",-4160);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+Number)).c_str()).OlePropertySet("WrapText",true);
//App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,6,Start+i-1)).c_str()).OlePropertyGet("Borders",7).OlePropertySet("Weight",-4138);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+Number)).c_str()).OlePropertyGet("Borders",7).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+Number)).c_str()).OlePropertyGet("Borders",8).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+Number)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+Number)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+Number)).c_str()).OlePropertyGet("Borders",11).OlePropertySet("Weight",4);
if (Number>=1)
{
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+Number)).c_str()).OlePropertyGet("Borders",12).OlePropertySet("Weight",4);
}

T="��������: _________________";
App.OlePropertyGet("Cells",Start+Number+4,2).OlePropertySet("Value",T.c_str());
App.OlePropertyGet("Cells",Start+Number+4,3).OlePropertySet("Value",Isp.c_str());
App.OlePropertyGet("Range",(Address(Sheet,3,Start+Number+4)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+Number+4)).c_str()).OlePropertySet("HorizontalAlignment",-4152);
App.OlePropertyGet("Cells",Start+Number+4,4).OlePropertySet("Value","________________________");

App.OlePropertyGet("Cells",Start+Number+5,2).OlePropertySet("Value","���������");
App.OlePropertyGet("Cells",Start+Number+5,2).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,2,Start+Number+5)).c_str()).OlePropertySet("IndentLevel",9);

App.OlePropertyGet("Cells",Start+Number+5,3).OlePropertySet("Value","�.�.�.");
App.OlePropertyGet("Cells",Start+Number+5,3).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,3,Start+Number+5)).c_str()).OlePropertySet("IndentLevel",7);

App.OlePropertyGet("Cells",Start+Number+5,4).OlePropertySet("Value","�������");
App.OlePropertyGet("Cells",Start+Number+5,4).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+Number+5)).c_str()).OlePropertySet("IndentLevel",7);

AnsiString Patch=ExtractFilePath(Application->ExeName);
AnsiString NameReport=Patch+"\Reports\\�-001.1.xls";
App.OlePropertySet("Visible",true);
//Application->BringToFront();

App=NULL;
Book=NULL;
Sheet=NULL;
*/
MP<TADODataSet>Posledstvie(Zast);
Posledstvie->Connection=Connect;

MP<TADODataSet>Tiagest(Zast);
Tiagest->Connection=Connect;

MP<TADODataSet>Prioritet(Zast);
Prioritet->Connection=Connect;

int Start=17;
AnsiString Text;
int Num;
AnsiString Ts;
//Ts="������ ������ ���������������� ����������  c "+Date1.DateString()+" �� "+Date2.DateString();
Ts="������ ������ ���������������� ����������";
App.OlePropertyGet("Cells",9,1).OlePropertySet("Value",Ts.c_str());
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������������� Where [����� �������������]="+IntToStr(Podr);
TempAspects->Active=true;
NP=TempAspects->FieldByName("�������� �������������")->Value;
Ts=NP;
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value",Ts.c_str());
Ts="������ - "+LFiltr;
App.OlePropertyGet("Cells",12,1).OlePropertySet("Value",Ts.c_str());

Report->First();
double G, O, R, S, T, L, N;
double C, B, F, Tp, Pr, UVZ, VVR;
int s1;
double s2, s3;
int p1;
double p2, p3;
int Number;
for(int i=0;i<Report->RecordCount;i++)
{
Number=i;
Num=Report->FieldByName("��� ����������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ���������� Where [����� ����������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ ����������")->Value;
App.OlePropertyGet("Cells",Start+i,1).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("������������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������������ Where [����� ������������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ ������������")->Value;
App.OlePropertyGet("Cells",Start+i,2).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������ Where [����� �������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ �������")->Value;
App.OlePropertyGet("Cells",Start+i,3).OlePropertySet("Value",Text.c_str());

Variant V=Report->FieldByName("�������������")->Value;
if (V.IsNull()==true)
{
Text="";
}
else
{
Text=Report->FieldByName("�������������")->Value;
}
App.OlePropertyGet("Cells",Start+i,4).OlePropertySet("Value",Text.c_str());



Num=Report->FieldByName("�����������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ����������� Where [����� �����������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ �����������")->Value;
App.OlePropertyGet("Cells",Start+i,5).OlePropertySet("Value",Text.c_str());


Num=Report->FieldByName("��������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From �������� Where [����� ��������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("�������� ��������")->Value;
App.OlePropertyGet("Cells",Start+i,6).OlePropertySet("Value",Text.c_str());





if(Report->FieldByName("G")->AsBoolean)
{
G=0.2;
}
else
{
G=0;
}

if(Report->FieldByName("O")->AsBoolean)
{
O=0.2;
}
else
{
O=0;
}

if(Report->FieldByName("R")->AsBoolean)
{
R=0.1;
}
else
{
R=0;
}

if(Report->FieldByName("S")->AsBoolean)
{
S=0.1;
}
else
{
S=0;
}

if(Report->FieldByName("T")->AsBoolean)
{
T=0.2;
}
else
{
T=0;
}

if(Report->FieldByName("L")->AsBoolean)
{
L=0.2;
}
else
{
L=0;
}

if(Report->FieldByName("N")->AsBoolean)
{
N=0.1;
}
else
{
N=0;
}



Num=Report->FieldByName("���������� �����������")->Value;
 Posledstvie->Active=false;
 Posledstvie->CommandText="Select * From [���������� �����������] where [����� ����������]="+IntToStr(Num);
 Posledstvie->Active=true;
 Posledstvie->First();
// Posledstvie->MoveBy(CProyav->ItemIndex);
 F=Posledstvie->FieldByName("����")->Value;

Num=Report->FieldByName("������� �����������")->Value;
 Tiagest->Active=false;
 Tiagest->CommandText="Select * From [������� �����������] where [����� �����������]="+IntToStr(Num);
 Tiagest->Active=true;
 Tiagest->First();
// Tiagest->MoveBy(CPosl->ItemIndex);
 UVZ=Tiagest->FieldByName("����")->Value;

Num=Report->FieldByName("��������������")->Value;
Prioritet->Active=false;
Prioritet->CommandText="Select * From ��� where [����� ��������������]="+IntToStr(Num);
Prioritet->Active=true;
Prioritet->First();
//Prioritet->MoveBy(CPrior->ItemIndex);
VVR=Prioritet->FieldByName("����")->Value;



 C=G+O+R+S+T+L+N;

if(C==0)
{
C=0.1;
}
 B=C*F;
 Tp=UVZ/VVR;
 Pr=Tp*B;

s1=Tp*100;
s2=s1;
s3=s2/100;
if(Tp-s3<0.01)
{
if(Tp-s3>=0.005)
{
s3=s3+0.01;
}
}

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

Text=FloatToStr(F);
App.OlePropertyGet("Cells",Start+i,7).OlePropertySet("Value",Text.c_str());

Text=FloatToStr(C);
App.OlePropertyGet("Cells",Start+i,8).OlePropertySet("Value",Text.c_str());

Text=FloatToStr(B);
App.OlePropertyGet("Cells",Start+i,9).OlePropertySet("Value",Text.c_str());

Text=FloatToStr(UVZ);
App.OlePropertyGet("Cells",Start+i,10).OlePropertySet("Value",Text.c_str());

Text=FloatToStr(VVR);
App.OlePropertyGet("Cells",Start+i,11).OlePropertySet("Value",Text.c_str());

Text=FloatToStr(Tp);
App.OlePropertyGet("Cells",Start+i,12).OlePropertySet("Value",Text.c_str());

Text=FloatToStr(p3);
App.OlePropertyGet("Cells",Start+i,13).OlePropertySet("Value",Text.c_str());

Text=Report->FieldByName("������������ ����������")->AsString;;
App.OlePropertyGet("Cells",Start+i,15).OlePropertySet("Value",Text.c_str());

/*
Num=Report->FieldByName("������� �����������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From [������� �����������] Where [����� �����������]="+IntToStr(Num);
TempAspects->Active=true;
Num=TempAspects->FieldByName("����")->Value;
App.OlePropertyGet("Cells",Start+i,14).OlePropertySet("Value",Num);

Num=Report->FieldByName("���������� �����������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From [���������� �����������] Where [����� ����������]="+IntToStr(Num);
TempAspects->Active=true;
Num=TempAspects->FieldByName("����")->Value;
App.OlePropertyGet("Cells",Start+i,15).OlePropertySet("Value",Num);

App.OlePropertyGet("Cells",Start+i,16).OlePropertySet("Value",Report->FieldByName("Z")->Value);
*/

bool Zn=Report->FieldByName("����������")->Value;
if (Zn==true)
{

Text="��������";
}
else
{

Text="����������";
}
App.OlePropertyGet("Cells",Start+i,14).OlePropertySet("Value",Text.c_str());



Report->Next();
}

App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,15,Start+Number)).c_str()).OlePropertyGet("Borders",7).OlePropertySet("Weight",-4138);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,15,Start+Number)).c_str()).OlePropertyGet("Borders",8).OlePropertySet("Weight",-4138);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,15,Start+Number)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",-4138);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,15,Start+Number)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",-4138);

App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,15,Start+Number)).c_str()).OlePropertyGet("Borders",11).OlePropertySet("Weight",2);

App.OlePropertyGet("Range",(Address(Sheet,6,Start)+":"+Address(Sheet,6,Start+Number)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",-4138);
App.OlePropertyGet("Range",(Address(Sheet,9,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",-4138);
App.OlePropertyGet("Range",(Address(Sheet,12,Start)+":"+Address(Sheet,12,Start+Number)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",-4138);
//App.OlePropertyGet("Range",(Address(Sheet,13,Start)+":"+Address(Sheet,13,Start+Number)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",-4138);

if (Number>=1)
{
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,15,Start+Number)).c_str()).OlePropertyGet("Borders",12).OlePropertySet("Weight",2);
}

App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,15,Start+Number)).c_str()).OlePropertySet("VerticalAlignment",-4160);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,15,Start+Number)).c_str()).OlePropertySet("WrapText",true);
App.OlePropertyGet("Range",(Address(Sheet,7,Start)+":"+Address(Sheet,13,Start+Number)).c_str()).OlePropertySet("HorizontalAlignment",-4108);


Ts="��������: _____________";
App.OlePropertyGet("Cells",Start+Number+4,1).OlePropertySet("Value",Ts.c_str());
App.OlePropertyGet("Cells",Start+Number+4,2).OlePropertySet("Value",Isp.c_str());
App.OlePropertyGet("Range",(Address(Sheet,2,Start+Number+4)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,3,Start+Number+4)).c_str()).OlePropertySet("HorizontalAlignment",-4152);
App.OlePropertyGet("Cells",Start+Number+4,3).OlePropertySet("Value","    ______________________");

App.OlePropertyGet("Cells",Start+Number+5,1).OlePropertySet("Value","���������");
App.OlePropertyGet("Cells",Start+Number+5,1).OlePropertyGet("Font").OlePropertySet("Size",8);
//App.OlePropertyGet("Range",(Address(Sheet,1,Start+Number+5)).c_str()).OlePropertySet("IndentLevel",9);

App.OlePropertyGet("Cells",Start+Number+5,2).OlePropertySet("Value","�.�.�.");
App.OlePropertyGet("Cells",Start+Number+5,2).OlePropertyGet("Font").OlePropertySet("Size",8);
//App.OlePropertyGet("Range",(Address(Sheet,2,Start+Number+5)).c_str()).OlePropertySet("IndentLevel",7);

App.OlePropertyGet("Cells",Start+Number+5,3).OlePropertySet("Value","�������");
App.OlePropertyGet("Cells",Start+Number+5,3).OlePropertyGet("Font").OlePropertySet("Size",8);
//App.OlePropertyGet("Range",(Address(Sheet,3,Start+Number+5)).c_str()).OlePropertySet("IndentLevel",7);

App.OlePropertyGet("Cells",Start+Number+6,1).OlePropertySet("Value","\"___\" _________ 20___ �");
App.OlePropertyGet("Cells",Start+Number+6,1).OlePropertyGet("Font").OlePropertySet("Size",10);
//App.OlePropertyGet("Range",(Address(Sheet,1,Start+Number+6)).c_str()).OlePropertySet("IndentLevel",7);

App.OlePropertyGet("Range",(Address(Sheet,1,Start+Number+5)+":"+Address(Sheet,3,Start+Number+5)).c_str()).OlePropertySet("HorizontalAlignment",-4108);


AnsiString Patch=ExtractFilePath(Application->ExeName);
AnsiString NameReport=Patch+"\Reports\\�-001.1.xls";
App.OlePropertySet("Visible",true);
//Application->BringToFront();

App=NULL;
Book=NULL;
Sheet=NULL;
}
/*
else
{
ShowMessage("��� ������ ��� ������");
}
*/
//}
//------------------------------------------------------------------------------------------------------------------
void Reports::CreateReport2(int Podr, TDateTime Date1, TDateTime Date2, AnsiString Isp, String Filtr2, String LFiltr)
{
AnsiString G;
if(Filtr2=="����������=False")
{
Filtr2="����������=True";
}

if(Role==4 | Role==3)
{
PrepareDemoReport();
String T;
if(Role==4)
{
T="TOP 2";
}
else
{
T="";
}
if (Filtr2=="")
{
G="SELECT "+T+" TempAspects.* FROM TempAspects INNER JOIN ������������� ON TempAspects.������������� = �������������.[����� �������������] WHERE (((�������������.ServerNum)="+IntToStr(Podr)+")) ORDER BY TempAspects.[����� �������];";
}
else
{
G="SELECT "+T+" TempAspects.* FROM TempAspects INNER JOIN ������������� ON TempAspects.������������� = �������������.[����� �������������] WHERE  "+Filtr2+" AND (((�������������.ServerNum)="+IntToStr(Podr)+")) ORDER BY TempAspects.[����� �������];";
}
}
else
{
PrepareReport();

if(Filtr2=="")
{
G="Select * From TempAspects Where  �������������="+IntToStr(Podr)+" AND ����������=True Order By [����� �������]";
}
else
{
G="Select * From TempAspects Where "+Filtr2+" AND �������������="+IntToStr(Podr)+" AND ����������=True Order By [����� �������]";
}
}

MP<TADODataSet>Report(Zast);
Report->Connection=Connect;

MP<TADODataSet>TempAspects(Zast);
TempAspects->Connection=Connect;

Report->Active=false;

Report->CommandText=G;
Report->Active=true;

AnsiString T="�-001.2 ";
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������������� Where [ServerNum]="+IntToStr(Podr);
TempAspects->Active=true;
AnsiString NP=TempAspects->FieldByName("�������� �������������")->Value;
T=T+" �������� '"+NP+"'";

AnsiString P1=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\�-001_2.xlt");
AnsiString P2=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\"+T+".xlt");
CopyFile(P1.c_str() ,P2.c_str() , false);

Variant App =Variant::CreateObject("Excel.Application");
App.OlePropertySet("Visible",false);
Variant Book=App.OlePropertyGet("Workbooks").OleFunction("Add", P2.c_str());
Variant Sheet=App.OlePropertyGet("ActiveSheet");
Sheet.OlePropertySet("Name","�-001.2");
App.OlePropertySet("Visible",false);

DeleteFile(P2);

/*
int Start=16;
AnsiString Text;
int Num;
T="������ �������� ������ � "+Date1.DateString()+" �� "+Date2.DateString();
App.OlePropertyGet("Cells",9,1).OlePropertySet("Value",T.c_str());
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������������� Where [ServerNum]="+IntToStr(Podr);
TempAspects->Active=true;
NP=TempAspects->FieldByName("�������� �������������")->Value;
T=NP;
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value",T.c_str());

T="������ - "+LFiltr;
App.OlePropertyGet("Cells",12,1).OlePropertySet("Value",T.c_str());

Report->First();
int Number=0;
for(int i=0;i<Report->RecordCount;i++)
{
Number=i;
App.OlePropertyGet("Cells",Start+i,1).OlePropertySet("Value",i+1);

Num=Report->FieldByName("������������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������������ Where [����� ������������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ ������������")->Value;
App.OlePropertyGet("Cells",Start+i,2).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������ Where [����� �������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ �������")->Value;
App.OlePropertyGet("Cells",Start+i,3).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("�����������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ����������� Where [����� �����������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ �����������")->Value;
App.OlePropertyGet("Cells",Start+i,4).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("��������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From �������� Where [����� ��������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("�������� ��������")->Value;
App.OlePropertyGet("Cells",Start+i,5).OlePropertySet("Value",Text.c_str());

int Z=Report->FieldByName("Z")->Value;
App.OlePropertyGet("Cells",Start+i,6).OlePropertySet("Value",Z);

AnsiString ZT=ClearString(Report->FieldByName("���������� � ��������")->AsString);
App.OlePropertyGet("Cells",Start+i,7).OlePropertySet("Value",ZT.c_str());

ZT=ClearString(Report->FieldByName("������������ ���������� � ��������")->AsString);

App.OlePropertyGet("Cells",Start+i,8).OlePropertySet("Value",ZT.c_str());

ZT=ClearString(Report->FieldByName("������������ �����������")->AsString);

App.OlePropertyGet("Cells",Start+i,9).OlePropertySet("Value",ZT.c_str());

Report->Next();
}
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertySet("VerticalAlignment",-4160);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertySet("WrapText",true);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertyGet("Borders",7).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertyGet("Borders",8).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertyGet("Borders",11).OlePropertySet("Weight",4);
if (Number>=1)
{
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertyGet("Borders",12).OlePropertySet("Weight",4);
}

App.OlePropertyGet("Cells",Start+Number+3,2).OlePropertySet("Value","�����������:");
App.OlePropertyGet("Range",(Address(Sheet,2,Start+Number+5)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+Number+6,2).OlePropertySet("Value","(���������)");
App.OlePropertyGet("Cells",Start+Number+6,2).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,2,Start+Number+6)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,2,Start+Number+8)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+Number+9,2).OlePropertySet("Value","(���������)");
App.OlePropertyGet("Cells",Start+Number+9,2).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,2,Start+Number+9)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,4,Start+Number+5)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+Number+6,4).OlePropertySet("Value","(�������)");
App.OlePropertyGet("Cells",Start+Number+6,4).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+Number+6)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,4,Start+Number+8)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+Number+9,4).OlePropertySet("Value","(�������)");
App.OlePropertyGet("Cells",Start+Number+9,4).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+Number+9)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+5)+":"+Address(Sheet,7,Start+Number+5)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+5)+":"+Address(Sheet,7,Start+Number+5)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+6)+":"+Address(Sheet,7,Start+Number+6)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Cells",Start+Number+6,6).OlePropertySet("Value","(�.�.�.)");
App.OlePropertyGet("Cells",Start+Number+6,6).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+6)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+8)+":"+Address(Sheet,7,Start+Number+8)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+8)+":"+Address(Sheet,7,Start+Number+8)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+9)+":"+Address(Sheet,7,Start+Number+9)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Cells",Start+Number+9,6).OlePropertySet("Value","(�.�.�.)");
App.OlePropertyGet("Cells",Start+Number+9,6).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+9)).c_str()).OlePropertySet("HorizontalAlignment",-4108);


App.OlePropertySet("Visible",true);
//Application->BringToFront();

App=NULL;
Book=NULL;
Sheet=NULL;
*/

int Start=17;
AnsiString Text;
int Num;
//T="������ �������� ������ ���������������� ���������� � "+Date1.DateString()+" �� "+Date2.DateString();
T="������ �������� ������ ���������������� ����������";//+Date1.DateString()+" �� "+Date2.DateString();

App.OlePropertyGet("Cells",9,1).OlePropertySet("Value",T.c_str());
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������������� Where [����� �������������]="+IntToStr(Podr);
TempAspects->Active=true;
NP=TempAspects->FieldByName("�������� �������������")->Value;
T=NP;
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value",T.c_str());

T="������ - "+LFiltr;
App.OlePropertyGet("Cells",12,1).OlePropertySet("Value",T.c_str());

Report->First();
int Number=0;
for(int i=0;i<Report->RecordCount;i++)
{
Number=i;

Num=Report->FieldByName("��� ����������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ���������� Where [����� ����������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ ����������")->Value;
App.OlePropertyGet("Cells",Start+i,1).OlePropertySet("Value",Text.c_str());

//App.OlePropertyGet("Cells",Start+i,1).OlePropertySet("Value",i+1);

Num=Report->FieldByName("������������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������������ Where [����� ������������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ ������������")->Value;
App.OlePropertyGet("Cells",Start+i,2).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������ Where [����� �������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ �������")->Value;
App.OlePropertyGet("Cells",Start+i,3).OlePropertySet("Value",Text.c_str());

App.OlePropertyGet("Range",(Address(Sheet,7,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

Text=Report->FieldByName("�������������")->AsString;
if(!Text.IsEmpty())
{
App.OlePropertyGet("Cells",Start+i,4).OlePropertySet("Value",Text.c_str());
}


Num=Report->FieldByName("�����������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From ����������� Where [����� �����������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ �����������")->Value;
App.OlePropertyGet("Cells",Start+i,5).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("��������")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From �������� Where [����� ��������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("�������� ��������")->Value;
App.OlePropertyGet("Cells",Start+i,6).OlePropertySet("Value",Text.c_str());

int Z=Report->FieldByName("Z")->Value;
App.OlePropertyGet("Cells",Start+i,7).OlePropertySet("Value",Z);

Text=Report->FieldByName("������������ ����������")->AsString;;
App.OlePropertyGet("Cells",Start+i,8).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("�����")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From �������������� Where [����� ��������������]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("������������ ��������������")->Value;
App.OlePropertyGet("Cells",Start+i,9).OlePropertySet("Value",Text.c_str());




AnsiString ZT=Report->FieldByName("������������� �����������")->Value;
while (ZT.Pos("\r"))
{
 int N=ZT.Pos("\r");
 AnsiString Z=ZT.Delete(N,1);
 ZT=Z;
}
App.OlePropertyGet("Cells",Start+i,10).OlePropertySet("Value",ZT.c_str());

ZT=Report->FieldByName("������������ �����������")->Value;
while (ZT.Pos("\r"))
{
 int N=ZT.Pos("\r");
 AnsiString Z=ZT.Delete(N,1);
 ZT=Z;
}
App.OlePropertyGet("Cells",Start+i,11).OlePropertySet("Value",ZT.c_str());

ZT=Report->FieldByName("���������� � ��������")->Value;
while (ZT.Pos("\r"))
{
 int N=ZT.Pos("\r");
 AnsiString Z=ZT.Delete(N,1);
 ZT=Z;
}
App.OlePropertyGet("Cells",Start+i,12).OlePropertySet("Value",ZT.c_str());

ZT=Report->FieldByName("������������ ���������� � ��������")->Value;
while (ZT.Pos("\r"))
{
 int N=ZT.Pos("\r");
 AnsiString Z=ZT.Delete(N,1);
 ZT=Z;
}
App.OlePropertyGet("Cells",Start+i,13).OlePropertySet("Value",ZT.c_str());

Report->Next();
}
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,13,Start+Number)).c_str()).OlePropertySet("VerticalAlignment",-4160);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,13,Start+Number)).c_str()).OlePropertySet("WrapText",true);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,13,Start+Number)).c_str()).OlePropertyGet("Borders",7).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,13,Start+Number)).c_str()).OlePropertyGet("Borders",8).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,13,Start+Number)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,13,Start+Number)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,13,Start+Number)).c_str()).OlePropertyGet("Borders",11).OlePropertySet("Weight",4);
if (Number>=1)
{
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,13,Start+Number)).c_str()).OlePropertyGet("Borders",12).OlePropertySet("Weight",4);
}

String Ts="��������: _____________";
App.OlePropertyGet("Cells",Start+Number+2,1).OlePropertySet("Value",Ts.c_str());
App.OlePropertyGet("Cells",Start+Number+2,2).OlePropertySet("Value",Isp.c_str());
App.OlePropertyGet("Range",(Address(Sheet,2,Start+Number+2)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,3,Start+Number+2)).c_str()).OlePropertySet("HorizontalAlignment",-4152);
App.OlePropertyGet("Cells",Start+Number+2,3).OlePropertySet("Value","    ______________________");

App.OlePropertyGet("Cells",Start+Number+3,1).OlePropertySet("Value","���������");
App.OlePropertyGet("Cells",Start+Number+3,1).OlePropertyGet("Font").OlePropertySet("Size",8);
//App.OlePropertyGet("Range",(Address(Sheet,1,Start+Number+5)).c_str()).OlePropertySet("IndentLevel",9);

App.OlePropertyGet("Cells",Start+Number+3,2).OlePropertySet("Value","�.�.�.");
App.OlePropertyGet("Cells",Start+Number+3,2).OlePropertyGet("Font").OlePropertySet("Size",8);
//App.OlePropertyGet("Range",(Address(Sheet,2,Start+Number+5)).c_str()).OlePropertySet("IndentLevel",7);

App.OlePropertyGet("Cells",Start+Number+3,3).OlePropertySet("Value","�������");
App.OlePropertyGet("Cells",Start+Number+3,3).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,1,Start+Number+3)+":"+Address(Sheet,3,Start+Number+3)).c_str()).OlePropertySet("HorizontalAlignment",-4108);


App.OlePropertyGet("Cells",Start+Number+5,1).OlePropertySet("Value","�����������:");
App.OlePropertyGet("Range",(Address(Sheet,1,Start+Number+7)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+Number+8,1).OlePropertySet("Value","(���������)");
App.OlePropertyGet("Cells",Start+Number+8,1).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,1,Start+Number+8)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,1,Start+Number+10)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+Number+11,1).OlePropertySet("Value","(���������)");
App.OlePropertyGet("Cells",Start+Number+11,1).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,1,Start+Number+11)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,3,Start+Number+7)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+Number+8,3).OlePropertySet("Value","(�������)");
App.OlePropertyGet("Cells",Start+Number+8,3).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,3,Start+Number+8)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,3,Start+Number+10)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+Number+11,3).OlePropertySet("Value","(�������)");
App.OlePropertyGet("Cells",Start+Number+11,3).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,3,Start+Number+11)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,5,Start+Number+7)+":"+Address(Sheet,6,Start+Number+7)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Range",(Address(Sheet,5,Start+Number+7)+":"+Address(Sheet,6,Start+Number+7)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,5,Start+Number+8)+":"+Address(Sheet,6,Start+Number+8)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Cells",Start+Number+8,5).OlePropertySet("Value","(�.�.�.)");
App.OlePropertyGet("Cells",Start+Number+8,5).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,5,Start+Number+8)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,5,Start+Number+10)+":"+Address(Sheet,6,Start+Number+10)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Range",(Address(Sheet,5,Start+Number+10)+":"+Address(Sheet,6,Start+Number+10)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,5,Start+Number+11)+":"+Address(Sheet,6,Start+Number+11)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Cells",Start+Number+11,5).OlePropertySet("Value","(�.�.�.)");
App.OlePropertyGet("Cells",Start+Number+11,5).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,5,Start+Number+11)).c_str()).OlePropertySet("HorizontalAlignment",-4108);


App.OlePropertySet("Visible",true);
//Application->BringToFront();

App=NULL;
Book=NULL;
Sheet=NULL;


}

//------------------------------------------------------------------------------------------------------------------
AnsiString Reports::Address(Variant Sheet,int x,int y)
{
return Sheet.OlePropertyGet("Cells",y,x).OlePropertyGet("Address");
}
//------------------------------------------------------------------------------------------------------------------
void Reports::PrepareReport()
{
Zast->MClient->Start();
//�������� ����������� �������
bool Ret=false;
MP<TADODataSet>LPodr(Form1);
LPodr->Connection=Connect;
LPodr->CommandText="SELECT �������������.[����� �������������], �������������.[�������� �������������], �������������.ServerNum FROM Logins INNER JOIN (������������� INNER JOIN ObslOtdel ON �������������.[����� �������������] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login;";
LPodr->Active=true;

MP<TADOCommand>Comm(Form1);
Comm->Connection=Connect;
Comm->CommandText="Delete * From TempAspects";
Comm->Execute();

MP<TADODataSet>LTemp(Form1);
LTemp->Connection=Connect;
LTemp->CommandText="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[������������ ����������], TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.�����,   TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;";


Table* RTemp=Zast->MClient->CreateTable(Report1, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

RTemp->SetCommandText("SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[������������ ����������], TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.�����,  TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects;");

Zast->MClient->PrepareLoadAspects(0 ,Form1->DBMemo1->Width, Form1->DBMemo31->Width, Form1->DBMemo2->Width, Form1->DBMemo4->Width);

LTemp->Active=false;
LTemp->Active=true;

RTemp->Active(false);
RTemp->Active(true);

Zast->MClient->LoadTable(RTemp, LTemp);

if(Zast->MClient->VerifyTable(LTemp, RTemp)==0)
{
MP<TADODataSet>Memo1(Form1);
Memo1->Connection=Connect;
Memo1->CommandText="SELECT Memo1.Num, Memo1.NumStr, Memo1.Text FROM Memo1 ORDER BY Memo1.Num, Memo1.NumStr;";
Memo1->Active=true;

Table* LMemo1=Zast->MClient->CreateTable(Report1, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);
LMemo1->SetCommandText("SELECT Memo1.Num, Memo1.NumStr, Memo1.Text FROM Memo1 ORDER BY Memo1.Num, Memo1.NumStr;");
LMemo1->Active(true);

Zast->MClient->LoadTable(LMemo1, Memo1);
Memo1->Active=false;
Memo1->Active=true;
if(Zast->MClient->VerifyTable(LMemo1, Memo1)==0)
{

MP<TADODataSet>Memo2(Form1);
Memo2->Connection=Connect;
Memo2->CommandText="SELECT Memo2.Num, Memo2.NumStr, Memo2.Text FROM Memo2 ORDER BY Memo2.Num, Memo2.NumStr;";
Memo2->Active=true;

Table* LMemo2=Zast->MClient->CreateTable(Report1, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);
LMemo2->SetCommandText("SELECT Memo2.Num, Memo2.NumStr, Memo2.Text FROM Memo2 ORDER BY Memo2.Num, Memo2.NumStr;");
LMemo2->Active(true);

Zast->MClient->LoadTable(LMemo2, Memo2);

if(Zast->MClient->VerifyTable(LMemo2, Memo2)==0)
{

MP<TADODataSet>Memo3(Form1);
Memo3->Connection=Connect;
Memo3->CommandText="SELECT Memo3.Num, Memo3.NumStr, Memo3.Text FROM Memo3 ORDER BY Memo3.Num, Memo3.NumStr;";
Memo3->Active=true;

Table* LMemo3=Zast->MClient->CreateTable(Report1, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);
LMemo3->SetCommandText("SELECT Memo3.Num, Memo3.NumStr, Memo3.Text FROM Memo3 ORDER BY Memo3.Num, Memo3.NumStr;");
LMemo3->Active(true);

Zast->MClient->LoadTable(LMemo3, Memo3);

if(Zast->MClient->VerifyTable(LMemo3, Memo3)==0)
{

MP<TADODataSet>Memo4(Form1);
Memo4->Connection=Connect;
Memo4->CommandText="SELECT Memo4.Num, Memo4.NumStr, Memo4.Text FROM Memo4 ORDER BY Memo4.Num, Memo4.NumStr;";
Memo4->Active=true;

Table* LMemo4=Zast->MClient->CreateTable(Report1, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);
LMemo4->SetCommandText("SELECT Memo4.Num, Memo4.NumStr, Memo4.Text FROM Memo4 ORDER BY Memo4.Num, Memo4.NumStr;");
LMemo4->Active(true);

Zast->MClient->LoadTable(LMemo4, Memo4);

if(Zast->MClient->VerifyTable(LMemo4, Memo4)==0)
{


PrepareMergeAspects();



Ret=true;

}
else
{
Ret=Ret & false;

}
Zast->MClient->DeleteTable(Report1, LMemo4);

}
else
{
Ret=Ret & false;

}
Zast->MClient->DeleteTable(Report1, LMemo3);

}
else
{
Ret=Ret & false;

}
Zast->MClient->DeleteTable(Report1, LMemo2);

}
else
{
Ret=Ret & false;

}
Zast->MClient->DeleteTable(Report1, LMemo1);
}
else
{
Ret=Ret & false;

}

Zast->MClient->DeleteTable(Report1, RTemp);
Zast->MClient->Stop();

}
//--------------------------------------------------------
void Reports::PrepareMergeAspects()
{
MP<TADODataSet>Temp(Form1);
Temp->Connection=Connect;
Temp->CommandText="Select * From TempAspects order by [����� �������]";
Temp->Active=true;

MP<TADODataSet>Memo(Form1);
Memo->Connection=Connect;


for(Temp->First();!Temp->Eof;Temp->Next())
{
 int TempNum=Temp->FieldByName("����� �������")->Value;

MP<TMemo>M(Form1);
M->Visible=false;
M->Parent=Form1;
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


}
//---------------------------------------------------------------------------
void Reports::PrepareDemoReport()
{
MP<TADOCommand>Comm(Form1);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * from TempAspects";
Comm->Execute();

/*
MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="select * from TempAspects";
Temp->Active=true;

MP<TADODataSet>Asp(this);
Asp->Connection=Zast->ADOUsrAspect;
As
*/

String S="INSERT INTO TempAspects ( [����� �������], �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, G, O, R, S, T, L, N, Z, ����������, [������������ ����������], [���������� �����������], [������� �����������], ��������������, �����, [������������� �����������], [������������ �����������], [���������� � ��������], [������������ ���������� � ��������], �����������, [���� ��������], [������ ��������], [����� ��������], ServerNum ) ";
S=S+" SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[������������ ����������], �������.[���������� �����������], �������.[������� �����������], �������.��������������, �������.�����, �������.[������������� �����������], �������.[������������ �����������], �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.�����������, �������.[���� ��������], �������.[������ ��������], �������.[����� ��������], �������.ServerNum ";
S=S+" FROM �������;";
Comm->CommandText=S;
Comm->Execute();
}
//--------------------------------------------------------------------------
String Reports::ClearString(String ZT)
{
while (ZT.Pos("\r"))
{
 int N=ZT.Pos("\r");
 AnsiString Z=ZT.Delete(N,1);
 ZT=Z;
}

while (ZT.Pos("\n"))
{
 int N=ZT.Pos("\n");
 AnsiString Z=ZT.Delete(N,1);
 ZT=Z;
}
return ZT;
}
//----------------------------------------
