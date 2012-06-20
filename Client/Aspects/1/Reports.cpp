//---------------------------------------------------------------------------

#include <vcl.h>
#include "Reports.h"
#pragma hdrstop

#include "Reports.h"
#include "MasterPointer.h"
#include "Zastavka.h"

#include "MainForm.h"
#include "Rep1.h"
//---------------------------------------------------------------------------

#pragma package(smart_init)
void Reports::CreateReport1(int Podr, TDateTime Date1, TDateTime Date2, AnsiString Isp, String Filtr1, String LFiltr)
{
AnsiString G;
if (Filtr1=="")
{
G="select * from TempAspects  Order By [����� �������]";
}
else
{
G="select * from TempAspects Where "+Filtr1+" Order By [����� �������]";
}

MP<TADODataSet>Report(Zast);
Report->Connection=Connect;

MP<TADODataSet>TempAspects(Zast);
TempAspects->Connection=Connect;

Report->Active=false;


Report->CommandText=G;
Report->Active=true;

AnsiString T="�-001_1 ";
TempAspects->Active=false;
AnsiString NP="��� �������������";
if(Podr!=0)
{
TempAspects->CommandText="Select * From ������������� Where [ServerNum]="+IntToStr(Podr);
TempAspects->Active=true;
NP=TempAspects->FieldByName("�������� �������������")->Value;
}

T=T+" �������� "+IntToStr(Podr);
AnsiString PP1=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\�-001_1.xlt");
AnsiString PP2=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\"+T+".xlt");
CopyFileTo(PP1.c_str() ,PP2.c_str());

Variant App =Variant::CreateObject("Excel.Application");
App.OlePropertySet("Visible",false);
Variant Book;


Book=App.OlePropertyGet("Workbooks").OleFunction("Add", PP2.c_str());


Variant Sheet=App.OlePropertyGet("ActiveSheet");
Sheet.OlePropertySet("Name","�-001.1");

DeleteFile(PP2);

int Start=17;
AnsiString Text;
int Num;
T="�������� ������������� �������� c "+Date1.DateString()+" �� "+Date2.DateString();
/*
App.OlePropertyGet("Cells",9,1).OlePropertySet("Value",T.c_str());
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������������� Where [ServerNum]="+IntToStr(Podr);
TempAspects->Active=true;
NP=TempAspects->FieldByName("�������� �������������")->Value;
*/
T=" "+NP;
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

App=NULL;
Book=NULL;
Sheet=NULL;
}
//------------------------------------------------------------------------------------------------------------------
void Reports::CreateReport2(int Podr, TDateTime Date1, TDateTime Date2, AnsiString Isp, String Filtr2, String LFiltr)
{
AnsiString G;
if (Filtr2=="")
{
G="select * from TempAspects  Order By [����� �������]";
}
else
{
G="select * from TempAspects Where "+Filtr2+" Order By [����� �������]";
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


AnsiString NP="��� �������������";
if(Podr!=0)
{
TempAspects->CommandText="Select * From ������������� Where [ServerNum]="+IntToStr(Podr);
TempAspects->Active=true;
NP=TempAspects->FieldByName("�������� �������������")->Value;
}


T=T+" �������� "+IntToStr(Podr);

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

int Start=16;
AnsiString Text;
int Num;
T="������ �������� ������������� �������� � "+Date1.DateString()+" �� "+Date2.DateString();
App.OlePropertyGet("Cells",9,1).OlePropertySet("Value",T.c_str());
/*
TempAspects->Active=false;
TempAspects->CommandText="Select * From ������������� Where [ServerNum]="+IntToStr(Podr);
TempAspects->Active=true;
NP=TempAspects->FieldByName("�������� �������������")->Value;
*/
T=" "+NP;
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

String S="INSERT INTO TempAspects ( [����� �������], �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, G, O, R, S, T, L, N, Z, ����������, [���������� �����������], [������� �����������], ��������������, [������������� �����������], [������������ �����������], [���������� � ��������], [������������ ���������� � ��������], �����������, [���� ��������], [������ ��������], [����� ��������], ServerNum ) ";
S=S+" SELECT �������.[����� �������], �������.�������������, �������.��������, �������.[��� ����������], �������.������������, �������.�������������, �������.������, �������.�����������, �������.G, �������.O, �������.R, �������.S, �������.T, �������.L, �������.N, �������.Z, �������.����������, �������.[���������� �����������], �������.[������� �����������], �������.��������������, �������.[������������� �����������], �������.[������������ �����������], �������.[���������� � ��������], �������.[������������ ���������� � ��������], �������.�����������, �������.[���� ��������], �������.[������ ��������], �������.[����� ��������], �������.ServerNum ";
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
