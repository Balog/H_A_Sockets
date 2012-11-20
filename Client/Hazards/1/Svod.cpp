//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Svod.h"
#include "Zastavka.h"
#include "MainForm.h"
#include "About.h"
#include "MasterPointer.h"
#include "MDBConnector.h"

#include <comobj.hpp>
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFSvod *FSvod;
//---------------------------------------------------------------------------
__fastcall TFSvod::TFSvod(TComponent* Owner)
        : TForm(Owner)
{
Registered=false;
}
//---------------------------------------------------------------------------

void __fastcall TFSvod::FormActivate(TObject *Sender)
{
this->Width=1024;
this->Height=742;
}
//---------------------------------------------------------------------------

void __fastcall TFSvod::FormShow(TObject *Sender)
{
Initialize();
}
//---------------------------------------------------------------------------
void __fastcall TFSvod::N13Click(TObject *Sender)
{
AnsiString Patch;
TDateTime Date;
int CountNum;
if (OpenDialog1->Execute())
{
 Patch=OpenDialog1->FileName;

 Temp->Connected=false;

Temp->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source="+Patch+";Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False";
Temp->Connected=true;

try
{


TempTable->Active=false;
TempTable->CommandText="Select * From Date_Time";
TempTable->Connection=Temp;
TempTable->Active=true;

float D;
D=TempTable->FieldByName("Date_Time")->Value;
Date=(TDateTime)D;

TempTable->Active=false;
TempTable->CommandText="Select * From Data";
TempTable->Active=true;

CountNum=TempTable->RecordCount;

TSvod->Active=false;
TSvod->Active=true;
int N=TSvod->RecordCount;
TSvod->Insert();
TSvod->FieldByName("Number")->Value=N+1;
TSvod->FieldByName("Date_Time")->Value=Date;
TSvod->FieldByName("Patch")->Value=Patch;
TSvod->FieldByName("Records")->Value=CountNum;
TSvod->FieldByName("Komm")->Value="Не обработан";
TSvod->Post();

}
catch(...)
{
ShowMessage("Неверная структура базы данных!\nВероятно вы указали не тот файл");
}

}


}
//---------------------------------------------------------------------------
void __fastcall TFSvod::N14Click(TObject *Sender)
{

TSvod->Delete();
TSvod->Active=false;
TSvod->Active=true;
TSvod->First();
for(int i=0;i<TSvod->RecordCount;i++)
{
TSvod->Edit();
TSvod->FieldByName("Number")->Value=i+1;
TSvod->Post();
TSvod->Next();
}
}
//---------------------------------------------------------------------------
void __fastcall TFSvod::N15Click(TObject *Sender)
{
int Number=TSvod->FieldByName("Number")->Value;
int N=TSvod->RecNo;
TSvod->Edit();
TSvod->FieldByName("Number")->Value=0;
TSvod->Post();

TSvod->First();
for(int i=0;i<TSvod->RecordCount;i++)
{
if(TSvod->FieldByName("Number")->Value==Number-1)
{
TSvod->Edit();
TSvod->FieldByName("Number")->Value=Number;
TSvod->Post();
break;
}
TSvod->Next();
}

TSvod->First();
for(int i=0;i<TSvod->RecordCount;i++)
{
if(TSvod->FieldByName("Number")->Value==0)
{
TSvod->Edit();
TSvod->FieldByName("Number")->Value=Number-1;
TSvod->Post();
break;
}
TSvod->Next();
}

TSvod->Active=false;
TSvod->Active=true;
TSvod->First();
TSvod->MoveBy(N-2);
}
//---------------------------------------------------------------------------
void __fastcall TFSvod::N16Click(TObject *Sender)
{
int Number=TSvod->FieldByName("Number")->Value;
int N=TSvod->RecNo;
TSvod->Edit();
TSvod->FieldByName("Number")->Value=0;
TSvod->Post();

TSvod->First();
for(int i=0;i<TSvod->RecordCount;i++)
{
if(TSvod->FieldByName("Number")->Value==Number+1)
{
TSvod->Edit();
TSvod->FieldByName("Number")->Value=Number;
TSvod->Post();
break;
}
TSvod->Next();
}

TSvod->First();
for(int i=0;i<TSvod->RecordCount;i++)
{
if(TSvod->FieldByName("Number")->Value==0)
{
TSvod->Edit();
TSvod->FieldByName("Number")->Value=Number+1;
TSvod->Post();
break;
}
TSvod->Next();
}

TSvod->Active=false;
TSvod->Active=true;
TSvod->First();
TSvod->MoveBy(N);
}
//---------------------------------------------------------------------------
void __fastcall TFSvod::Button1Click(TObject *Sender)
{

Zast->MClient->BlockServer("CreateMainSvod");

}

//---------------------------------------------------------------------------
AnsiString TFSvod::Address(Variant Sheet,int x,int y)
{
return Sheet.OlePropertyGet("Cells",y,x).OlePropertyGet("Address");
}
//--------------------------------------------------
void TFSvod::Initialize()
{
MP<TADODataSet>MReestr(this);
MReestr->Connection=Zast->ADOAspect;
MReestr->CommandText="select * From MainSvodReestr";
MReestr->Active=true;

MP<TADODataSet>DateTime(this);
DateTime->CommandText="Select * From Date_Time";

MP<TADODataSet>Data(this);
Data->CommandText="Select * From Data";

String MPatch=MReestr->FieldByName("Patch")->AsString;

if(FileExists(MPatch))
{
MainPatch->Font->Color=clBlack;
MainPatch->Text=MPatch;
Button1->Enabled=true;

try
{
Temp->Connected=false;
Temp->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source="+MPatch+";Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False";
Temp->Connected=true;

DateTime->Connection=Temp;
DateTime->Active=true;
DateTime->Active=false;

Data->Connection=Temp;
Data->Active=true;
Data->Active=false;

Temp->Connected=false;

String Path=ExtractFilePath(Application->ExeName);
String MainDatabase3=MPatch.SubString(1,MPatch.Length()-ExtractFileExt(MPatch).Length());

MP<TIniFile> Ini(Path+"Hazards.ini");
int Days=Ini->ReadInteger("Main","StoreArchive",0);

Zast->ADOSvod=new MDBConnector(ExtractFilePath(MainDatabase3), ExtractFileName(MainDatabase3), this);

Zast->ADOSvod->ClearArchive(Days);

Zast->ADOSvod->PackDB();

if(Days!=-1)
{
Zast->ADOSvod->SetPatchBackUp("Archive");
Zast->ADOSvod->Backup("Archive");
}

Zast->ADOSvod->ConnectDB();

DateTime->Connection=Zast->ADOSvod;
DateTime->Active=true;


Data->Connection=Zast->ADOSvod;
Data->Active=true;

TSvod->Connection=Zast->ADOAspect;
TSvod->CommandText="Select * From SlaveReestr Order by Number";
TSvod->Active=true;

MP<TADODataSet>SCount(this);
SCount->CommandText="SELECT Count(Data.Подразделение) AS [Count] FROM Data;";


for(TSvod->First();!TSvod->Eof;TSvod->Next())
{
try
{
String SPatch=TSvod->FieldByName("Patch")->Value;

Temp->Connected=false;
Temp->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source="+SPatch+";Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False";
Temp->Connected=true;

SCount->Connection=Temp;
SCount->Active=true;

TSvod->Edit();
TSvod->FieldByName("Records")->Value=SCount->FieldByName("Count")->Value;
TSvod->FieldByName("Komm")->Value="Не обработан";
TSvod->Post();


}
catch(...)
{
TSvod->Edit();
TSvod->FieldByName("Records")->Value=0;
TSvod->FieldByName("Komm")->Value="Не найден!";
TSvod->Post();
}
}
}
catch(...)
{
MainPatch->Text="Задайте путь к файлу главного сводного реестра";
MainPatch->Font->Color=clRed;
Button1->Enabled=false;
}

}
else
{
MainPatch->Text="Задайте путь к файлу главного сводного реестра";
MainPatch->Font->Color=clRed;
Button1->Enabled=false;
}
}
//--------------------------------------------------------------------
void __fastcall TFSvod::N1Click(TObject *Sender)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"HELP_ASPECTS.HLP";
Application->HelpJump("СводныйРеестр");

}
//---------------------------------------------------------------------------

void __fastcall TFSvod::N8Click(TObject *Sender)
{
Zast->Close();
}
//---------------------------------------------------------------------------

void TFSvod::CreateMainSvod()
{
}
//---------------------------------------------------------------------------
void TFSvod::CreateRep(TADODataSet *TempTable, Variant App, Variant Book, Variant Sheet, int &Start, int &NN, int Number)
{
TempTable->First();
AnsiString Text;
/*
for(int j=0;j<TempTable->RecordCount;j++)
{
NN=j;
App.OlePropertyGet("Cells",Start+j,1).OlePropertySet("Value",Number);
Text=TempTable->FieldByName("Название подразделения")->Value;
App.OlePropertyGet("Cells",Start+j,2).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Наименование деятельности")->Value;
App.OlePropertyGet("Cells",Start+j,3).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Наименование аспекта")->Value;
App.OlePropertyGet("Cells",Start+j,4).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Наименование воздействия")->Value;
App.OlePropertyGet("Cells",Start+j,5).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Название ситуации")->Value;
App.OlePropertyGet("Cells",Start+j,6).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Z")->Value;
App.OlePropertyGet("Cells",Start+j,7).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Мониторинг и контроль")->Value;
App.OlePropertyGet("Cells",Start+j,8).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Мониторинг и контроль")->Value;
App.OlePropertyGet("Cells",Start+j,9).OlePropertySet("Value",Text.c_str());


Number++;
TempTable->Next();
}

App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+NN)).c_str()).OlePropertySet("VerticalAlignment",-4160);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+NN)).c_str()).OlePropertySet("WrapText",true);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+NN)).c_str()).OlePropertyGet("Borders",7).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+NN)).c_str()).OlePropertyGet("Borders",8).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+NN)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+NN)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+NN)).c_str()).OlePropertyGet("Borders",11).OlePropertySet("Weight",4);
if (Number>1)
{
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+NN)).c_str()).OlePropertyGet("Borders",12).OlePropertySet("Weight",4);
}
*/
for(int j=0;j<TempTable->RecordCount;j++)
{
NN=j;

Text=TempTable->FieldByName("Название подразделения")->AsString;
App.OlePropertyGet("Cells",Start+j,1).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Наименование территории")->AsString;
App.OlePropertyGet("Cells",Start+j,2).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Наименование деятельности")->AsString;
App.OlePropertyGet("Cells",Start+j,3).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Наименование аспекта")->AsString;
App.OlePropertyGet("Cells",Start+j,4).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Специальность")->AsString;
App.OlePropertyGet("Cells",Start+j,5).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Наименование воздействия")->AsString;
App.OlePropertyGet("Cells",Start+j,6).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Название ситуации")->AsString;
App.OlePropertyGet("Cells",Start+j,7).OlePropertySet("Value",Text.c_str());
double T=TempTable->FieldByName("Z")->AsFloat;
Text=FloatToStr(T);
App.OlePropertyGet("Cells",Start+j,8).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Наименование значимости")->AsString;
App.OlePropertyGet("Cells",Start+j,9).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Наименование приоритетности")->AsString;
App.OlePropertyGet("Cells",Start+j,10).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Выполняющиеся мероприятия")->AsString;
App.OlePropertyGet("Cells",Start+j,11).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Предлагаемые мероприятия")->AsString;
App.OlePropertyGet("Cells",Start+j,12).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Мониторинг и контроль")->AsString;
App.OlePropertyGet("Cells",Start+j,13).OlePropertySet("Value",Text.c_str());
Text=TempTable->FieldByName("Мониторинг и контроль")->AsString;
App.OlePropertyGet("Cells",Start+j,14).OlePropertySet("Value",Text.c_str());


Number++;
TempTable->Next();
}

App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,14,Start+NN)).c_str()).OlePropertySet("VerticalAlignment",-4160);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,14,Start+NN)).c_str()).OlePropertySet("WrapText",true);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,14,Start+NN)).c_str()).OlePropertyGet("Borders",7).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,14,Start+NN)).c_str()).OlePropertyGet("Borders",8).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,14,Start+NN)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,14,Start+NN)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,14,Start+NN)).c_str()).OlePropertyGet("Borders",11).OlePropertySet("Weight",4);
if (Number>1)
{
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,14,Start+NN)).c_str()).OlePropertyGet("Borders",12).OlePropertySet("Weight",4);
}
}
//----------------------------------------------------------------------------
void TFSvod::EndSvod(Variant App, Variant Sheet, int Start)
{
App.OlePropertyGet("Cells",Start+3,2).OlePropertySet("Value","Согласовано:");
App.OlePropertyGet("Range",(Address(Sheet,2,Start+5)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+6,2).OlePropertySet("Value","(должность)");
App.OlePropertyGet("Cells",Start+6,2).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,2,Start+6)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,2,Start+8)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+9,2).OlePropertySet("Value","(должность)");
App.OlePropertyGet("Cells",Start+9,2).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,2,Start+9)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,4,Start+5)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+6,4).OlePropertySet("Value","(подпись)");
App.OlePropertyGet("Cells",Start+6,4).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+6)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,4,Start+8)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+9,4).OlePropertySet("Value","(подпись)");
App.OlePropertyGet("Cells",Start+9,4).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+9)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,6,Start+5)+":"+Address(Sheet,7,Start+5)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+5)+":"+Address(Sheet,7,Start+5)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+6)+":"+Address(Sheet,7,Start+6)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Cells",Start+6,6).OlePropertySet("Value","(Ф.И.О.)");
App.OlePropertyGet("Cells",Start+6,6).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+6)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,6,Start+8)+":"+Address(Sheet,7,Start+8)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+8)+":"+Address(Sheet,7,Start+8)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+9)+":"+Address(Sheet,7,Start+9)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Cells",Start+9,6).OlePropertySet("Value","(Ф.И.О.)");
App.OlePropertyGet("Cells",Start+9,6).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+9)).c_str()).OlePropertySet("HorizontalAlignment",-4108);


App.OlePropertySet("Visible",true);
}
//----------------------------------------------------------------------------
void __fastcall TFSvod::Button2Click(TObject *Sender)
{
if(OpenDialog1->Execute())
{
MP<TADODataSet>MReestr(this);
MReestr->Connection=Zast->ADOAspect;
MReestr->CommandText="select * From MainSvodReestr";
MReestr->Active=true;

MReestr->Edit();
MReestr->FieldByName("Patch")->Value=OpenDialog1->FileName;
MReestr->Post();

Initialize();
}
}
//---------------------------------------------------------------------------
void TFSvod::PrepareMergeAspects()
{

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOAspect;
Temp->CommandText="Select * From TempAspects order by [Номер аспекта]";
Temp->Active=true;

MP<TADODataSet>Memo(this);
Memo->Connection=Zast->ADOAspect;


for(Temp->First();!Temp->Eof;Temp->Next())
{
 int TempNum=Temp->FieldByName("Номер аспекта")->Value;

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
Temp->FieldByName("Выполняющиеся мероприятия")->Assign(TT);
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
Temp->FieldByName("Предлагаемые мероприятия")->Assign(TT);
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
Temp->FieldByName("Мониторинг и контроль")->Assign(TT);
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
Temp->FieldByName("Предлагаемый мониторинг и контроль")->Assign(TT);
Temp->Post();
}

}
//------------------------------------------------------------------------

void __fastcall TFSvod::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"Hazards.HLP";
  Application->HelpJump("IDH_СВОДНЫЙ_ОТЧЕТ");
}
}
//---------------------------------------------------------------------------
void TFSvod::ContSvodReport()
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
/*
String CT="INSERT INTO TempSvodReestr ( [Название подразделения], [Наименование деятельности], [Наименование аспекта], [Наименование воздействия], [Название ситуации], Z, [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Значимость ) ";
CT=CT+" SELECT Подразделения.[Название подразделения], Деятельность.[Наименование деятельности], Аспект.[Наименование аспекта], Воздействия.[Наименование воздействия], Ситуации.[Название ситуации], TempAspects.Z, TempAspects.[Мониторинг и контроль], TempAspects.[Предлагаемый мониторинг и контроль], TempAspects.Значимость ";
CT=CT+" FROM (Подразделения INNER JOIN (Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN TempAspects ON Деятельность.[Номер деятельности] = TempAspects.Деятельность) ON Аспект.[Номер аспекта] = TempAspects.Аспект) ON Воздействия.[Номер воздействия] = TempAspects.Воздействие) ON Ситуации.[Номер ситуации] = TempAspects.Ситуация) ON Подразделения.ServerNum = TempAspects.Подразделение) INNER JOIN (Logins INNER JOIN ObslOtdel ON Logins.Num = ObslOtdel.Login) ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel ";
CT=CT+" WHERE (((Logins.Role)=3));";
*/
String CT="INSERT INTO TempSvodReestr ( [Название подразделения], [Наименование территории], [Наименование деятельности], [Наименование аспекта], Специальность, [Наименование воздействия], [Название ситуации], Z, [Наименование значимости], [Выполняющиеся мероприятия], [Предлагаемые мероприятия], [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Значимость, [Наименование приоритетности] ) ";
CT=CT+" SELECT Подразделения.[Название подразделения], Территория.[Наименование территории], Деятельность.[Наименование деятельности], Аспект.[Наименование аспекта], TempAspects.Специальность, Воздействия.[Наименование воздействия], Ситуации.[Название ситуации], TempAspects.Z, TempAspects.[Наименование значимости], TempAspects.[Выполняющиеся мероприятия], TempAspects.[Предлагаемые мероприятия], TempAspects.[Мониторинг и контроль], TempAspects.[Предлагаемый мониторинг и контроль], TempAspects.Значимость, ВВР.[Наименование приоритетности] ";
CT=CT+" FROM ВВР INNER JOIN (Logins INNER JOIN ((Ситуации INNER JOIN ((Аспект INNER JOIN (Деятельность INNER JOIN (Территория INNER JOIN (TempAspects INNER JOIN Подразделения ON TempAspects.Подразделение = Подразделения.[ServerNum]) ON Территория.[Номер территории] = TempAspects.[Вид территории]) ON Деятельность.[Номер деятельности] = TempAspects.Деятельность) ON Аспект.[Номер аспекта] = TempAspects.Аспект) INNER JOIN Воздействия ON TempAspects.Воздействие = Воздействия.[Номер воздействия]) ON Ситуации.[Номер ситуации] = TempAspects.Ситуация) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login) ON ВВР.[Номер приоритетности] = TempAspects.Приоритетность ";
CT=CT+" WHERE (((Logins.Role)=3))";

Comm->CommandText=CT;
Comm->Execute();

MP<TADODataSet>SlReestr(this);
SlReestr->Connection=Zast->ADOAspect;
SlReestr->CommandText="select * from SlaveReestr where Records>0 Order by Number";
SlReestr->Active=true;

MP<TADODataSet>DT(this);
DT->CommandText="Select * from Date_Time";

MP<TADODataSet>D(this);
D->CommandText="select * from Data";

 MP<TADODataSet>TempSvod(this);
 TempSvod->Connection=Zast->ADOAspect;
 TempSvod->CommandText="Select * From TempSvodReestr Order by [Название подразделения]";
 TempSvod->Active=true;


for(SlReestr->First();!SlReestr->Eof;SlReestr->Next())
{
 String Patch=SlReestr->FieldByName("Patch")->Value;

 Temp->Connected=false;
 Temp->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source="+Patch+";Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False";
 Temp->Connected=true;

 DT->Active=false;
 DT->Connection=Temp;
 DT->Active=true;

 D->Active=false;
 D->Connection=Temp;
 D->Active=true;

 String Res="Обработан";
 for(D->First();!D->Eof;D->Next())
 {
 try
 {
  TempSvod->Insert();
  TempSvod->FieldByName("Название подразделения")->Value=D->FieldByName("Подразделение")->Value;
  TempSvod->FieldByName("Наименование деятельности")->Value=D->FieldByName("Деятельность")->Value;
  TempSvod->FieldByName("Наименование аспекта")->Value=D->FieldByName("Аспект")->Value;
  TempSvod->FieldByName("Наименование воздействия")->Value=D->FieldByName("Воздействие")->Value;
  TempSvod->FieldByName("Название ситуации")->Value=D->FieldByName("Ситуация")->Value;
  TempSvod->FieldByName("Z")->Value=D->FieldByName("Z")->Value;
  TempSvod->FieldByName("Мониторинг и контроль")->Value=D->FieldByName("Требования к качественным")->Value;
  TempSvod->FieldByName("Предлагаемый мониторинг и контроль")->Value=D->FieldByName("Требования к средствам")->Value;
  TempSvod->FieldByName("Значимость")->Value=D->FieldByName("Значимость")->Value;

  TempSvod->Post();
  }
  catch(...)
  {
  Res="Не обработан";
  }
 }

 SlReestr->Edit();
 SlReestr->FieldByName("Komm")->Value=Res;
 SlReestr->Post();
}

//Реализовать запись данных в собственную таблицу MainSvodReestr

 Temp->Connected=false;
 Temp->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source="+MainPatch->Text+";Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False";
 Temp->Connected=true;

 TempTable->Active=false;
 TempTable->Connection=Temp;
 TempTable->CommandText="Select * from Date_Time";
 TempTable->Active=true;

 TempTable->Edit();
 TempTable->FieldByName("Date_Time")->Value=Now();
 TempTable->Post();

 TempTable->Active=false;
 TempTable->Connection=Temp;
 TempTable->CommandText="Select * from Data";
 TempTable->Active=true;

 TempSvod->Active=false;
 TempSvod->Active=true;

 Comm->Connection=Temp;
 Comm->CommandText="Delete * from Data";
 Comm->Execute();

 for(TempSvod->First();!TempSvod->Eof;TempSvod->Next())
 {
  TempTable->Insert();
  TempTable->FieldByName("Подразделение")->Value=TempSvod->FieldByName("Название подразделения")->Value;
  TempTable->FieldByName("Деятельность")->Value=TempSvod->FieldByName("Наименование деятельности")->Value;
  TempTable->FieldByName("Аспект")->Value=TempSvod->FieldByName("Наименование аспекта")->Value;
  TempTable->FieldByName("Воздействие")->Value=TempSvod->FieldByName("Наименование воздействия")->Value;
  TempTable->FieldByName("Ситуация")->Value=TempSvod->FieldByName("Название ситуации")->Value;
  TempTable->FieldByName("Z")->Value=TempSvod->FieldByName("Z")->Value;
  TempTable->FieldByName("Требования к качественным")->Value=TempSvod->FieldByName("Мониторинг и контроль")->Value;
  TempTable->FieldByName("Требования к средствам")->Value=TempSvod->FieldByName("Предлагаемый мониторинг и контроль")->Value;
  TempTable->FieldByName("Значимость")->Value=TempSvod->FieldByName("Значимость")->Value;

  TempTable->Post();
 }
////////////////////////////////////////////////////////////
int NumFiles=0;
TempTable->Active=false;
TempTable->CommandText="Select * From TempSvodReestr";
TempTable->Connection=Zast->ADOAspect;
TempTable->Active=true;

if(TempTable->RecordCount!=0)
{
Variant App =Variant::CreateObject("Excel.Application");
Variant App1 =Variant::CreateObject("Excel.Application");
try
{
AnsiString T="Ф-001.3 ";
AnsiString T1="Ф-001.3 ";

T=T+" Сводный реестр значимых";
T1=T1+" Сводный реестр незначимых";

AnsiString P1=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\Ф-001_3.xlt");
AnsiString P2=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\"+T+".xlt");
CopyFile(P1.c_str() ,P2.c_str() , false);

AnsiString P11=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\Ф-001_3a.xlt");
AnsiString P21=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\"+T1+".xlt");
CopyFile(P11.c_str() ,P21.c_str() , false);


//App.OlePropertySet("Visible",true);
Variant Book=App.OlePropertyGet("Workbooks").OleFunction("Add", P2.c_str());
Variant Sheet=App.OlePropertyGet("ActiveSheet");
Sheet.OlePropertySet("Name","Ф-001.2");

App1.OlePropertySet("Visible",false);
Variant Book1=App1.OlePropertyGet("Workbooks").OleFunction("Add", P21.c_str());
Variant Sheet1=App1.OlePropertyGet("ActiveSheet");
Sheet1.OlePropertySet("Name","Ф-001.2");
App.OlePropertySet("Visible",false);
App1.OlePropertySet("Visible",false);

DeleteFile(P2);
DeleteFile(P21);

int Start=14;
int Start1=14;


TSvod->First();
AnsiString Patch;
int Number=1;

try
{


TempTable->Active=false;
TempTable->Connection=Zast->ADOAspect;
TempTable->CommandText="Select * From TempSvodReestr Where Значимость=true;";
TempTable->Active=true;

int NN=0;

CreateRep(TempTable, App, Book, Sheet, Start, NN, Number);
Start=Start+NN+1;

int NN1=0;

TempTable->Active=false;
TempTable->Connection=Zast->ADOAspect;
TempTable->CommandText="Select * From TempSvodReestr Where Значимость=false;";
TempTable->Active=true;

CreateRep(TempTable, App1, Book1, Sheet1, Start1, NN1, Number);

Start1=Start1+NN1+1;
NumFiles++;
}
catch (EOleException &)
{
 //ShowMessage("Неверный путь или файл");
 TSvod->Edit();
 TSvod->FieldByName("Komm")->Value="Нет файла";
 TSvod->Post();


}


EndSvod(App, Sheet, Start);
EndSvod(App1, Sheet1, Start1);

//ShowMessage("окончание свода");
}
catch(...)
{
App.OlePropertySet("Visible",true);

}



TSvod->Active=false;
TSvod->Active=true;
}
else
{
 ShowMessage("Нет данных для отчета");
}
}
