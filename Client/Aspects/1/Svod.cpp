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

MP<TIniFile> Ini(Path+"NetAspects.ini");
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
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
  Application->HelpJump("IDH_СВОДНЫЙ_ОТЧЕТ");
}
}
//---------------------------------------------------------------------------
void TFSvod::ContSvodReport()
{

}
