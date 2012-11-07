//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Diary.h"
#include "MasterPointer.h"
#include "inifiles.hpp";
#include "Main.h"
#include "ClientClass.h"
#include "Zastavka.h"
#include "Progress.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFDiary *FDiary;
//---------------------------------------------------------------------------
__fastcall TFDiary::TFDiary(TComponent* Owner)
        : TForm(Owner)
{
Register=false;
}
//---------------------------------------------------------------------------
void __fastcall TFDiary::FormShow(TObject *Sender)
{

FDiary->Initialize();

FDiary->Refresh();
}
//---------------------------------------------------------------------------
void __fastcall TFDiary::EnNDateClick(TObject *Sender)
{
PB->Visible=true;
PB->Position=0;
PB->Max=6;
NDate->Enabled=EnNDate->Checked;
LoadDiary();
}
//---------------------------------------------------------------------------
void __fastcall TFDiary::EnKDateClick(TObject *Sender)
{
PB->Visible=true;
PB->Position=0;
PB->Max=6;
KDate->Enabled=EnKDate->Checked;
LoadDiary();
//Refresh();
}
//---------------------------------------------------------------------------
void  TFDiary::Initialize()
{
MP<TADODataSet>Comp(this);
Comp->Connection=Zast->MClient->Diary;
Comp->CommandText="SELECT Events.Comp FROM Events GROUP BY Events.Comp ORDER BY Events.Comp;";
Comp->Active=true;

Comps->Clear();
int i=0;
for(Comp->First();!Comp->Eof;Comp->Next())
{
 Comps->Items->Add(Comp->FieldByName("Comp")->AsString);
 if(Comp->FieldByName("Comp")->AsString=="Не определен")
 {
 Comps->Checked[i]=false;
 }
 else
 {
 Comps->Checked[i]=true;
 }

 i++;
}

MP<TADODataSet>Login(this);
Login->Connection=Zast->MClient->Diary;
Login->CommandText="SELECT Events.Login FROM Events GROUP BY Events.Login ORDER BY Events.Login;";
Login->Active=true;

i=0;
Logins->Clear();
for(Login->First();!Login->Eof;Login->Next())
{
Logins->Items->Add(Login->FieldByName("Login")->AsString);
if(Login->FieldByName("Login")->AsString=="Не известен")
{
Logins->Checked[i]=false;
}
else
{
Logins->Checked[i]=true;
}

i++;
}

MP<TADODataSet>Type(this);
Type->Connection=Zast->MClient->Diary;
Type->CommandText="SELECT TypeOp.NameType FROM TypeOp ORDER BY TypeOp.NameType;";
Type->Active=true;
i=0;
Types->Clear();
for(Type->First();!Type->Eof;Type->Next())
{
 Types->Items->Add(Type->FieldByName("NameType")->AsString);
 if(Type->FieldByName("NameType")->AsString=="Служебное")
 {
 Types->Checked[i]=false;
 }
 else
 {
 Types->Checked[i]=true;
 }
 i++;
}
}
//-----------------------------------------------------------------------
void TFDiary::Refresh()
{


String Filtr;
String Filtr1="";
String Filtr2="";
String Filtr3;
String Filtr4;
String Filtr5;



Word Y;
Word M;
Word D;

int Year;
int Month;
int Day;
String Dat;

if(EnNDate->Checked)
{
DecodeDate(NDate->Date, Y, M, D);
Year=(int)Y;
Month=(int)M;
Day=(int)D;
Dat=IntToStr(Month)+"/"+IntToStr(Day)+"/"+IntToStr(Year);
Filtr1=" Date_Time>#"+Dat+"# ";
}

if(EnKDate->Checked)
{
DecodeDate(KDate->Date+1, Y, M, D);
Year=(int)Y;
Month=(int)M;
Day=(int)D;
Dat=IntToStr(Month)+"/"+IntToStr(Day)+"/"+IntToStr(Year);
Filtr2=" Date_Time<=#"+Dat+"# ";
}
for(int i=0;i<Comps->Count;i++)
{
 if(Comps->Checked[i])
 {
  if(Filtr3=="")
  {
   Filtr3="Events.Comp='"+Comps->Items->Strings[i]+"'";
  }
  else
  {
   Filtr3=Filtr3+" OR Events.Comp='"+Comps->Items->Strings[i]+"'";
  }
 }
 else
 {
  if(Filtr3=="")
  {
   Filtr3="Events.Comp<>'"+Comps->Items->Strings[i]+"'";
  }
  else
  {
   Filtr3="("+Filtr3+") AND Events.Comp<>'"+Comps->Items->Strings[i]+"'";
  }
 }
}
Filtr3=" "+Filtr3+" ";

for(int i=0;i<Logins->Count;i++)
{
 if(Logins->Checked[i])
 {
  if(Filtr4=="")
  {
   Filtr4="Events.Login='"+Logins->Items->Strings[i]+"'";
  }
  else
  {
   Filtr4=Filtr4+" OR Events.Login='"+Logins->Items->Strings[i]+"'";
  }
 }
 else
 {
  if(Filtr4=="")
  {
   Filtr4="Events.Login<>'"+Logins->Items->Strings[i]+"'";
  }
  else
  {
   Filtr4="("+Filtr4+") AND Events.Login<>'"+Logins->Items->Strings[i]+"'";
  }
 }
}
Filtr4=" "+Filtr4+" ";

for(int i=0;i<Types->Count;i++)
{
 if(Types->Checked[i])
 {
  if(Filtr5=="")
  {
   Filtr5="TypeOp.NameType='"+Types->Items->Strings[i]+"'";
  }
  else
  {
   Filtr5=Filtr5+" OR TypeOp.NameType='"+Types->Items->Strings[i]+"'";
  }
 }
 else
 {
  if(Filtr5=="")
  {
   Filtr5="TypeOp.NameType<>'"+Types->Items->Strings[i]+"'";
  }
  else
  {
   Filtr5="("+Filtr5+") AND TypeOp.NameType<>'"+Types->Items->Strings[i]+"'";
  }
 }
}
Filtr5=" "+Filtr5+" ";

if(Filtr1!="")
{
if(Filtr2!="")
{
Filtr=Filtr1+" AND "+Filtr2;
}
else
{
Filtr=Filtr1;
}
}
else
{
if(Filtr2!="")
{
Filtr=Filtr2;
}
}

if(Filtr!="")
{
if(Filtr3!=" ")
{
Filtr=Filtr+" AND ("+Filtr3+")";
}
}
else
{
if(Filtr3!=" ")
{
Filtr="("+Filtr3+")";
}
}

if(Filtr!="")
{
if(Filtr4!=" ")
{
Filtr=Filtr+" AND ("+Filtr4+")";
}
}
else
{
if(Filtr4!=" ")
{
Filtr="("+Filtr4+")";
}
}

if(Filtr!="")
{
if(Filtr5!="  ")
{
Filtr=Filtr+" AND ("+Filtr5+")";
}
}
else
{
if(Filtr5!="  ")
{
Filtr="("+Filtr5+")";
}
}

String CT="SELECT Events.Num, Events.Date_Time, Events.Comp, Events.Login, TypeOp.NameType, Operations.NameOperation, Events.Prim FROM TypeOp INNER JOIN (Operations INNER JOIN Events ON Operations.Num = Events.Operation) ON TypeOp.Num = Operations.Type ";

if(Filtr!="")
{
CT=CT+" Where "+Filtr+" ORDER BY Events.Num DESC;";
}
else
{
CT=CT+" ORDER BY Events.Num DESC;";
}

Events->Active=false;
Events->Connection=Zast->MClient->Diary;
Events->CommandText=CT;
Events->Active=true;

PB->Visible=false;
}
//------------------------------------------------------
void __fastcall TFDiary::NDateClick(TObject *Sender)
{
PB->Visible=true;
PB->Position=0;
PB->Max=6;
LoadDiary();
}
//---------------------------------------------------------------------------

void __fastcall TFDiary::NTimeChange(TObject *Sender)
{
PB->Visible=true;
PB->Position=0;
LoadDiary();
}
//---------------------------------------------------------------------------


void __fastcall TFDiary::Button1Click(TObject *Sender)
{
PB->Max=6;
PB->Visible=true;
PB->Min=0;
PB->Position=0;
LoadDiary();
}
//---------------------------------------------------------------------------
void TFDiary::LoadDiary()
{
Prog->PB->Position++;
PB->Position++;
Zast->LoadTypeOp->Execute();
}
//----------------------------------------------------------------------------
void TFDiary::MergeTypeOp()
{
MP<TADODataSet>Type(this);
Type->Connection=Zast->MClient->Diary;
Type->CommandText="select * From TypeOp";
Type->Active=true;

MP<TADODataSet>TempType(this);
TempType->Connection=Zast->MClient->Diary;
TempType->CommandText="select * From TempTypeOp";
TempType->Active=true;

for(TempType->First();!TempType->Eof;TempType->Next())
{
 int N=TempType->FieldByName("Num")->Value;

 if(!Type->Locate("Num",N,SO))
 {
  Type->Insert();
  Type->FieldByName("Num")->Value=TempType->FieldByName("Num")->Value;
  Type->FieldByName("NameType")->Value=TempType->FieldByName("NameType")->Value;
  Type->Post();
 }
}

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->MClient->Diary;
Comm->CommandText="DELETE TypeOp.* FROM TypeOp LEFT JOIN TempTypeOp ON TypeOp.[Num] = TempTypeOp.[Num] WHERE (((TempTypeOp.Num) Is Null));";
Comm->Execute();
}

//----------------------------------------------------------------------------
void TFDiary::MergeOperations()
{

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->MClient->Diary;
Comm->CommandText="Delete * from Operations";
Comm->Execute();

Comm->CommandText="INSERT INTO Operations ( Num, Type, NameOperation ) SELECT TempOperations.Num, TempOperations.Type, TempOperations.NameOperation FROM TempOperations;";
Comm->Execute();

Comm->CommandText="Delete * from TempOperations";
Comm->Execute();
}
//-----------------------------------------------------------------------------
void __fastcall TFDiary::CompsClickCheck(TObject *Sender)
{
Refresh();
}
//---------------------------------------------------------------------------

void __fastcall TFDiary::Button2Click(TObject *Sender)
{
AnsiString P1=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\Diary.xlt");
Variant App;
Variant Book;
Variant Sheet;
try
{
App =Variant::CreateObject("Excel.Application");
App.OlePropertySet("Visible",false);
Book=App.OlePropertyGet("Workbooks").OleFunction("Add", P1.c_str());
Sheet=App.OlePropertyGet("ActiveSheet");
App.OlePropertySet("Visible",false);

MP<TADODataSet>Event(this);
Event->Connection=Zast->MClient->Diary;
Event->CommandText=Events->CommandText;
Event->Active=true;

Text="за период от "+NDate->Date.DateString()+" по "+KDate->Date.DateString();
App.OlePropertyGet("Cells",2,1).OlePropertySet("Value",Text.c_str());

PB->Visible=true;
PB->Min=0;
PB->Position=0;
PB->Max=Event->RecordCount;

int Start=5;
int i=0;
String Text;
for(Event->First();!Event->Eof;Event->Next())
{
Text=Event->FieldByName("Date_Time")->AsString;
App.OlePropertyGet("Cells",Start+i,1).OlePropertySet("Value",Text.c_str());

Text=Event->FieldByName("Comp")->AsString;
App.OlePropertyGet("Cells",Start+i,2).OlePropertySet("Value",Text.c_str());

Text=Event->FieldByName("Login")->AsString;
App.OlePropertyGet("Cells",Start+i,3).OlePropertySet("Value",Text.c_str());

Text=Event->FieldByName("NameType")->AsString;
App.OlePropertyGet("Cells",Start+i,4).OlePropertySet("Value",Text.c_str());

Text=Event->FieldByName("NameOperation")->AsString;
App.OlePropertyGet("Cells",Start+i,5).OlePropertySet("Value",Text.c_str());

Text=Event->FieldByName("Prim")->AsString;
App.OlePropertyGet("Cells",Start+i,6).OlePropertySet("Value",Text.c_str());
i++;
PB->Position=i;
this->Repaint();
}
PB->Visible=false;
i--;
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,6,Start+i)).c_str()).OlePropertyGet("Borders",7).OlePropertySet("Weight",-4138);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,6,Start+i)).c_str()).OlePropertyGet("Borders",8).OlePropertySet("Weight",-4138);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,6,Start+i)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",-4138);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,6,Start+i)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",-4138);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,6,Start+i)).c_str()).OlePropertyGet("Borders",11).OlePropertySet("Weight",-4138);
if (i>=1)
{
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,6,Start+i)).c_str()).OlePropertyGet("Borders",12).OlePropertySet("Weight",2);
}

App.OlePropertySet("Visible",true);

}
catch(...)
{
ShowMessage("Ошибка составления отчета");
}
App=NULL;
Book=NULL;
Sheet=NULL;
}
//------------------------------------------------------------------------------------------------------------------
AnsiString TFDiary::Address(Variant Sheet,int x,int y)
{
return Sheet.OlePropertyGet("Cells",y,x).OlePropertyGet("Address");
}
void __fastcall TFDiary::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
if(Form1->CBDatabase->Text=="Аспекты")
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"ADMINARM.HLP";
}
else
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"ADMINARM-H.HLP";
}
Application->HelpJump("IDH_ПРОСМОТР_ЖУРНАЛА_СОБЫТИЙ");
}
}
//---------------------------------------------------------------------------





void __fastcall TFDiary::FormCreate(TObject *Sender)
{
NDate->Date=Now();
KDate->Date=Now();
EnNDate->Checked=true;
EnKDate->Checked=true;

try
{





}
catch(...)
{

}
        
}
//---------------------------------------------------------------------------

