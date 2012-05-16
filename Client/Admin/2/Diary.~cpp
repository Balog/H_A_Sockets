//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Diary.h"
//#include "UServer.h"
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
//Sleep(1000);
FDiary->Initialize();

FDiary->Refresh();
}
//---------------------------------------------------------------------------
void __fastcall TFDiary::EnNDateClick(TObject *Sender)
{
PB->Visible=true;
PB->Position=0;
NDate->Enabled=EnNDate->Checked;
LoadDiary();
//Refresh();
}
//---------------------------------------------------------------------------
void __fastcall TFDiary::EnKDateClick(TObject *Sender)
{
PB->Visible=true;
PB->Position=0;
KDate->Enabled=EnKDate->Checked;
LoadDiary();
//Refresh();
}
//---------------------------------------------------------------------------
void  TFDiary::Initialize()
{
//ADODiary->Connected=false;
//ADODiary->Connected=true;



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
//if(Login->FieldByName("Login")->AsString!="")
//{
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
//}
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
/*
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
}
Filtr3=" "+Filtr3+" ";
*/
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
/*
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
}
Filtr4=" "+Filtr4+" ";
*/

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



//Filtr=" ("+Filtr1+" OR "+Filtr2+") AND ("+Filtr3+") AND ("+Filtr4+") AND ("+Filtr5+") ";

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
LoadDiary();
//Refresh();
}
//---------------------------------------------------------------------------

void __fastcall TFDiary::NTimeChange(TObject *Sender)
{
PB->Visible=true;
PB->Position=0;
LoadDiary();
//Refresh();
}
//---------------------------------------------------------------------------


void __fastcall TFDiary::Button1Click(TObject *Sender)
{
PB->Max=6;
PB->Visible=true;
PB->Min=0;
PB->Position=0;
LoadDiary();
//Refresh();
}
//---------------------------------------------------------------------------
void TFDiary::LoadDiary()
{
Prog->PB->Position++;
PB->Position++;
Zast->LoadTypeOp->Execute();
/*
//Main->MClient->Start();

if(!Register)
{
 Main->MClient->RegForm(this);
 Register=true;
}

MP<TADOCommand>Comm(this);
Comm->Connection=ADODiary;
Comm->CommandText="Delete * From TempTypeOp";
Comm->Execute();

Comm->CommandText="Delete * From TempOperations";
Comm->Execute();

MP<TADODataSet>LType(this);
LType->Connection=ADODiary;
LType->CommandText="Select Num, NameType from TempTypeOp order by Num";
LType->Active=true;

Table* RType=Main->MClient->CreateTable(this,  Main->ServerName, Main->VDB[Main->GetIDDBName("Diary")].ServerDB);
RType->SetCommandText("Select Num, NameType from TypeOp order by Num");
RType->Active(true);

Main->MClient->LoadTable(RType, LType);

if(Main->MClient->VerifyTable(LType, RType)==0)
{
MP<TADODataSet>LOp(this);
LOp->Connection=ADODiary;
LOp->CommandText="Select Num, Type, NameOperation from TempOperations order by Num";
LOp->Active=true;

Table* ROp=Main->MClient->CreateTable(this,  Main->ServerName, Main->VDB[Main->GetIDDBName("Diary")].ServerDB);
ROp->SetCommandText("Select Num, Type, NameOperation from Operations order by Num");
ROp->Active(true);

Main->MClient->LoadTable(ROp, LOp);

if(Main->MClient->VerifyTable(LOp, ROp)==0)
{

MergeTypeOp();

MergeOperations();

Word Y;
Word M;
Word D;

int Year;
int Month;
int Day;
String Dat, Dat1;

MP<TADODataSet>Tab(this);
Tab->Connection=ADODiary;

Table* Remote;
Remote=Main->MClient->CreateTable(this,  Main->ServerName, Main->VDB[Main->GetIDDBName("Diary")].ServerDB);
DecodeDate(NDate->Date, Y, M, D);
Year=(int)Y;
Month=(int)M;
Day=(int)D;
Dat=IntToStr(Month)+"/"+IntToStr(Day)+"/"+IntToStr(Year);

DecodeDate(KDate->Date+1, Y, M, D);
Year=(int)Y;
Month=(int)M;
Day=(int)D;
Dat1=IntToStr(Month)+"/"+IntToStr(Day)+"/"+IntToStr(Year);

if(EnNDate->Checked & EnKDate->Checked)
{
//ShowMessage("Включены оба");
//Включены оба


Tab->CommandText="SELECT Events.Num,  Events.Date_Time, Events.Comp, Events.Login, Events.Operation, Events.Prim FROM Events WHERE (((Events.Date_Time)>=#"+Dat+"#) AND ((Events.Date_Time)<=#"+Dat1+"#));";

Remote->SetCommandText("SELECT Events.Num,  Events.Date_Time, Events.Comp, Events.Login, Events.Operation, Events.Prim FROM Events WHERE (((Events.Date_Time)>=#"+Dat+"#) AND ((Events.Date_Time)<=#"+Dat1+"#))");

}
else
{
if(EnNDate->Checked)
{
//ShowMessage("Включен начальный");
//включен только начальный
Tab->CommandText="SELECT Events.Num,  Events.Date_Time, Events.Comp, Events.Login, Events.Operation, Events.Prim FROM Events WHERE (((Events.Date_Time)>=#"+Dat+"#));";

Remote->SetCommandText("SELECT Events.Num,  Events.Date_Time, Events.Comp, Events.Login, Events.Operation, Events.Prim FROM Events WHERE (((Events.Date_Time)>=#"+Dat+"#));");

}
else
{
if(EnKDate->Checked)
{
//ShowMessage("Включен конечный");
//включен только конечный
Tab->CommandText="SELECT Events.Num,  Events.Date_Time, Events.Comp, Events.Login, Events.Operation, Events.Prim FROM Events WHERE (((Events.Date_Time)<=#"+Dat1+"#));";

Remote->SetCommandText("SELECT Events.Num,  Events.Date_Time, Events.Comp, Events.Login, Events.Operation, Events.Prim FROM Events WHERE (((Events.Date_Time)<=#"+Dat1+"#))");

}
else
{
//ShowMessage("Выключены оба");
//не включен ниодин
Tab->CommandText="SELECT Events.Num,  Events.Date_Time, Events.Comp, Events.Login, Events.Operation, Events.Prim FROM Events;";

Remote->SetCommandText("SELECT Events.Num,  Events.Date_Time, Events.Comp, Events.Login, Events.Operation, Events.Prim FROM Events ");

}
}
}

Tab->Active=true;
Remote->Active(true);
if(Tab->RecordCount!=Remote->RecordCount())
{
PB->Min=0;
PB->Position=0;
PB->Visible=true;
PB->Max=Remote->RecordCount();
PB->Position=PB->Max;
PB->Min=Tab->RecordCount;
PB->Position=PB->Min;
//ShowMessage(PB->Max);
// ShowMessage("Есть различия");
for(Remote->First();!Remote->eof();Remote->Next())
{
 int N=Remote->FieldByName("Num");

 if(!Tab->Locate("Num",N,SO))
 {
  Tab->Insert();
  Tab->FieldByName("Num")->Value=Remote->FieldByName("Num");
  Tab->FieldByName("Date_Time")->Value=Remote->FieldByName("Date_Time");
  Tab->FieldByName("Comp")->Value=Remote->FieldByName("Comp");
  Tab->FieldByName("Login")->Value=Remote->FieldByName("Login");
  Tab->FieldByName("Operation")->Value=Remote->FieldByName("Operation");
  Tab->FieldByName("Prim")->Value=Remote->FieldByName("Prim");
  Tab->Post();
 PB->Position++;
 //PB->Repaint();
 this->Repaint();
 }

}
PB->Visible=false;
if(Events->Active)
{
Events->Active=false;
Events->Active=true;
}
}
Main->MClient->DeleteTable(this, Remote);
}
else
{
 ShowMessage("Ошибка копирования операций");
}
Main->MClient->DeleteTable(this, ROp);
}
else
{
 ShowMessage("Ошибка копирования типов операций");
}
Main->MClient->DeleteTable(this, RType);


Main->MClient->Stop();
*/
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
/*
MP<TADODataSet>Op(this);
Op->Connection=ADODiary;
Op->CommandText="select * From Operations";
Op->Active=true;

MP<TADODataSet>TempOp(this);
TempOp->Connection=ADODiary;
TempOp->CommandText="select * From TempOperations";
TempOp->Active=true;

for(TempOp->First();!TempOp->Eof;TempOp->Next())
{
 int N=TempOp->FieldByName("Num")->Value;

 if(!Op->Locate("Num",N,SO))
 {
  Op->Insert();
  Op->FieldByName("Num")->Value=TempOp->FieldByName("Num")->Value;
  Op->FieldByName("Type")->Value=TempOp->FieldByName("Type")->Value;
  Op->FieldByName("NameOperation")->Value=TempOp->FieldByName("NameOperation")->Value;
  Op->Post();
 }
}

MP<TADOCommand>Comm(this);
Comm->Connection=ADODiary;
Comm->CommandText="DELETE Operations.* FROM Operations LEFT JOIN TempOperations ON Operations.[Num] = TempOperations.[Num] WHERE (((TempOperations.Num) Is Null));";
Comm->Execute();
*/
}
//-----------------------------------------------------------------------------
void __fastcall TFDiary::CompsClickCheck(TObject *Sender)
{
Refresh();
}
//---------------------------------------------------------------------------

void __fastcall TFDiary::Button2Click(TObject *Sender)
{
//LoadDiary();

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
Application->HelpFile=ExtractFilePath(Application->ExeName)+"ADMINARM.HLP";
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

/*
String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Admin.ini");
String DPatch=Ini->ReadString("Main","DiaryBase","");

String DiaryPath=Path+DPatch+".dtb";
*/
/*
for(int i=0;i<NumBase;i++)
{
 String NameSect="Base"+IntToStr(i+1);

 String Name=Ini->ReadString(NameSect,"Name","");
 if(Name=="Diary")
 {
  String Name=Ini->ReadString(NameSect,"Name","");

  int AbsPath=Ini->ReadInteger(NameSect, "AbsPath",0);

  if(AbsPath==0)
  {
   DiaryPath=Path+Ini->ReadString(NameSect,"Base","");
  }
  else
  {
   DiaryPath= Ini->ReadString(NameSect,"Base","");
  }

  break;
 }
}
*/

/*
ADODiary->Connected=false;
ADODiary->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+DiaryPath+";Persist Security Info=False";
ADODiary->LoginPrompt=false;
ADODiary->Connected=true;
 */
try
{





}
catch(...)
{

}
        
}
//---------------------------------------------------------------------------

