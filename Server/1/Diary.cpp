//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Diary.h"
#include "Main.h"
#include "MasterPointer.h"
#include "inifiles.hpp";
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFDiary *FDiary;
//---------------------------------------------------------------------------
__fastcall TFDiary::TFDiary(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TFDiary::FormShow(TObject *Sender)
{
NDate->Date=Now();
KDate->Date=Now();
EnNDate->Checked=true;
EnKDate->Checked=true;


String Path=ExtractFilePath(Application->ExeName);
MP<TIniFile>Ini(Path+"Server.ini");
int NumBase=Ini->ReadInteger("Main","NumDatabases",0);

String DiaryPath="";
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
ADODiary->Connected=false;
ADODiary->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+DiaryPath+";Persist Security Info=False";
ADODiary->LoginPrompt=false;
ADODiary->Connected=true;

Initialize();
Refresh();
}
//---------------------------------------------------------------------------
void __fastcall TFDiary::EnNDateClick(TObject *Sender)
{
NDate->Enabled=EnNDate->Checked;
Refresh();
}
//---------------------------------------------------------------------------
void __fastcall TFDiary::EnKDateClick(TObject *Sender)
{
KDate->Enabled=EnKDate->Checked;
Refresh();
}
//---------------------------------------------------------------------------
void  TFDiary::Initialize()
{
MP<TADODataSet>Comp(this);
Comp->Connection=ADODiary;
Comp->CommandText="SELECT Events.Comp FROM Events GROUP BY Events.Comp ORDER BY Events.Comp;";
Comp->Active=true;

Comps->Clear();
int i=0;
for(Comp->First();!Comp->Eof;Comp->Next())
{
 Comps->Items->Add(Comp->FieldByName("Comp")->AsString);
 if(Comp->FieldByName("Comp")->AsString=="�� ���������")
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
Login->Connection=ADODiary;
Login->CommandText="SELECT Events.Login FROM Events GROUP BY Events.Login ORDER BY Events.Login;";
Login->Active=true;

i=0;
Logins->Clear();
for(Login->First();!Login->Eof;Login->Next())
{
//if(Login->FieldByName("Login")->AsString!="")
//{
Logins->Items->Add(Login->FieldByName("Login")->AsString);
if(Login->FieldByName("Login")->AsString=="��� �� ���������")
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
Type->Connection=ADODiary;
Type->CommandText="SELECT TypeOp.NameType FROM TypeOp ORDER BY TypeOp.NameType;";
Type->Active=true;
i=0;
Types->Clear();
for(Type->First();!Type->Eof;Type->Next())
{
 Types->Items->Add(Type->FieldByName("NameType")->AsString);
 Types->Checked[i]=true;
 i++;
}
}
//-----------------------------------------------------------------------
void TFDiary::Refresh()
{


String Filtr;
String Filtr1;
String Filtr2;
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
Events->Connection=ADODiary;
Events->CommandText=CT;
Events->Active=true;
}
//------------------------------------------------------
void __fastcall TFDiary::NDateClick(TObject *Sender)
{
Refresh();
}
//---------------------------------------------------------------------------

void __fastcall TFDiary::NTimeChange(TObject *Sender)
{
Refresh();
}
//---------------------------------------------------------------------------


void __fastcall TFDiary::Button1Click(TObject *Sender)
{
Refresh();
}
//---------------------------------------------------------------------------
