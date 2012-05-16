//---------------------------------------------------------------------------


#pragma hdrstop

#include "ServerClass.h"
#include "MasterPointer.h"
#include "Main.h"
//---------------------------------------------------------------------------

#pragma package(smart_init)

Clients::Clients(TComponent* Owner)
{
pOwner=Owner;

}
//-------------------------------------------------------
Clients::~Clients()
{
for(unsigned int i=0; i<VBases.size();i++)
{
  delete VBases[i].Database;
//  delete VBases[i].Command;
}
VBases.clear();

for(unsigned int i=0;i<VClients.size();i++)
{
delete VClients[i];
}
VClients.clear();
}
//-------------------------------------------------------
void Clients::WriteDiaryEvent(String Comp, String Login, String Type, String Name, String Prim)
{
DiaryEvent->WriteEvent(Now(), Comp, Login, Type, Name,Prim);

}
//----------------------------------------------------
void Clients::WriteDiaryEvent(String Comp, String Login, String Type, String Name)
{
DiaryEvent->WriteEvent(Now(), Comp, Login, Type, Name,"");

}
//----------------------------------------------------
void Clients::ConnectDiary(String DiaryPatch)
{
this->DiaryBase=DiaryPatch;

DiaryEvent=new Diary(pOwner, DiaryPatch);
}
//----------------------------------------------------
//********************************************************
Client::Client(Clients* Cls)
{
 Parent=Cls;
 Login="Неизвестен";
// Role=-1;

}
//**************************************************
Client::~Client()
{
for(unsigned int i=0;i<VForm.size();i++)
{
delete VForm[i];
}
VForm.clear();
}
//***************************************************
void Client::CommandExec(int Comm, vector<String>Parameters)
{
if(LastCommand==Comm | LastCommand==0)
{
 switch(Comm)
 {
  case 1:
  {



   //Ответ на команду запроса IP у клиента
   this->IP=Parameters[0];
   this->AppPatch=Parameters[1];
   //Передача сигнала готовности к дальнейшей работе
   //Ответ на него не предусмотрен
   LastCommand=0;
   this->Socket->SendText("Command:2;0|");

  Form1->ListBox1->Clear();
for(unsigned int i=0; i<Parent->VClients.size();i++)
{
String App;
if(ExtractFileName(Parent->VClients[i]->AppPatch)=="AdminARM.exe")
{
 App="AdminARM";
}
if(ExtractFileName(Parent->VClients[i]->AppPatch)=="NetAspects.exe")
{
 App="Aspects";
}
if(ExtractFileName(Parent->VClients[i]->AppPatch)=="Hazards.exe")
{
 App="Hazards";
}
 Form1->ListBox1->Items->Add(Parent->VClients[i]->IP+" "+App);
}
  break;
  }
  case 3:
  {
  //Организовать регистрацию формы на сервере
//   ShowMessage(Parameters[0]);
   LastCommand=0;
   mForm *F=new mForm();
   F->IDF=VForm.size();
   F->NameForm=Parameters[0];
   VForm.push_back(F);
   //Пока передаем один и тот же параметр а должен передаваться номер формы на сервере
   this->Socket->SendText("Command:3;1|"+IntToStr(IntToStr(F->IDF).Length())+"#"+IntToStr(F->IDF)+"|");
   break;
  }
  case 4:
  {
  LastCommand=0;
  for(unsigned int i=0;i<Parent->VBases.size();i++)
  {
   if(Parent->VBases[i].Name==Parameters[0])
   {
    if(Parameters[2]=="true")
    {
    Parent->VBases[i].Database->Connected=true;
    }
    else
    {
     Parent->VBases[i].Database->Connected=false;
    }
   //удачное открытие базы
    if(Parameters[1]!="-1")
    {

   this->Socket->SendText("Command:4;3|1#1|"+IntToStr(IntToStr(Parent->VBases[i].LicCount).Length())+"#"+IntToStr(Parent->VBases[i].LicCount)+"|"+IntToStr(Parent->VBases[i].Name.Length())+"#"+Parent->VBases[i].Name+"|");
    }
    else
    {
   this->Socket->SendText("Command:4;3|1#1|2#-1|1#" "");
    }


   break;
      }
  }
  break;
  }
  case 5:
   {
   //Прием команды чтения таблицы

   String Text=TableToStr(Parameters[0], Parameters[1]);
   //ShowMessage(Text);
   Text="Command:5;1|"+IntToStr(Text.Length())+"#"+Text+"|";

   this->Socket->SendText(Text);
   //Text=Text.SubString(Rec+1,Text.Length());

   break;
   }
   case 6:
   {

   if(Parameters[2]=="1")
   {
   Parent->WriteDiaryEvent(this->IP, Parameters[0], "Служебное","Пользователь идентифицирован");
   this->Login=Parameters[0];
   this->Socket->SendText("Command:6;0|");

   }
   else
   {
   //Parent->WriteDiaryEvent(this->IP, Parameters[0], "Служебное","Идентификация пользователя");
   Parent->WriteDiaryEvent(this->IP, Parameters[0], "Служебное","Пользователь не идентифицирован","Pass: "+Parameters[1]);
   }
   break;
   }
   case 7:
   {
   /*
   ShowMessage(Parameters[0]);
   ShowMessage(Parameters[1]);
   ShowMessage(Parameters[2]);
   ShowMessage(Parameters[3]);
   ShowMessage(Parameters[4]);
   */
//void WriteDiaryEvent(String Comp, String Login, String Type, String Name, String Prim);

Parent->WriteDiaryEvent(Parameters[0], Parameters[1], Parameters[2], Parameters[3], Parameters[4]);
   //this->Socket->SendText("Command:7;0|");
   break;
   }
   case 8:
   {
   /*
   ShowMessage(Parameters[0]);  //NameDB
   ShowMessage(Parameters[1]);  //SQL
   ShowMessage(Parameters[2]);  //Data
   */
   DecodeTable(Parameters[0],Parameters[1], Parameters[2]);

   this->Socket->SendText("Command:8;0|");
   break;
   }
   case 9:
   {
   //Команда для объединения данных от админа
   MergeLogins(Parameters[0]);
   this->Socket->SendText("Command:9;0|");
   break;
   }
 }

}
}
//---------------------------------------------------------------------------
TADOConnection* Client::GetDatabase(String NameDB)
{

 for(unsigned int i=0;i<Parent->VBases.size();i++)
 {
  if(Parent->VBases[i].Name==NameDB)
  {
   return Parent->VBases[i].Database;
  }
 }
 return 0;//Ошибка, не найдена база данных
}
//---------------------------------------------------------------------------
String Client::TableToStr(String NameDB, String SQLText)
{
   MP<TADODataSet>Tab((TForm*)Parent->pOwner);
   Tab->Connection=GetDatabase(NameDB);
   Tab->CommandText=SQLText;
   Tab->Active=true;

   //Формат передачи таблицы
   //Начало таблицы 27 1
   //int Esc=27;
   int S1=1;
   int S2=2;
   int S3=3;
   int S4=4;
   int S5=5;
   int S6=6;
   int S7=7;
   
   char C=VK_ESCAPE;
   char C1=((char)S1);
   char C2=((char)S2);
   char C3=((char)S3);
   char C4=((char)S4);
   char C5=((char)S5);
   char C6=((char)S6);
   char C7=((char)S7);

   String BeginTable=(String)C+(String)C1;
   //Конец таблицы 27 2
   String EndTable=(String)C+(String)C2;
   //Начало записи 27 3
   String BeginRecord=(String)C+(String)C3;
   //Конец записи 27 4
   String EndRecord=(String)C+(String)C4;
   //Начало поля 27 5
   String BeginField=(String)C+(String)C5;
   //Конец поля 27 6
   String EndField=(String)C+(String)C6;
   //Поле состоит из типа поля и значения разделенного символами 27 6
   String FieldSeparator=(String)C+(String)C7;
   //Типы полей
   //ftString - 1
   //ftInteger - 2
   //ftBoolean - 3
   //ftFloat - 4
   //ftDateTime - 5
   //ftMemo - 6
   //Поле Memo пока представляется как строка а там посмотрим
   //Все поля имеют текстовое предстваление, Boolean true - 1, false - 0
   String Ret=BeginTable;
   for(Tab->First();!Tab->Eof;Tab->Next())
   {
   Ret=Ret+BeginRecord;
    for(int i=0; i<Tab->FieldCount;i++)
    {
    Ret=Ret+BeginField;
     switch (Tab->FieldList->Fields[i]->DataType)
     {
      case ftWideString:
      {
      Ret=Ret+(String)C1+Tab->FieldList->Fields[i]->AsString;
      break;
      }
      case ftInteger:
      {
      Ret=Ret+(String)C2+FieldSeparator+IntToStr(Tab->FieldList->Fields[i]->AsInteger);
      break;
      }
      case ftAutoInc:
      {
      Ret=Ret+(String)C2+FieldSeparator+IntToStr(Tab->FieldList->Fields[i]->AsInteger);
      break;
      }
      case ftBoolean:
      {
      if(Tab->FieldList->Fields[i]->AsBoolean)
      {
      Ret=Ret+(String)C3+FieldSeparator+((char)1);
      }
      else
      {
      Ret=Ret+(String)C3+FieldSeparator+((char)0);
      }
      break;
      }
      case ftFloat:
      {
      Ret=Ret+(String)C4+FieldSeparator+FloatToStr(Tab->FieldList->Fields[i]->AsFloat);
      break;
      }
      case ftDateTime:
      {
      Ret=Ret+(String)C5+FieldSeparator+Tab->FieldList->Fields[i]->AsDateTime.DateTimeString();
      break;
      }
      case ftMemo:
      {
      Ret=Ret+(String)C6+FieldSeparator+Tab->FieldList->Fields[i]->AsString;
      break;
      }
     }
     Ret=Ret+EndField;
    }
    Ret=Ret+EndRecord;
   }
   Ret=Ret+EndTable;
   return Ret;
}
//--------------------------------------------------------------------------
void Client::DecodeTable(String NameDB, String ServerSQL, String Text)
{
TADOConnection* DB;

DB=GetDatabase(NameDB);

int FromPos=ServerSQL.LowerCase().Pos("from");
String DelText="Delete * "+ServerSQL.SubString(FromPos, ServerSQL.Length());

 MP<TADOCommand>Comm((TForm*)Parent->pOwner);
 Comm->Connection=DB;
 Comm->CommandText=DelText;
 Comm->Execute();

 MP<TADODataSet>Tab((TForm*)Parent->pOwner);
 Tab->Connection=DB;
 Tab->CommandText=ServerSQL;
 Tab->Active=true;

    //Формат передачи таблицы
   //Начало таблицы 27 1
   //int Esc=27;
   int S1=1;
   int S2=2;
   int S3=3;
   int S4=4;
   int S5=5;
   int S6=6;
   int S7=7;

   char C=VK_ESCAPE;
   char C1=((char)S1);
   char C2=((char)S2);
   char C3=((char)S3);
   char C4=((char)S4);
   char C5=((char)S5);
   char C6=((char)S6);
   char C7=((char)S7);

   String BeginTable=(String)C+(String)C1;
   //Конец таблицы 27 2
   String EndTable=(String)C+(String)C2;
   //Начало записи 27 3
   String BeginRecord=(String)C+(String)C3;
   //Конец записи 27 4
   String EndRecord=(String)C+(String)C4;
   //Начало поля 27 5
   String BeginField=(String)C+(String)C5;
   //Конец поля 27 6
   String EndField=(String)C+(String)C6;
   //Поле состоит из типа поля и значения разделенного символами 27 6
   String FieldSeparator=(String)C+(String)C7;
   //Типы полей
   //ftString - 1
   //ftInteger - 2
   //ftBoolean - 3
   //ftFloat - 4
   //ftDateTime - 5
   //ftMemo - 6
   //Поле Memo пока представляется как строка а там посмотрим
   //Все поля имеют текстовое предстваление, Boolean true - 1, false - 0

   if(Text.SubString(1,2)==BeginTable | Text.SubString(Text.Length()-2,2)==EndTable)
   {
    Text=Text.SubString(3,Text.Length()-4);
    do
    {
    if(Text.SubString(1,2)==BeginRecord)
    {
     int EndRecordPos=Text.Pos(EndRecord);
     String Record=Text.SubString(3,EndRecordPos-3);

     Tab->Insert();
     int i=0;
     do
     {
      if(Record.SubString(1,2)==BeginField)
      {
       int EndFieldPos=Record.Pos(EndField);
       String Field=Record.SubString(3, EndFieldPos-3);

       char T=Field[1];
       int Type=T;

       switch(Type)
       {
        case 1:
        {
        Tab->FieldList->Fields[i]->Value=Field.SubString(2, Field.Length());
        break;
        }
        case 2:
        {
        String F=Field.SubString(4, Field.Length());
        Tab->FieldList->Fields[i]->Value=StrToInt(F);
        break;
        }
        case 3:
        {

        break;
        }
        case 4:
        {

        }
        case 5:
        {

        break;
        }
        case 6:
        {

        break;
        }

       }
       i++;

       Record=Record.SubString(EndFieldPos+2, Record.Length());
      }
      else
      {
       ShowMessage("Нет знака начала поля");
      }
     }
     while(Record.Length()!=0);
       Tab->Post();
     Text=Text.SubString(EndRecordPos+2, Text.Length());
    }
    else
    {
//     ShowMessage("Нет знака начала записи");
    }
    }
    while(Text.Length()!=0);
   }
   else
   {
    ShowMessage("Нет знаков начала и конца таблицы "+ServerSQL+" базы "+NameDB);
   }
}
//--------------------------------------------------------------------------
void Client::MergeLogins(String NameDB)
{
//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Администратор", "Объединение логинов", "IDC="+IntToStr(IDC())+"IDDB="+IntToStr(IDDB));
TLocateOptions SO;
//try
//{
/*
TADOConnection *Database=NULL;
for(unsigned int i=0;i<VDatabase.size();i++)
{
 if(VDatabase[i]->IDDB==IDDB)
 {
  Database=VDatabase[i]->Database;
  break;
 }
}
*/
//if(Database!=NULL)
//{

MP<TADOCommand>Comm((TForm*)Parent->pOwner);
Comm->Connection=GetDatabase(NameDB);;
Comm->CommandText="UPDATE Logins SET Logins.Del = False;";
Comm->Execute();

MP<TADODataSet>TempLogins((TForm*)Parent->pOwner);
TempLogins->Connection=GetDatabase(NameDB);;
TempLogins->CommandText="Select * From TempLogins";
TempLogins->Active=true;

MP<TADODataSet>Logins((TForm*)Parent->pOwner);
Logins->Connection=GetDatabase(NameDB);;
Logins->CommandText="Select * From Logins Order by Num";
Logins->Active=true;

MP<TADODataSet>TempObslOtdel((TForm*)Parent->pOwner);
TempObslOtdel->Connection=GetDatabase(NameDB);;
TempObslOtdel->CommandText="Select * from TempObslOtdel";
TempObslOtdel->Active=true;

MP<TADODataSet>ObslOtdel((TForm*)Parent->pOwner);
ObslOtdel->Connection=GetDatabase(NameDB);;
ObslOtdel->CommandText="Select * from ObslOtdel";
ObslOtdel->Active=true;

MP<TADODataSet>TempPodr((TForm*)Parent->pOwner);
TempPodr->Connection=GetDatabase(NameDB);;
TempPodr->CommandText="Select * from TempПодразделения";
TempPodr->Active=true;

for(TempLogins->First();!TempLogins->Eof;TempLogins->Next())
{
 String Log=TempLogins->FieldByName("Login")->Value;
 if(Logins->Locate("Login",Log,SO))
 {
  int Role=Logins->FieldByName("Role")->Value;
  int TempRole=TempLogins->FieldByName("Role")->Value;
  if(Role==4 & TempRole!=Role)
  {
   Logins->Delete();
  }
 }
}

for( Logins->First();!Logins->Eof;Logins->Next())
{
 int N=Logins->FieldByName("Num")->AsInteger;
 if(TempLogins->Locate("ServerNum",N,SO))
 {
Logins->Edit();
Logins->FieldByName("Login")->Value=TempLogins->FieldByName("Login")->Value;
Logins->FieldByName("Code1")->Value=TempLogins->FieldByName("Code1")->Value;
Logins->FieldByName("Code2")->Value=TempLogins->FieldByName("Code2")->Value;
Logins->FieldByName("Role")->Value=TempLogins->FieldByName("Role")->Value;
Logins->Post();

TempLogins->Edit();
TempLogins->FieldByName("Del")->Value=true;
TempLogins->Post();
 }
 else
 {
  Logins->Edit();
  Logins->FieldByName("Del")->Value=true;
  Logins->Post();
 }
}

Comm->CommandText="DELETE Logins.Del FROM Logins WHERE (((Logins.Del)=True));";
Comm->Execute();

TempLogins->Active=false;
TempLogins->CommandText="Select * From TempLogins Where Del=false";
TempLogins->Active=true;


for(TempLogins->First();!TempLogins->Eof;TempLogins->Next())
{
Logins->Insert();
Logins->FieldByName("Login")->Value=TempLogins->FieldByName("Login")->Value;
Logins->FieldByName("Code1")->Value=TempLogins->FieldByName("Code1")->Value;
Logins->FieldByName("Code2")->Value=TempLogins->FieldByName("Code2")->Value;
Logins->FieldByName("Role")->Value=TempLogins->FieldByName("Role")->Value;
Logins->Post();

Logins->Active=false;
Logins->Active=true;
Logins->Last();

TempLogins->Edit();
TempLogins->FieldByName("ServerNum")->Value=Logins->FieldByName("Num")->Value;
TempLogins->Post();
}


Comm->CommandText="Delete * from ObslOtdel";
Comm->Execute();

TempLogins->Active=false;
TempLogins->CommandText="Select * From TempLogins";
TempLogins->Active=true;
int Log2;
int Otd2;
for(TempObslOtdel->First();!TempObslOtdel->Eof;TempObslOtdel->Next())
{
 int Log1=TempObslOtdel->FieldByName("Login")->Value;
 if(TempLogins->Locate("Num", Log1, SO))
 {
  Log2=TempLogins->FieldByName("ServerNum")->Value;
 }
 else
 {
//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка объединения таблицы ObslOtdel по логинам", "IDC="+IntToStr(IDC())+" Log1="+IntToStr(Log1)+" Log2="+IntToStr(Log2));
 }

 int Otd1=TempObslOtdel->FieldByName("NumObslOtdel")->Value;
 if(TempPodr->Locate("Номер подразделения", Otd1, SO))
 {
  Otd2=TempPodr->FieldByName("ServerNum")->Value;

  TempObslOtdel->Edit();
  TempObslOtdel->FieldByName("Login")->Value=Log2;
  TempObslOtdel->FieldByName("NumObslOtdel")->Value=Otd2;
  TempObslOtdel->Post();
 }
 else
 {
  TempObslOtdel->Edit();
  TempObslOtdel->FieldByName("Del")->Value=True;
  TempObslOtdel->Post();

//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка объединения таблицы ObslOtdel по подразделениям", "IDC="+IntToStr(IDC())+" Log1="+IntToStr(Otd1)+" Log2="+IntToStr(Otd2));

 }
}

Comm->CommandText="DELETE * From TempObslOtdel where Del=true";
Comm->Execute();

Comm->CommandText="DELETE * From ObslOtdel";
Comm->Execute();

Comm->CommandText="DELETE * From TempПодразделения";
Comm->Execute();

Comm->CommandText="INSERT INTO ObslOtdel ( Login, NumObslOtdel ) SELECT TempObslOtdel.Login, TempObslOtdel.NumObslOtdel FROM TempObslOtdel;";
Comm->Execute();

//return true;
//}
//else
//{
//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка объединения логинов (нет базы данных)", "IDC="+IntToStr(IDC())+" IDDB="+IntToStr(IDDB));

//return false;
//}
//}
//catch(...)
//{
//Form1->Block=-1;
//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Служебная ошибка", "Ошибка объединения логинов", "IDC="+IntToStr(IDC())+" IDDB="+IntToStr(IDDB));

//return false;
//}
/*
TLocateOptions SO;
MP<TADOCommand>Comm((TForm*)Parent->pOwner);
Comm->Connection=GetDatabase(NameDB);
Comm->CommandText="Delete * from Logins where Del=true";


MP<TADODataSet>TempLogins((TForm*)Parent->pOwner);
TempLogins->Connection=GetDatabase(NameDB);
TempLogins->CommandText="Select * From TempLogins";
TempLogins->Active=true;

MP<TADODataSet>Logins((TForm*)Parent->pOwner);
Logins->Connection=GetDatabase(NameDB);
Logins->CommandText="Select * From Logins Order by Num";
Logins->Active=true;

MP<TADODataSet>TempObslOtdel((TForm*)Parent->pOwner);
TempObslOtdel->Connection=GetDatabase(NameDB);

//Пройти по Logins, переписать данные совпадающих логинов,удаляя строки из TempLogins
//пометить к удалению те, у которых нет совпадений в TempLogins и удалить их в конце

for(Logins->First();!Logins->Eof;Logins->Next())
{
Logins->Edit();
Logins->FieldByName("Del")->Value=false;
Logins->Post();

String Log=Logins->FieldByName("Login")->Value;
bool B=TempLogins->Locate("Login",Log,SO);
if(B)
{
//есть совпадение
int NumLogin=Logins->FieldByName("Num")->Value;
int NumTempLogin= TempLogins->FieldByName("Num")->Value;

TempObslOtdel->Active=false;
TempObslOtdel->CommandText="Select * From TempObslOtdel where Login="+IntToStr(NumTempLogin);
TempObslOtdel->Active=true;
for(TempObslOtdel->First();!TempObslOtdel->Eof;TempObslOtdel->Next())
{
TempObslOtdel->Edit();
TempObslOtdel->FieldByName("Login")->Value=NumLogin;
TempObslOtdel->Post();
}

Logins->Edit();
Logins->FieldByName("Code1")->Value=TempLogins->FieldByName("Code1")->Value;
Logins->FieldByName("Code2")->Value=TempLogins->FieldByName("Code2")->Value;
Logins->FieldByName("Role")->Value=TempLogins->FieldByName("Role")->Value;
Logins->Post();
TempLogins->Delete();
}
else
{
Logins->Edit();
Logins->FieldByName("Del")->Value=true;
Logins->Post();
}
}
Comm->Execute();
//пройти по TempLogins и перенести в Logins оставшиеся
for(TempLogins->First();!TempLogins->Eof;TempLogins->Next())
{
Logins->Insert();
Logins->FieldByName("Login")->Value=TempLogins->FieldByName("Login")->Value;
Logins->FieldByName("Code1")->Value=TempLogins->FieldByName("Code1")->Value;
Logins->FieldByName("Code2")->Value=TempLogins->FieldByName("Code2")->Value;
Logins->FieldByName("Role")->Value=TempLogins->FieldByName("Role")->Value;
Logins->Post();

Logins->Active=false;
Logins->Active=true;
Logins->Last();

int NumLogin=Logins->FieldByName("Num")->Value;
int NumTempLogin= TempLogins->FieldByName("Num")->Value;

TempObslOtdel->Active=false;
TempObslOtdel->CommandText="Select * From TempObslOtdel where Login="+IntToStr(NumTempLogin);
TempObslOtdel->Active=true;
for(TempObslOtdel->First();!TempObslOtdel->Eof;TempObslOtdel->Next())
{
TempObslOtdel->Edit();
TempObslOtdel->FieldByName("Login")->Value=NumLogin;
TempObslOtdel->Post();
}
}

//очистить TempLogins
Comm->CommandText="Delete * from TempLogins";
Comm->Execute();
*/
}
//--------------------------------------------------------------------------

//***************************************************************************
 mForm::mForm()
 {

 }
 //**************************************************************************
 mForm::~mForm()
 {

 }
 //**************************************************************************
 //.........................................................................
 Diary::Diary(TComponent* Owner, String PatchDiary)
{
this->Form=(TForm*)Owner;



DiaryBase=new TADOConnection(this->Form);

DiaryBase->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source="+PatchDiary+";Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False";
DiaryBase->LoginPrompt=false;
DiaryBase->Connected=true;
}
//................................................
Diary::~Diary()
{
delete DiaryBase;
}
//.................................................
void Diary::WriteEvent(TDateTime DT, String Comp, String Login, String Type, String Name, String Prim)
{
MP<TADODataSet>TypeOp(Form);
TypeOp->Connection=DiaryBase;
TypeOp->CommandText="Select * From TypeOp order by Num";
TypeOp->Active=true;

int NumType;
if(TypeOp->Locate("NameType",Type, SO))
{
//Найдено, прочитать
NumType=TypeOp->FieldByName("Num")->Value;
}
else
{
//Не найдено, записать
TypeOp->Insert();
TypeOp->FieldByName("NameType")->Value=Type;
TypeOp->Post();
TypeOp->Active=false;
TypeOp->Active=true;
TypeOp->Last();
NumType=TypeOp->FieldByName("Num")->Value;
}

MP<TADODataSet>Operation(Form);
Operation->Connection=DiaryBase;
Operation->CommandText="Select * From Operations order by Num";
Operation->Active=true;

/*
Variant locvalues[] = {EDep->Text, EFam->Text};
Table1->Locate("Dep;Fam", VarArrayOf(locvalues,1),      SearchOptions<<loPartialKey<<loCaseInsensitive);
*/
int NumOp;
Variant locvalues[] = {Name, NumType};
if(Operation->Locate("NameOperation;Type",VarArrayOf(locvalues,1),SO))
{
//найдено, читать
NumOp=Operation->FieldByName("Num")->Value;
}
else
{
//Ненайдено, записать
Operation->Insert();
Operation->FieldByName("NameOperation")->Value=Name;
Operation->FieldByName("Type")->Value=NumType;
Operation->Post();
Operation->Active=false;
Operation->Active=true;
Operation->Last();
NumOp=Operation->FieldByName("Num")->Value;
}

MP<TADODataSet>Event(Form);
Event->Connection=DiaryBase;
Event->CommandText="Select * From Events order by Num";
Event->Active=true;

Event->Insert();
Event->FieldByName("Date_Time")->Value=DT;
Event->FieldByName("Comp")->Value=Comp;
Event->FieldByName("Login")->Value=Login;
Event->FieldByName("Operation")->Value=NumOp;
Event->FieldByName("Prim")->Value=Prim;
Event->Post();
}
//................................................
void Diary::WriteEvent(TDateTime DT, String Comp, String Login, String Type, String Name)
{
WriteEvent(DT, Comp, Login, Type, Name, "");
}
//.................................................

