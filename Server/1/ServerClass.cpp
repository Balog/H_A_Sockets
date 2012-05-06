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
   this->Socket->SendText("Command:5;1|"+IntToStr(Text.Length())+"#"+Text+"|");
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

