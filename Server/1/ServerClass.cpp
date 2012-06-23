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
Block=0;
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
void Clients::SearchDubl(String IP, String Login, String AppPatch, TCustomWinSocket *S)
{
bool Dubl=false;
int Di=0;
IVC=VClients.begin();
for(unsigned int i=0;i<VClients.size();i++)
{
if(VClients[i]->Login==Login & VClients[i]->Socket!=S)
{
//Найден дубликат
Dubl=true;
Di=i;
break;
}
IVC++;
}
if(Dubl)
{
 //найден дубдикат надо сделать проверку его работоспособности
VClients[Di]->LastCommand=0;
VClients[Di]->SDubl=true;

try
{
VClients[Di]->Socket->SendText("Command:0;0|");
}
catch(...)
{
//Cl->VClients[i]->Socket->SendText("Command:10;1|1#1");
delete VClients[Di];
VClients.erase(IVC);


}



 /*
delete VClients[i];
VClients.erase(IVC);

  Form1->ListBox1->Clear();
for(unsigned int i=0; i<VClients.size();i++)
{
String App;
if(ExtractFileName(VClients[i]->AppPatch)=="AdminARM.exe")
{
 App="AdminARM";
}
if(ExtractFileName(VClients[i]->AppPatch)=="NetAspects.exe")
{
 App="Aspects";
}
if(ExtractFileName(VClients[i]->AppPatch)=="Hazards.exe")
{
 App="Hazards";
}
 Form1->ListBox1->Items->Add(VClients[i]->IP+" "+App);
}
 */
}
else
{
S->SendText("Command:6;0|");
}
Form1->ListBox1->Clear();
for(unsigned int i=0; i<VClients.size();i++)
{
String App;
if(ExtractFileName(VClients[i]->AppPatch)=="AdminARM.exe")
{
 App="AdminARM";
}
if(ExtractFileName(VClients[i]->AppPatch)=="NetAspects.exe")
{
 App="Aspects";
}
if(ExtractFileName(VClients[i]->AppPatch)=="Hazards.exe")
{
 App="Hazards";
}
 Form1->ListBox1->Items->Add(VClients[i]->IP+" "+App+" "+VClients[i]->Login);
}
}
//-------------------------------------------------------------------------
bool Clients::IfBlock(TCustomWinSocket *Socket, String Flag)
{
bool Ret=false;
if(Block==0)
{
 //Сервер не заблокирован
 if(Flag=="1")
 {
 //Команда блокировать
 //Блокируем
 Block=Socket;
 Ret=true;
 }
 else
 {
 //Сервер не заблокирован
 //Но разблокировать его конечно можно хоть и бесполезно
  Ret=true;
 }
}
else
{
 //Сервер уже заблокирован
 //Проверим, заблокирован ли он правильным клиентом
 if(Block==Socket)
 {
  //Клиент правильный
  if(Flag=="0")
  {
   //Команда разблокировки
   Block=0;
   Ret=true;
  }
  else
  {
   //Повторная блокировка тем же клиентом
   //Можно хоть и бесполезно
   Ret=true;
  }
 }
 else
 {
 //Не тот клиент
 //Нельзя ни заблокировать ни разблокировать
  Ret=false;
 }
}
return Ret;
}
//-------------------------------------------------------------------------
//********************************************************
Client::Client(Clients* Cls)
{
 Parent=Cls;
 Login="Не известен";
// Role=-1;
SDubl=false;
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
SDubl=false;
if(LastCommand==Comm | LastCommand==0)
{
 switch(Comm)
 {
  case 0:
  {
    try
    {
   //если существующий логин ответил то надо подать команду на отключение нового клиента
   //this->Socket->SendText("Command:10;1|1#0");
   for(unsigned int i=0; i<Parent->VClients.size();i++)
   {
    if(Parent->VClients[i]->Login==this->Login & Parent->VClients[i]->Socket!=this->Socket)
    {
    Parent->WriteDiaryEvent(Parent->VClients[i]->IP, Login, "Сервер", "Передача команды прекращения работы", "Уже работает на IP: "+IP);
     Parent->VClients[i]->Socket->SendText("Command:10;1|1#0");
    }
   }
   }
   catch(...)
   {

   Form1->UnBlockClients();
   }
  break;
  }
  case 1:
  {

//WriteEvent();


   //Ответ на команду запроса IP у клиента
   this->IP=Parameters[0];
   this->AppPatch=Parameters[1];
  Parent->WriteDiaryEvent(IP, Login, "Сервер", "IP получен", "Путь: "+Parameters[1]);
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
 Form1->ListBox1->Items->Add(Parent->VClients[i]->IP+" "+App+" "+Parent->VClients[i]->Login);
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
   Parent->WriteDiaryEvent(IP, Login, "Служебное", "Регистрация формы", "Имя: "+F->NameForm+" Номер="+IntToStr(F->IDF));

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
   Parent->WriteDiaryEvent(IP, Login, "Сервер", "Открытие базы данных", "Имя: "+Parent->VBases[i].Name+" Число лицензий="+IntToStr(Parent->VBases[i].LicCount));

   this->Socket->SendText("Command:4;3|1#1|"+IntToStr(IntToStr(Parent->VBases[i].LicCount).Length())+"#"+IntToStr(Parent->VBases[i].LicCount)+"|"+IntToStr(Parent->VBases[i].Name.Length())+"#"+Parent->VBases[i].Name+"|");
    }
    else
    {
    Parent->WriteDiaryEvent(IP, Login, "Сервер", "Открытие базы данных", "Имя: "+Parent->VBases[i].Name+" Число лицензий - Неограничено");
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
   try
   {
   Parent->WriteDiaryEvent(IP, Login, "Служебное", "Передача таблицы", "Имя: "+Parameters[0]+" SQL: "+Parameters[1]);

   String Text=TableToStr(Parameters[0], Parameters[1]);
   //ShowMessage(Text);
   Text="Command:5;1|"+IntToStr(Text.Length())+"#"+Text+"|";

   this->Socket->SendText(Text);
   //Text=Text.SubString(Rec+1,Text.Length());
   }
   catch(...)
   {

   Form1->UnBlockClients();
   }
   break;
   }
   case 6:
   {

   if(Parameters[2]=="1")
   {

   Parent->WriteDiaryEvent(this->IP, Parameters[0], "Служебное","Пользователь идентифицирован");
   this->Login=Parameters[0];

   Parent->SearchDubl(IP, Login, AppPatch, this->Socket);



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
   try
   {
Parent->WriteDiaryEvent(IP, Login, "Служебное", "Прием таблицы", "Имя: "+Parameters[0]+" SQL: "+Parameters[1]);

   DecodeTable(Parameters[0],Parameters[1], Parameters[2]);

   this->Socket->SendText("Command:8;0|");
   }
   catch(...)
   {
    Form1->UnBlockClients();
   }
   break;
   }
   case 9:
   {
   //Команда для объединения данных от админа
Parent->WriteDiaryEvent(IP, Login, "AdminARM", "Запись логинов", "Имя: "+Parameters[0]+" SQL: "+Parameters[1]);

   MergeLogins(Parameters[0]);
   this->Socket->SendText("Command:9;0|");
   break;
   }
   case 11:
   {
   try
   {
   MergePodr(Parameters[0]);
   this->Socket->SendText("Command:11;0|");
   }
   catch(...)
   {
    Form1->UnBlockClients();
   }
   break;
   }
   case 12:
   {
   try
   {
   MergeCrit(Parameters[0], Parameters[1]);
   this->Socket->SendText("Command:12;0|");
   }
   catch(...)
   {
    Form1->UnBlockClients();
   }
   break;
   }
   case 13:
   {
   try
   {
   MergeSit(Parameters[0], Parameters[1]);
   this->Socket->SendText("Command:13;0|");
   }
   catch(...)
   {
    Form1->UnBlockClients();
   }
   break;
   }
   case 14:
   {
   try
   {
   MergeNodeBranch(Parameters[0], Parameters[1], Parameters[2], Parameters[3], Parameters[4], Parameters[5], Parameters[6], Parameters[7]);
   this->Socket->SendText("Command:14;0|");
   }
   catch(...)
   {
    Form1->UnBlockClients();
   }
   break;
   }
   case 15:
   {
   try
   {
   MergeNodeBranch(Parameters[0], Parameters[1], Parameters[2]);
   this->Socket->SendText("Command:15;0|");
   }
   catch(...)
   {
    Form1->UnBlockClients();
   }
   break;
   }
   case 16:
   {
   String Login=PodrToLogin(Parameters[0],StrToInt(Parameters[1]));
   this->Socket->SendText("Command:16;1|"+IntToStr(Login.Length())+"#"+Login+"|");
   break;
   }
   case 17:
   {
   try
   {
   MergeAspectsMainSpec();
   this->Socket->SendText("Command:17;0|");
   }
   catch(...)
   {
    Form1->UnBlockClients();
   }
   break;
   }
   case 18:
   {
   try
   {
   MergeAspectsUser(StrToInt(Parameters[0]));
   this->Socket->SendText("Command:18;0|");
   }
   catch(...)
   {
    Form1->UnBlockClients();
   }
   break;
   }
   case 19:
   {
   //:IfBlock(TCustomWinSocket *Socket, bool Flag)
   try
   {
   bool Res=Parent->IfBlock(this->Socket,"1");
   if(Res)
   {
    //Сервер удалось заблокировать
    Parent->WriteDiaryEvent(IP, Login, "Сервер", "Сервер заблокирован", "Путь: "+AppPatch);
    this->Socket->SendText("Command:19;1|1#1");
   }
   else
   {
    //Сервер не удалось заблокировать
    Parent->WriteDiaryEvent(IP, Login, "Сервер", "Сервер заблокировать не удалось", "Путь: "+AppPatch);
    this->Socket->SendText("Command:19;1|1#0");
   }
   }
   catch(...)
   {
    Form1->UnBlockClients();
   }
   break;
   }
   case 20:
   {
   try
   {
   bool Res=Parent->IfBlock(this->Socket,"0");
   if(Res)
   {
    //Сервер разблокирован
    Parent->WriteDiaryEvent(IP, Login, "Сервер", "Сервер разблокирован", "Путь: "+AppPatch);
    this->Socket->SendText("Command:20;1|1#1");
   }
   else
   {
    //Сервер разблокировать не удалось
    Parent->WriteDiaryEvent(IP, Login, "Сервер", "Сервер разблокировать не удалось", "Путь: "+AppPatch);
    this->Socket->SendText("Command:20;1|1#0");
   }
   }
   catch(...)
   {
    Form1->UnBlockClients();
   }
   break;
   }
 }

}
Form1->ImageList1->GetIcon(1, Application->Icon);
Form1->NID.hIcon=Application->Icon->Handle;
Shell_NotifyIcon(NIM_MODIFY, &Form1->NID);
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
FromPos=DelText.LowerCase().Pos("order");
if(FromPos>0)
{
DelText=DelText.SubString(0, FromPos-2);
}

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
        String F=Field.SubString(4, Field.Length());
        if(F=="\x01")
        {
        Tab->FieldList->Fields[i]->Value=true;
        }
        else
        {
        Tab->FieldList->Fields[i]->Value=false;
        }
        break;
        }
        case 4:
        {

        }
        case 5:
        {
        String F=Field.SubString(4, Field.Length());
        Tab->FieldList->Fields[i]->Value=StrToDateTime(F);
        break;
        }
        case 6:
        {
        String F=Field.SubString(4, Field.Length());
        Tab->FieldList->Fields[i]->Value=F;
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
void Client::MergePodr(String DB1)
{
TADOConnection *Database2=GetDatabase(DB1);

MP<TADOCommand>Comm2(Form1);
Comm2->Connection=Database2;
Comm2->CommandText="UPDATE Подразделения SET Подразделения.Del = False;";
Comm2->Execute();
Comm2->CommandText="UPDATE TempПодразделения SET TempПодразделения.Del = False;";
Comm2->Execute();

MP<TADODataSet>Temp(Form1);
Temp->Connection=Database2;
Temp->CommandText="Select * From TempПодразделения Order by [Номер подразделения]";
Temp->Active=true;

MP<TADODataSet>Podr(Form1);
Podr->Connection=Database2;
Podr->CommandText="Select * From Подразделения Order by [Номер подразделения]";
Podr->Active=true;

for(Podr->First();!Podr->Eof;Podr->Next())
{
int Num=Podr->FieldByName("Номер подразделения")->AsInteger;

if(Temp->Locate("ServerNum", Num, SO))
{
 //Найдено
 Podr->Edit();
 Podr->FieldByName("Название подразделения")->Value=Temp->FieldByName("Название подразделения")->Value;
 Podr->Post();

 Temp->Edit();
 Temp->FieldByName("Del")->Value=true;
 Temp->Post();

}
else
{
 //Ненайдено
 Podr->Edit();
 Podr->FieldByName("Del")->Value=true;
 Podr->Post();
}
}

Comm2->CommandText="Delete * From подразделения Where Del=true";
Comm2->Execute();

Temp->Active=false;
Temp->CommandText="Select * From TempПодразделения where Del=false Order by [Номер подразделения]";
Temp->Active=true;

for(Temp->First();!Temp->Eof;Temp->Next())
{
Podr->Insert();
Podr->FieldByName("Название подразделения")->Value=Temp->FieldByName("Название подразделения")->Value;
Podr->Post();
Podr->Active=false;
Podr->Active=true;
Podr->Last();
Temp->Edit();
Temp->FieldByName("ServerNum")->Value=Podr->FieldByName("Номер подразделения")->Value;
Temp->Post();
}


}
//--------------------------------------------------------------------------
void Client::MergeCrit(String DB1, String DB2)
{
TADOConnection *Database1=GetDatabase(DB1);


MP<TADOCommand>Comm1(Form1);
Comm1->Connection=Database1;
Comm1->CommandText="Delete * From Значимость";
Comm1->Execute();

Comm1->CommandText="INSERT INTO Значимость ( [Номер значимости], [Наименование значимости], Критерий1, Критерий, [Мин граница], [Макс граница], [Необходимая мера] ) SELECT TempZn.[Номер значимости], TempZn.[Наименование значимости], TempZn.Критерий1, TempZn.Критерий, TempZn.[Мин граница], TempZn.[Макс граница], TempZn.[Необходимая мера] FROM TempZn;";
Comm1->Execute();

//Aspects
TADOConnection *Database2=GetDatabase(DB2);

MP<TADOCommand>Comm2(Form1);
Comm2->Connection=Database2;
Comm2->CommandText="Delete * From Значимость";
Comm2->Execute();

MP<TADODataSet>Tab1(Form1);
Tab1->Connection=Database1;
Tab1->CommandText="Select * From Значимость order by [Номер значимости]";
Tab1->Active=true;

MP<TADODataSet>Tab2(Form1);
Tab2->Connection=Database2;
Tab2->CommandText="Select * From Значимость";
Tab2->Active=true;

for(Tab1->First();!Tab1->Eof;Tab1->Next())
{
 Tab2->Insert();
 Tab2->FieldByName("Номер значимости")->Value=Tab1->FieldByName("Номер значимости")->Value;
 Tab2->FieldByName("Наименование значимости")->Value=Tab1->FieldByName("Наименование значимости")->Value;
 Tab2->FieldByName("Критерий1")->Value=Tab1->FieldByName("Критерий1")->Value;
 Tab2->FieldByName("Критерий")->Value=Tab1->FieldByName("Критерий")->Value;
 Tab2->FieldByName("Мин граница")->Value=Tab1->FieldByName("Мин граница")->Value;
 Tab2->FieldByName("Макс граница")->Value=Tab1->FieldByName("Макс граница")->Value;
 Tab2->FieldByName("Необходимая мера")->Value=Tab1->FieldByName("Необходимая мера")->Value;
 Tab2->Post();
}

Comm1->CommandText="Delete * From TempZn";
Comm1->Execute();
}
//--------------------------------------------------------------------------
void Client::MergeSit(String DB1, String DB2)
{
TADOConnection *Database1=GetDatabase(DB1);

MP<TADOCommand>Comm1(Form1);
Comm1->Connection=Database1;
Comm1->CommandText="Delete * From Ситуации";
Comm1->Execute();

Comm1->CommandText="INSERT INTO Ситуации ( [Номер ситуации], [Название ситуации] ) SELECT TempSit.[Номер ситуации], TempSit.[Название ситуации] FROM TempSit;";
Comm1->Execute();

TADOConnection *Database2=GetDatabase(DB2);

MP<TADOCommand>Comm2(Form1);
Comm2->Connection=Database2;
Comm2->CommandText="UPDATE Ситуации SET Ситуации.Del = False;";
Comm2->Execute();

MP<TADODataSet>Temp(Form1);
Temp->Connection=Database2;
Temp->CommandText="Select * From TempSit Order by [Номер ситуации]";
Temp->Active=true;

MP<TADODataSet>Sit(Form1);
Sit->Connection=Database2;
Sit->CommandText="Select * From Ситуации Where [Номер ситуации]<>0 Order by [Номер ситуации]";
Sit->Active=true;

for(Sit->First();!Sit->Eof;Sit->Next())
{
int Num=Sit->FieldByName("номер ситуации")->Value;

if(Temp->Locate("Номер ситуации", Num, SO))
{
 //Найдено
 Sit->Edit();
 Sit->FieldByName("Номер ситуации")->Value=Temp->FieldByName("Номер ситуации")->Value;
 Sit->FieldByName("Название ситуации")->Value=Temp->FieldByName("Название ситуации")->Value;
 Sit->FieldByName("Показ")->Value=true;
 Sit->Post();

 Temp->Delete();
}
else
{
 //Ненайдено
 Sit->Edit();
 Sit->FieldByName("Del")->Value=true;
 Sit->Post();
}
}
MP<TADODataSet>Asp(Form1);
Asp->Connection=Database2;
Asp->CommandText="SELECT Аспекты.* FROM Ситуации INNER JOIN Аспекты ON Ситуации.[Номер ситуации] = Аспекты.Ситуация WHERE (((Ситуации.Del)=True));";
Asp->Active=true;

for(Asp->First();!Asp->Eof;Asp->Next())
{
 Asp->Edit();
 Asp->FieldByName("Ситуация")->Value=0;
 Asp->Post();
}

Comm2->CommandText="Delete * From Ситуации Where Del=true AND Показ=True";
Comm2->Execute();

Comm2->CommandText="INSERT INTO Ситуации ( [Номер ситуации], [Название ситуации], [показ] ) SELECT TempSit.[Номер ситуации], TempSit.[Название ситуации], true FROM TempSit Where [Номер ситуации]<>0;";
Comm2->Execute();

//Temp->CommandText="Select * From TempSit Order by [Номер ситуации]";
//Temp->Active=true;

Comm2->CommandText="Delete * From TempSit";
Comm2->Execute();
}
//-------------------------------------------------------------------------
void Client::MergeNodeBranch(String NameDatabase1, String NameNode, String NameBranch, String NameDatabase2, String NameTable, String NameField, String NameKey, String Name)
{
TADOConnection *Database1=GetDatabase(NameDatabase1);

if(Database1!=NULL)
{
MP<TADODataSet>TempNode(Form1);
TempNode->Connection=Database1;
TempNode->CommandText="Select * From TempNode";
TempNode->Active=true;

MP<TADODataSet>Node(Form1);
Node->Connection=Database1;
Node->CommandText="Select * From "+NameNode;
Node->Active=true;

MP<TADOCommand>Comm(Form1);
Comm->Connection=Database1;
Comm->CommandText="UPDATE "+NameNode+" SET "+NameNode+".Del = False, "+NameNode+".CopyNum = -1;";
Comm->Execute();

for(Node->First();!Node->Eof;Node->Next())
{
int NumNode=Node->FieldByName("Номер узла")->Value;

if(TempNode->Locate("Номер узла", NumNode, SO))
{
//Найдено
Node->Edit();
Node->FieldByName("Родитель")->Value=TempNode->FieldByName("Родитель")->Value;
Node->FieldByName("Название")->Value=TempNode->FieldByName("Название")->Value;
Node->Post();

TempNode->Delete();
}
else
{
//Ненайдено
Node->Edit();
Node->FieldByName("Del")->Value=true;
Node->Post();
}
}

Comm->CommandText="DELETE "+NameNode+".Del FROM "+NameNode+" WHERE ((("+NameNode+".Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO "+NameNode+" ( CopyNum, Родитель, Название ) SELECT TempNode.[Номер узла], TempNode.Родитель, TempNode.Название FROM TempNode;";
Comm->Execute();


MP<TADODataSet>TempBranch(Form1);
TempBranch->Connection=Database1;
TempBranch->CommandText="Select * From TempBranch";
TempBranch->Active=true;

MP<TADODataSet>Branch(Form1);
Branch->Connection=Database1;
Branch->CommandText="Select * From "+NameBranch;
Branch->Active=true;

Comm->CommandText="UPDATE "+NameBranch+" SET "+NameBranch+".Del = False, "+NameBranch+".NumCopy = -1; ";
Comm->Execute();
Node->Active=false;
Node->Active=true;

for(Branch->First();!Branch->Eof;Branch->Next())
{
 int NumBranch=Branch->FieldByName("Номер ветви")->Value;

 if(TempBranch->Locate("Номер ветви",NumBranch,SO))
 {
  Branch->Edit();
  int NumPar=TempBranch->FieldByName("Номер родителя")->Value;
  if(Node->Locate("CopyNum",NumPar,SO))
  {
  NumPar=Node->FieldByName("Номер узла")->Value;
  }

  Branch->FieldByName("Номер родителя")->Value=NumPar;
  Branch->FieldByName("Название")->Value=TempBranch->FieldByName("Название")->Value;
  Branch->FieldByName("Показ")->Value=TempBranch->FieldByName("Показ")->Value;
  Branch->Post();

  TempBranch->Delete();
 }
 else
 {
  Branch->Edit();
  Branch->FieldByName("Del")->Value=true;
  Branch->Post();
 }
}

Comm->CommandText="DELETE "+NameBranch+".Del FROM "+NameBranch+" WHERE ((("+NameBranch+".Del)=True));";
Comm->Execute();

for(TempBranch->First();!TempBranch->Eof;TempBranch->Next())
{
 int NumTemp=TempBranch->FieldByName("Номер родителя")->Value;

 if(Node->Locate("CopyNum", NumTemp, SO))
 {
  int NumPar=Node->FieldByName("Номер узла")->Value;

  Branch->Insert();
  Branch->FieldByName("Номер родителя")->Value=NumPar;
  Branch->FieldByName("Название")->Value=TempBranch->FieldByName("Название")->Value;
  Branch->FieldByName("Показ")->Value=TempBranch->FieldByName("Показ")->Value;
  Branch->Post();
 }
 else
 {
 if(Node->Locate("Номер узла", NumTemp, SO))
 {
  int NumPar=Node->FieldByName("Номер узла")->Value;

  Branch->Insert();
  Branch->FieldByName("Номер родителя")->Value=NumPar;
  Branch->FieldByName("Название")->Value=TempBranch->FieldByName("Название")->Value;
  Branch->FieldByName("Показ")->Value=TempBranch->FieldByName("Показ")->Value;
  Branch->Post();
 }
 else
 {
//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка записи структуры", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase1+" Номер узла: "+IntToStr(NumTemp)+"NameNode "+NameNode+" NameBravch "+NameBranch);
Parent->WriteDiaryEvent(IP, Login, "Ошибка обработки данных", "Ошибка записи структуры", " DB: "+NameDatabase1+" Номер узла: "+IntToStr(NumTemp)+"NameNode "+NameNode+" NameBravch "+NameBranch);
 }
 }
}
}
else
{
//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка записи узлов и ветвей 1 (Full) (нет базы данных)", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase1);
Parent->WriteDiaryEvent(IP, Login, "Ошибка обработки данных", "Ошибка записи узлов и ветвей 1 (Full) (нет базы данных)", " DB: "+NameDatabase1);

}




TADOConnection *Database2=GetDatabase(NameDatabase2);



if(Database2!=NULL)
{
//План таков:
//На всякий случай снимаем все пометки в целевой таблице
//проходим по таблице ветви в Reference и ищем совпадение номера в целевой таблице базы аспектов
//если находим - обновляем название и помечаем запись в целевой таблице
//если не находим - добавляем запись в целевую таблицу и помечаем
//после окончания прохода все непомеченные записи целевой таблицы уже отсутствуют в новой версии справочника.
//запросом на исправление в таблице аспектов заменяем индекс ссылки на непомеченые записи нулями.
//Потом удаляем непомеченые записи целевой и снимаем пометки целевой таблицы.

//Database1 (Comm) - Reference
//Database2 (Comm2) - аспекты

MP<TADOCommand>Comm2(Form1);
Comm2->Connection=Database2;
Comm2->CommandText="UPDATE "+NameTable+" SET "+NameTable+".Del = False;";
Comm2->Execute();

MP<TADODataSet>Branch(Form1);
Branch->Connection=Database1;
Branch->CommandText="Select * From "+NameBranch+" Where показ=true Order by [Номер ветви]";
Branch->Active=true;

MP<TADODataSet>CurrTable(Form1);
CurrTable->Connection=Database2;
CurrTable->CommandText="Select * from "+NameTable+" Where показ=true";
CurrTable->Active=true;


for(Branch->First();!Branch->Eof;Branch->Next())
{
 int NumBranch=Branch->FieldByName("Номер ветви")->Value;
 if(CurrTable->Locate(NameKey, NumBranch, SO))
 {
  //Найдено
  CurrTable->Edit();
  CurrTable->FieldByName(Name)->Value=Branch->FieldByName("Название")->Value;
  CurrTable->FieldByName("Показ")->Value=Branch->FieldByName("Показ")->Value;
  CurrTable->FieldByName("Del")->Value=true;
  CurrTable->Post();
 }
 else
 {
  //Ненайдено
  CurrTable->Insert();
  CurrTable->FieldByName(NameKey)->Value=Branch->FieldByName("Номер ветви")->Value;
  CurrTable->FieldByName(Name)->Value=Branch->FieldByName("Название")->Value;
  CurrTable->FieldByName("Показ")->Value=Branch->FieldByName("Показ")->Value;
  CurrTable->FieldByName("Del")->Value=true;
  CurrTable->Post();
 }
}

Comm2->CommandText="UPDATE ["+NameTable+"] INNER JOIN Аспекты ON ["+NameTable+"].["+NameKey+"] = Аспекты.["+NameField+"] SET Аспекты.["+NameField+"] = 0 WHERE (((["+NameTable+"].Показ)=True) AND ((["+NameTable+"].Del)=False)); ";
//Comm2->CommandText="UPDATE "+NameTable+" INNER JOIN Аспекты ON "+NameTable+".[Номер воздействия] = Аспекты."+NameField+" SET Аспекты."+NameField+" = 0, "+NameTable+".Показ = True WHERE ((("+NameTable+".Del)=False)); ";
//                  UPDATE Территория INNER JOIN Аспекты ON Территория.[Номер территории] = Аспекты.[Вид территории] SET Аспекты.[Вид территории] = 0 WHERE (((Территория.Показ)=True) AND ((Территория.Del)=False));
Comm2->Execute();

Comm2->CommandText="Delete * From "+NameTable+" Where "+NameTable+".Del=false AND "+NameTable+".Показ=true";
Comm2->Execute();

Comm2->CommandText="UPDATE "+NameTable+" SET "+NameTable+".Del = False";
Comm2->Execute();

MP<TADOCommand>Comm(Form1);
Comm->Connection=Database1;
Comm->CommandText="Delete * From TempBranch";
Comm->Execute();
}
else
{
//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Сбой", "Сбой записи узлов и ветвей (Full) (нет базы данных)", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase2);
Parent->WriteDiaryEvent(IP, Login, "Сбой", "Сбой записи узлов и ветвей (Full) (нет базы данных)", " DB: "+NameDatabase2);
}
}
//--------------------------------------------------------------------------
void Client::MergeNodeBranch(String NameDatabase1, String NameNode, String NameBranch)
{
TADOConnection *Database1=GetDatabase(NameDatabase1);



if(Database1!=NULL)
{
MP<TADOCommand>Comm(Form1);
Comm->Connection=Database1;
Comm->CommandText="Delete * From "+NameNode;
Comm->Execute();

Comm->CommandText="INSERT INTO "+NameNode+" ( [Номер узла], Родитель, Название ) SELECT TempNode.[Номер узла], TempNode.Родитель, TempNode.Название FROM TempNode;";
Comm->Execute();

Comm->CommandText="INSERT INTO "+NameBranch+" ( [Номер ветви], [Номер родителя], Название, Показ ) SELECT TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.Название, TempBranch.Показ FROM TempBranch;";
Comm->Execute();

Comm->CommandText="Delete * From TempNode";
Comm->Execute();

}
else
{
//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Сбой", "Сбой записи узлов и ветвей (Low) (нет базы данных)", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase1);
Parent->WriteDiaryEvent(IP, Login, "Сбой", "Сбой записи узлов и ветвей (Low) (нет базы данных)", " DB: "+NameDatabase1);


}

}
//---------------------------------------------------------------------------
String Client::PodrToLogin(String NameDB, int NumPodr)
{
String Ret="";
 MP<TADODataSet>Tab(Form1);
 Tab->Connection=GetDatabase(NameDB);
 Tab->CommandText="SELECT Logins.Login, ObslOtdel.NumObslOtdel FROM Logins INNER JOIN ObslOtdel ON Logins.Num = ObslOtdel.Login WHERE (((ObslOtdel.NumObslOtdel)="+IntToStr(NumPodr)+"));";
 Tab->Active=true;
 if(Tab->RecordCount!=0)
 {
  Ret=Tab->FieldByName("Login")->Value;
 }

 return Ret;
}
//---------------------------------------------------------------------------
void Client::MergeAspectsMainSpec()
{
MP<TADOCommand>Comm(Form1);
Comm->Connection=GetDatabase("Аспекты");
Comm->CommandText="UPDATE Аспекты SET Аспекты.Del = False";
Comm->Execute();

MP<TADODataSet>Aspects(Form1);
Aspects->Connection=GetDatabase("Аспекты");
Aspects->CommandText="Select * From Аспекты";
Aspects->Active=true;

MP<TADODataSet>Temp(Form1);
Temp->Connection=GetDatabase("Аспекты");
Temp->CommandText="Select * From TempAspects";
Temp->Active=true;

for(Aspects->First();!Aspects->Eof;Aspects->Next())
{
 int N=Aspects->FieldByName("номер аспекта")->Value;

 if(Temp->Locate("Номер аспекта", N, SO))
 {
  Aspects->Edit();
  Aspects->FieldByName("Подразделение")->Value=Temp->FieldByName("Подразделение")->Value;
  Aspects->Post();
 }
 else
 {
//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Ошибка обработки данных", "Ошибка объединения аспектов главным специалистом (не найден аспект)", "IDC="+IntToStr(IDC())+" Номер="+IntToStr(N));

 }
}
}
//---------------------------------------------------------------------------
void Client::MergeAspectsUser(int NumLogin)
{
TADOConnection *Database=GetDatabase("Аспекты");
MP<TADODataSet>TempPodr(Form1);
TempPodr->Connection=Database;
TempPodr->CommandText="Select * From TempПодразделения";
TempPodr->Active=true;

MP<TADODataSet>TempAsp(Form1);
TempAsp->Connection=Database;
TempAsp->CommandText="select * From TempAspects";
TempAsp->Active=true;

MP<TADODataSet>Asp(Form1);
Asp->Connection=Database;
Asp->CommandText="SELECT Аспекты.*, ObslOtdel.Login FROM (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel WHERE (((ObslOtdel.Login)="+IntToStr(NumLogin)+"));";
Asp->Active=true;
/*
for(TempAsp->First();!TempAsp->Eof;TempAsp->Next())
{
 int N=TempAsp->FieldByName("Подразделение")->Value;

 if(TempPodr->Locate("номер подразделения",N,SO))
 {
 int NumPodr=TempPodr->FieldByName("ServerNum")->Value;

 TempAsp->Edit();
 TempAsp->FieldByName("Подразделение")->Value=NumPodr;
 TempAsp->Post();
 }
 else
 {
//DiaryEvent->WriteEvent(Now(), this->pNameComp, this->Login, "Сбой", "Сбой объединения аспектов (не найдено подразделение)", "IDC="+IntToStr(IDC())+" DB: "+NameDatabase+" NumLogin: "+IntToStr(NumLogin));

 }
}
 */

MP<TADOCommand>Comm(Form1);
Comm->Connection=Database;
Comm->CommandText="UPDATE (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel SET Аспекты.Del = False WHERE (((ObslOtdel.Login)="+IntToStr(NumLogin)+"));";
Comm->Execute();

for(Asp->First();!Asp->Eof;Asp->Next())
{
 int N=Asp->FieldByName("Номер аспекта")->Value;

 if(TempAsp->Locate("Номер аспекта", N, SO))
 {
  Asp->Edit();
  Asp->FieldByName("Подразделение")->Value=TempAsp->FieldByName("Подразделение")->Value;
  Asp->FieldByName("Ситуация")->Value=TempAsp->FieldByName("Ситуация")->Value;
  Asp->FieldByName("Вид территории")->Value=TempAsp->FieldByName("Вид территории")->Value;
  Asp->FieldByName("Деятельность")->Value=TempAsp->FieldByName("Деятельность")->Value;
  Asp->FieldByName("Специальность")->Value=TempAsp->FieldByName("Специальность")->Value;
  Asp->FieldByName("Аспект")->Value=TempAsp->FieldByName("Аспект")->Value;
  Asp->FieldByName("Воздействие")->Value=TempAsp->FieldByName("Воздействие")->Value;
  Asp->FieldByName("G")->Value=TempAsp->FieldByName("G")->Value;
  Asp->FieldByName("O")->Value=TempAsp->FieldByName("O")->Value;
  Asp->FieldByName("R")->Value=TempAsp->FieldByName("R")->Value;
  Asp->FieldByName("S")->Value=TempAsp->FieldByName("S")->Value;
  Asp->FieldByName("T")->Value=TempAsp->FieldByName("T")->Value;
  Asp->FieldByName("L")->Value=TempAsp->FieldByName("L")->Value;
  Asp->FieldByName("N")->Value=TempAsp->FieldByName("N")->Value;
  Asp->FieldByName("Z")->Value=TempAsp->FieldByName("Z")->Value;
  Asp->FieldByName("Значимость")->Value=TempAsp->FieldByName("Значимость")->Value;
  Asp->FieldByName("Проявление воздействия")->Value=TempAsp->FieldByName("Проявление воздействия")->Value;
  Asp->FieldByName("Тяжесть последствий")->Value=TempAsp->FieldByName("Тяжесть последствий")->Value;
  Asp->FieldByName("Приоритетность")->Value=TempAsp->FieldByName("Приоритетность")->Value;
  Asp->FieldByName("Выполняющиеся мероприятия")->Value=TempAsp->FieldByName("Выполняющиеся мероприятия")->Value;
  Asp->FieldByName("Предлагаемые мероприятия")->Value=TempAsp->FieldByName("Предлагаемые мероприятия")->Value;
  Asp->FieldByName("Мониторинг и контроль")->Value=TempAsp->FieldByName("Мониторинг и контроль")->Value;
  Asp->FieldByName("Предлагаемый мониторинг и контроль")->Value=TempAsp->FieldByName("Предлагаемый мониторинг и контроль")->Value;
  Asp->FieldByName("Исполнитель")->Value=TempAsp->FieldByName("Исполнитель")->Value;
  Asp->FieldByName("Дата создания")->Value=TempAsp->FieldByName("Дата создания")->Value;
  Asp->FieldByName("Начало действия")->Value=TempAsp->FieldByName("Начало действия")->Value;
  Asp->FieldByName("Конец действия")->Value=TempAsp->FieldByName("Конец действия")->Value;
  Asp->Post();

  TempAsp->Delete();
 }
 else
 {
  Asp->Edit();
  Asp->FieldByName("Del")->Value=true;
  Asp->Post();
 }
}

Comm->CommandText="DELETE Аспекты.*, Аспекты.Del, ObslOtdel.Login FROM (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel WHERE (((Аспекты.Del)=True) AND ((ObslOtdel.Login)="+IntToStr(NumLogin)+"));";
Comm->Execute();

for(TempAsp->First();!TempAsp->Eof;TempAsp->Next())
{
  Asp->Insert();
  Asp->FieldByName("Подразделение")->Value=TempAsp->FieldByName("Подразделение")->Value;
  Asp->FieldByName("Ситуация")->Value=TempAsp->FieldByName("Ситуация")->Value;
  Asp->FieldByName("Вид территории")->Value=TempAsp->FieldByName("Вид территории")->Value;
  Asp->FieldByName("Деятельность")->Value=TempAsp->FieldByName("Деятельность")->Value;
  Asp->FieldByName("Специальность")->Value=TempAsp->FieldByName("Специальность")->Value;
  Asp->FieldByName("Аспект")->Value=TempAsp->FieldByName("Аспект")->Value;
  Asp->FieldByName("Воздействие")->Value=TempAsp->FieldByName("Воздействие")->Value;
  Asp->FieldByName("G")->Value=TempAsp->FieldByName("G")->Value;
  Asp->FieldByName("O")->Value=TempAsp->FieldByName("O")->Value;
  Asp->FieldByName("R")->Value=TempAsp->FieldByName("R")->Value;
  Asp->FieldByName("S")->Value=TempAsp->FieldByName("S")->Value;
  Asp->FieldByName("T")->Value=TempAsp->FieldByName("T")->Value;
  Asp->FieldByName("L")->Value=TempAsp->FieldByName("L")->Value;
  Asp->FieldByName("N")->Value=TempAsp->FieldByName("N")->Value;
  Asp->FieldByName("Z")->Value=TempAsp->FieldByName("Z")->Value;
  Asp->FieldByName("Значимость")->Value=TempAsp->FieldByName("Значимость")->Value;
  Asp->FieldByName("Проявление воздействия")->Value=TempAsp->FieldByName("Проявление воздействия")->Value;
  Asp->FieldByName("Тяжесть последствий")->Value=TempAsp->FieldByName("Тяжесть последствий")->Value;
  Asp->FieldByName("Приоритетность")->Value=TempAsp->FieldByName("Приоритетность")->Value;
  Asp->FieldByName("Выполняющиеся мероприятия")->Value=TempAsp->FieldByName("Выполняющиеся мероприятия")->Value;
  Asp->FieldByName("Предлагаемые мероприятия")->Value=TempAsp->FieldByName("Предлагаемые мероприятия")->Value;
  Asp->FieldByName("Мониторинг и контроль")->Value=TempAsp->FieldByName("Мониторинг и контроль")->Value;
  Asp->FieldByName("Предлагаемый мониторинг и контроль")->Value=TempAsp->FieldByName("Предлагаемый мониторинг и контроль")->Value;
  //Asp->FieldByName("Исполнитель")->Value=TempAsp->FieldByName("Исполнитель")->Value;
  Asp->FieldByName("Дата создания")->Value=TempAsp->FieldByName("Дата создания")->Value;
  Asp->FieldByName("Начало действия")->Value=TempAsp->FieldByName("Начало действия")->Value;
  Asp->FieldByName("Конец действия")->Value=TempAsp->FieldByName("Конец действия")->Value;
  Asp->Post();

  Asp->Active=false;
  Asp->Active=true;
  Asp->Last();

  TempAsp->Edit();
  TempAsp->FieldByName("ServerNum")->Value=Asp->FieldByName("Номер аспекта")->Value;
  TempAsp->Post();
}

Comm->CommandText="Delete * From TempПодразделения";
Comm->Execute();
}
//---------------------------------------------------------------------------
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

