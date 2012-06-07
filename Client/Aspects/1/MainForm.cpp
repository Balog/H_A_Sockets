//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "MainForm.h"
#include "Zastavka.h"
#include "About.h"
#include "MasterPointer.h"
//#include "InputFiltr.h"
//#include "LoadKeyFile.h"
//#include "Svod.h"
//#include "Rep1.h"
//#include "File_operations.cpp"

#include "Progress.h"
//#include "F_Vvedenie.h"
//#include "Metod.h"
//#include "Rep1.h"
#include "About.h"
#include "InpDocs.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
Aspects->Connection=Zast->ADOUsrAspect;
Podrazdel->Connection=Zast->ADOUsrAspect;
Situaciya->Connection=Zast->ADOUsrAspect;
Territoriya->Connection=Zast->ADOUsrAspect;
Deyatelnost->Connection=Zast->ADOUsrAspect;
Aspect->Connection=Zast->ADOUsrAspect;
Vozdeystvie->Connection=Zast->ADOUsrAspect;
Posledstvie->Connection=Zast->ADOUsrAspect;
Tiagest->Connection=Zast->ADOUsrAspect;
Prioritet->Connection=Zast->ADOUsrAspect;
TempAspects->Connection=Zast->ADOUsrAspect;
Znachimost->Connection=Zast->ADOUsrAspect;
Report->Connection=Zast->ADOUsrAspect;

Filtr1="";
Filtr2="Значимость=True";
/*
CurrentRecord=0;
Aspects->Active=false;
Aspects->CommandText="select * from Аспекты ORDER BY [Номер аспекта]";
Aspects->Active=true;
Filtr1="select * from Аспекты";
Filtr2="SELECT * From Аспекты WHERE Значимость=True";
*/

Registered=false;
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormShow(TObject *Sender)
{

//ShowMessage("Пользователь");
//Pass->Hide();
//Pass->Close();
//Zast->MClient->Start();

//******************************************
//       Удаление демо-записей аспектов
//     если они попали не демопользователю
//         в результате манипуляций
//******************************************
 MP<TADOCommand>Comm(this);
 Comm->Connection=Zast->ADOUsrAspect;
if(!Demo)
{
 Comm->CommandText="DELETE Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+") AND ((Аспекты.Demo)=True));";
  //Comm->CommandText="DELETE Logins.AdmNum, Аспекты.Demo, Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.AdmNum)="+IntToStr(NumLogin)+") AND ((Аспекты.Demo)=True)); ";

 Comm->Execute();
}
else
{
//******************************************
// Пометить все аспекты демопользователя как Demo
//******************************************
 Comm->CommandText="UPDATE Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login SET Аспекты.Demo = True WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+")); ";
 Comm->Execute();
}
//****************************************

N3->Enabled=!Demo;
Initialize();
//N9->Enabled=!Demo;

/*
if(!Registered)
{
Zast->MClient->Start();
Zast->MClient->RegForm(this);
Registered=true;

MP<TADODataSet>Logins(this);
Logins->Connection=Zast->ADOUsrAspect;
Logins->CommandText="Select * From Logins Where ServerNum="+IntToStr(NumLogin);
Logins->Active=true;

Login=Logins->FieldByName("Login")->AsString;
if(Demo)
{
this->Caption="ДЕМОНСТРАЦИОННЫЙ РЕЖИМ!   Пользователь: "+Login;
Zast->MClient->WriteDiaryEvent("NetAspects","Деморежим",Login);

}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects","Пользователь",Login);
this->Caption="Пользователь: "+Login;
}

Initialize();
Zast->MClient->Stop();
}


Zast->MClient->WriteDiaryEvent("NetAspects","Запуск модуля пользователя завершен","");
*/
}
//---------------------------------------------------------------------------
void TForm1::Initialize()
{
//ShowMessage("Начало обновления логинов");
SetAspects(Login);
//ShowMessage("Обновление справочников");
LoadSpravs();
//ShowMessage("Начало");

SetCombo();

Aspects->Active=true;
if(Podrazdel->RecordCount!=0)
{
/*
if(Aspects->RecordCount==0)
{
 NewRecord();

}
*/
}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects","Для пользователя не назначено подразделений",Login);
ShowMessage("Для пользователя "+Login+" не назначено подразделений!\rЗавершение работы программы.");
Zast->Close();
}
 InitCombo();
}
//----------------------------------------------------------------------------
void TForm1::SetAspects(String Login)
{
//Список логинов, подразделений и обслуживаемых подразделений уже обновлен в локальной базе.
//Если подразделение уже отсутствует то аспекты удалены автоматически.
//Если изменены обслуживаемые подразделения - аспекты остаются на месте

//Через логин, таблицу обслуживаемых подразделений и таблицу подразделений
//запрашиваем из локальной базы все аспекты что обслуживаются текущим логином
//Задаем условия фильтрации обслуживаемых аспектов

//Задаем условия фильтрации списка подразделений для заполнения списка подразделений

Aspects->Active=false;
Aspects->CommandText="SELECT  Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+")) ORDER BY Аспекты.[Номер аспекта];";
Aspects->Active=true;

Podrazdel->Active=false;
Podrazdel->CommandText="SELECT Подразделения.* FROM Logins INNER JOIN (Подразделения INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+")) ORDER BY Подразделения.[Название подразделения];";
Podrazdel->Active=true;

Situaciya->Active=false;
Situaciya->CommandText="select * from Ситуации where [Номер ситуации]<>0";
Situaciya->Active=true;
}
//---------------------------------------------------------
bool TForm1::LoadSpravs()
{
/*
bool Ret=true;
//Проверка обновления всех справочников и синхронизация таблиц.
//После синхронизации каждой таблицы обработка отфильтрованой таблицы аспектов (принадлежащей текущему логину)
Zast->MClient->WriteDiaryEvent("NetAspects","Начало проверки необходимости загрузки справочников (автоматическое)","");
try
{
FProg->Progress->Position=0;
FProg->Progress->Min=0;
FProg->Progress->Max=7;
FProg->Visible=true;


FProg->Progress->Position++;
FProg->Label1->Caption="Чтение справочника ситуаций...";
FProg->Label1->Repaint();
if(VerifySit())
{
Ret=Ret & LoadSit();
}

FProg->Progress->Position++;
FProg->Label1->Caption="Чтение справочника территорий...";
FProg->Label1->Repaint();
if(VerifyTerr())
{
Ret=Ret & LoadTerr();
}

FProg->Progress->Position++;
FProg->Label1->Caption="Чтение справочника видов деятельностей...";
FProg->Label1->Repaint();
if(VerifyDeyat())
{
Ret=Ret & LoadDeyat();
}

FProg->Progress->Position++;
FProg->Label1->Caption="Чтение справочника аспектов...";
FProg->Label1->Repaint();
if(VerifyAsp())
{
Ret=Ret & LoadAsp();
}

FProg->Progress->Position++;
FProg->Label1->Caption="Чтение справочника видов воздействий...";
FProg->Label1->Repaint();
if(VerifyVozd())
{
Ret=Ret & LoadVozd();
}

FProg->Progress->Position++;
FProg->Label1->Caption="Чтение справочника критериев оценки...";
FProg->Label1->Repaint();
if(VerifyCryt())
{
Ret=Ret & LoadCryt();
}

FProg->Progress->Position++;
FProg->Label1->Caption="Чтение справочника мероприятий...";
FProg->Label1->Repaint();
VerifyMeropr();

FProg->Visible=false;

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой проверки необходимости загрузки справочников (автоматическое)","");
ShowMessage("Во время обновления справочников произошли ошибки");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка проверки необходимости загрузки справочников (автоматическое)"," Ошибка "+IntToStr(GetLastError()));
}
Zast->MClient->WriteDiaryEvent("NetAspects","Конец проверки необходимости загрузки справочников (автоматическое)","");
return Ret;
*/
}
//---------------------------------------------------------
void TForm1::SetCombo()
{
Aspects->Active=false;
Podrazdel->Active=false;
Situaciya->Active=false;
Territoriya->Active=false;
Deyatelnost->Active=false;
Aspect->Active=false;
Vozdeystvie->Active=false;
Posledstvie->Active=false;
Tiagest->Active=false;
Prioritet->Active=false;

Aspects->Active=true;
Podrazdel->Active=true;
Situaciya->Active=true;
Territoriya->Active=true;
Deyatelnost->Active=true;
Aspect->Active=true;
Vozdeystvie->Active=true;
Posledstvie->Active=true;
Tiagest->Active=true;
Prioritet->Active=true;

CPodrazdel->Clear();
for(Podrazdel->First();!Podrazdel->Eof;Podrazdel->Next())
{
 CPodrazdel->Items->Add(Podrazdel->FieldByName("Название подразделения")->AsString);
}

CSit->Clear();
for(Situaciya->First();!Situaciya->Eof;Situaciya->Next())
{
 CSit->Items->Add(Situaciya->FieldByName("Название ситуации")->AsString);
}

CProyav->Clear();
for(Posledstvie->First();!Posledstvie->Eof;Posledstvie->Next())
{
 CProyav->Items->Add(Posledstvie->FieldByName("Наименование проявления")->AsString);
}

CPosl->Clear();
for(Tiagest->First();!Tiagest->Eof;Tiagest->Next())
{
 CPosl->Items->Add(Tiagest->FieldByName("Наименование последствия")->AsString);
}

CPrior->Clear();
for(Prioritet->First();!Prioritet->Eof;Prioritet->Next())
{
 CPrior->Items->Add(Prioritet->FieldByName("Наименование приоритетности")->AsString);
}
}
//---------------------------------------------------------
void TForm1::InitCombo()
{
Aspects->Active=true;
Podrazdel->Active=true;
Situaciya->Active=true;
Territoriya->Active=true;
Deyatelnost->Active=true;
Aspect->Active=true;
Vozdeystvie->Active=true;
Posledstvie->Active=true;
Tiagest->Active=true;
Prioritet->Active=true;
//Zmax->Text=Zast->Zmax;
 int N1=0,N2=0,N3=0,N4=0,N5=0;
// int Num1,Num2,Num3,Num4,Num5;

if(Aspects->RecordCount!=0)
{
EnabledForm(true);



 N1=Aspects->FieldByName("Подразделение")->Value;
 N2=Aspects->FieldByName("Ситуация")->Value;
 N3=Aspects->FieldByName("Проявление воздействия")->Value;
 N4=Aspects->FieldByName("Тяжесть последствий")->Value;
 N5=Aspects->FieldByName("Приоритетность")->Value;



if(Podrazdel->Locate("Номер подразделения",N1, SO))
{
CPodrazdel->ItemIndex=Podrazdel->RecNo-1;
}
else
{
ShowMessage("Ошибка подразделения");
}

if(N2==0)
{
CSit->Style=csDropDown;
CSit->Text="Выберите ситуацию";
}
else
{
if(Situaciya->Locate("Номер ситуации",N2, SO))
{
CSit->Style=csDropDownList;
CSit->ItemIndex=Situaciya->RecNo-1;
}
else
{
ShowMessage("Ошибка ситуации");
}
}

if(Posledstvie->Locate("Номер проявления",N3, SO))
{
CProyav->ItemIndex=Posledstvie->RecNo-1;
}
else
{
ShowMessage("Ошибка проявления");
}

if(Tiagest->Locate("Номер последствия",N4, SO))
{
CPosl->ItemIndex=Tiagest->RecNo-1;
}
else
{
ShowMessage("Ошибка проявления");
}

if(Prioritet->Locate("Номер приоритетности",N5, SO))
{
CPrior->ItemIndex=Prioritet->RecNo-1;
}
else
{
ShowMessage("Ошибка приоритетности");
}



if (Aspects->RecordCount!=0)
{
DateTimePicker1->Date=Aspects->FieldByName("Дата создания")->Value;
DateTimePicker2->Date=Aspects->FieldByName("Начало действия")->Value;
DateTimePicker3->Date=Aspects->FieldByName("Конец действия")->Value;
}

//----Инициализация из документов-----------------------

int K;
if(Aspects->RecordCount!=0)
{
K=Aspects->FieldByName("Вид территории")->Value;
}
else
{
K=0;
}
Territoriya->Active=false;
Territoriya->CommandText="Select * From Территория Where [Номер территории]="+IntToStr(K);
Territoriya->Active=true;
Edit1->Text=Territoriya->FieldByName("Наименование территории")->Value;

if(Aspects->RecordCount!=0)
{
K=Aspects->FieldByName("Деятельность")->Value;
}
else
{
K=0;
}
Deyatelnost->Active=false;
Deyatelnost->CommandText="Select * From Деятельность Where [Номер деятельности]="+IntToStr(K);
Deyatelnost->Active=true;
Edit2->Text=Deyatelnost->FieldByName("Наименование деятельности")->Value;

if(Aspects->RecordCount!=0)
{
K=Aspects->FieldByName("Аспект")->Value;
}
else
{
K=0;
}
Aspect->Active=false;
Aspect->CommandText="Select * From Аспект Where [Номер аспекта]="+IntToStr(K);
Aspect->Active=true;
Edit4->Text=Aspect->FieldByName("Наименование аспекта")->Value;

if(Aspects->RecordCount!=0)
{
K=Aspects->FieldByName("Воздействие")->Value;
}
else
{
K=0;
}
Vozdeystvie->Active=false;
Vozdeystvie->CommandText="Select * From Воздействия Where [Номер воздействия]="+IntToStr(K);
Vozdeystvie->Active=true;
Edit5->Text=Vozdeystvie->FieldByName("Наименование воздействия")->Value;


//----------
if(Aspects->RecordCount!=0)
{
int N=Aspects->FieldByName("Номер аспекта")->Value;

TempAspects->Active=false;
TempAspects->CommandText="Select * From Аспекты Where [Номер аспекта]="+IntToStr(N);
TempAspects->Active=true;
if(TempAspects->RecordCount!=0)
{
if(TempAspects->FieldByName("G")->Value==1)
{

C1=true;

}
else
{

C1=false;

}

if(TempAspects->FieldByName("R")->Value==1)
{

C2=true;

}
else
{

C2=false;

}

if(TempAspects->FieldByName("T")->Value==1)
{

C3=true;

}
else
{
//
C3=false;

}

if(TempAspects->FieldByName("L")->Value==1)
{

C4=true;

}
else
{

C4=false;

}

if(TempAspects->FieldByName("O")->Value==1)
{

C5=true;

}
else
{

C5=false;

}

if(TempAspects->FieldByName("S")->Value==1)
{

C6=true;

}
else
{

C6=false;


}

if(TempAspects->FieldByName("N")->Value==1)
{

C7=true;

}
else
{

C7=false;

}

if(C1==1)
{
CheckBox1->Checked=true;
}
else
{
CheckBox1->Checked=false;
}

if(C2==1)
{
CheckBox2->Checked=true;
}
else
{
CheckBox2->Checked=false;
}

if(C3==1)
{
CheckBox3->Checked=true;
}
else
{
CheckBox3->Checked=false;
}

if(C4==1)
{
CheckBox4->Checked=true;
}
else
{
CheckBox4->Checked=false;
}

if(C5==1)
{
CheckBox5->Checked=true;
}
else
{
CheckBox5->Checked=false;
}

if(C6==1)
{
CheckBox6->Checked=true;
}
else
{
CheckBox6->Checked=false;
}

if(C7==1)
{
CheckBox7->Checked=true;
}
else
{
CheckBox7->Checked=false;
}
}
}



//----Позиция записи------------
if(Aspects->RecordCount==0 & Podrazdel->RecordCount!=0)
{
 NewRecord();
}
CountRecord->Text=IntToStr(Aspects->RecordCount);

NumRecord->Text=IntToStr(Aspects->RecNo);


Calc();
}
else
{
EnabledForm(false);
NumRecord->Text="0";
CountRecord->Text="0";
}
}
//---------------------------------------------------------------------------
void TForm1::NewRecord()
{
Podrazdel->Active=false;
Podrazdel->Active=true;
Situaciya->Active=false;
Situaciya->Active=true;
Posledstvie->Active=true;
Tiagest->Active=true;
Prioritet->Active=true;



/*
if (Podrazdel->RecordCount==0)
{
 Podrazdel->Append();
 Podrazdel->FieldByName("Название подразделения")->Value="Новое подразделение";
 Podrazdel->Post();
}
*/
if (Situaciya->RecordCount==0)
{
 Situaciya->Append();
 Situaciya->FieldByName("Название ситуации")->Value="Новая ситуация";
 Situaciya->Post();
}

Podrazdel->First();
Situaciya->First();
Posledstvie->First();
Tiagest->First();
Prioritet->First();

 Aspects->Append();
 Aspects->FieldByName("Вид территории")->Value=0;
 Aspects->FieldByName("Деятельность")->Value=0;
 Aspects->FieldByName("Аспект")->Value=0;
 Aspects->FieldByName("Воздействие")->Value=0;
 //Aspects->FieldByName("Значимость")->Value=1;

 int NP=Podrazdel->FieldByName("Номер подразделения")->Value;
 Aspects->FieldByName("Подразделение")->Value=NP;
 int NS=Situaciya->FieldByName("Номер ситуации")->Value;
 Aspects->FieldByName("Ситуация")->Value=NS;
 Aspects->FieldByName("Специальность")->Value=" ";
// Aspects->FieldByName("Подразделение")->Value=Podrazdel->FieldByName("Номер подразделения")->Value;
// Aspects->FieldByName("Ситуация")->Value=Situaciya->FieldByName("Номер ситуации")->Value;
 Aspects->FieldByName("Тяжесть последствий")->Value=Tiagest->FieldByName("Номер последствия")->Value;
 Aspects->FieldByName("Проявление воздействия")->Value=Posledstvie->FieldByName("Номер проявления")->Value;
 Aspects->FieldByName("Приоритетность")->Value=Prioritet->FieldByName("Номер приоритетности")->Value;
 Aspects->FieldByName("Предлагаемые мероприятия")->Value="";
 Aspects->FieldByName("Дата создания")->Value=Now();
 Aspects->FieldByName("Начало действия")->Value=Now();
 Aspects->FieldByName("Конец действия")->Value=Now();
 Aspects->FieldByName("Demo")->Value=Demo;

 Aspects->Post();

 CheckBox1->Checked=false;
CheckBox2->Checked=false;
CheckBox3->Checked=false;
CheckBox4->Checked=false;
CheckBox5->Checked=false;
CheckBox6->Checked=false;
CheckBox7->Checked=false;

CProyav->ItemIndex=0;
CPosl->ItemIndex=0;
CPodrazdel->ItemIndex=0;
CSit->ItemIndex=0;
 Calc();
 InitCombo();

}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormClose(TObject *Sender, TCloseAction &Action)
{
switch (Application->MessageBoxA("Записать все изменения на сервер?","Выход из программы",MB_YESNOCANCEL	+MB_ICONQUESTION+MB_DEFBUTTON1))
{
 case IDYES:
 {
Zast->Stop=true;
Zast->MClient->Act.WaitCommand=0;

Zast->Saved=false;
//Zast->PrepareSaveLogins->Execute();


 Action=caNone;
 break;
 }
 case IDNO:
 {
Zast->MClient->WriteDiaryEvent("NetAspects","Завершение работы отказ от записи","");
Zast->Stop=true;
Sleep(2000);
Zast->Close();
 Action=caNone;
 break;
 }
 case IDCANCEL:
 {
 Action=caNone;
 break;
 }
}


}
//---------------------------------------------------------------------------
void __fastcall TForm1::N8Click(TObject *Sender)
{
//Zast->Close();
this->Close();
}
//---------------------------------------------------------------------------

void TForm1::SynhronTerr()
{

try
{
Territoriya->Active=false;
Territoriya->CommandText="Select * From Территория Where Показ=true";
Territoriya->Active=true;
Territoriya->First();
for(int i=0;i<Territoriya->RecordCount;i++)
{
int Num=Territoriya->FieldByName("Номер территории")->Value;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_5 Where Показ=true and [Номер ветви]="+IntToStr(Num);
Branches->Active=true;
if (Branches->RecordCount==0)
{
 TempAspects->Active=false;
 TempAspects->CommandText="Select Аспекты.[Номер аспекта], Аспекты.[Вид территории] From Аспекты Where [Вид территории]="+IntToStr(Num);
 TempAspects->Active=true;
 TempAspects->First();
 for(int j=0;j<TempAspects->RecordCount;j++)
 {
  TempAspects->Edit();
  TempAspects->FieldByName("Вид территории")->Value=0;
  TempAspects->Post();
  TempAspects->Next();
 }
 Territoriya->Delete();
}
else
{
 AnsiString Text=Branches->FieldByName("Название")->Value;
 Territoriya->Edit();
 Territoriya->FieldByName("Наименование территории")->Value=Text;
 Territoriya->Post();
}
 Territoriya->Next();
}
Branches->Active=false;
Branches->CommandText="Select * From Ветви_5 Where Показ=true";
Branches->Active=true;
Branches->First();
for(int i=0;i<Branches->RecordCount;i++)
{
int Num=Branches->FieldByName("Номер ветви")->Value;
 Territoriya->Active=false;
 Territoriya->CommandText="Select * From Территория Where [Номер территории]="+IntToStr(Num);
 Territoriya->Active=true;
 if (Territoriya->RecordCount==0)
 {
  int Num=Branches->FieldByName("Номер ветви")->Value;
  AnsiString Text=Branches->FieldByName("Название")->Value;

 Territoriya->Active=false;
 Territoriya->CommandText="Select * From Территория";
 Territoriya->Active=true;

  Territoriya->Append();
  Territoriya->FieldByName("Номер территории")->Value=Num;
  Territoriya->FieldByName("Наименование территории")->Value=Text;
  Territoriya->FieldByName("Показ")->Value=true;
  Territoriya->Post();
 }
Branches->Next();
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка синхронизации территорий"," Ошибка "+IntToStr(GetLastError()));
}
}
//---------------------------------------------------------------------------
void TForm1::SynhronDeyat()
{
try
{
Deyatelnost->Active=false;
Deyatelnost->CommandText="Select * From Деятельность Where Показ=true";
Deyatelnost->Active=true;
Deyatelnost->First();
for(int i=0;i<Deyatelnost->RecordCount;i++)
{
int Num=Deyatelnost->FieldByName("Номер деятельности")->Value;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_6 Where Показ=true and [Номер ветви]="+IntToStr(Num);
Branches->Active=true;

if (Branches->RecordCount==0)
{
 TempAspects->Active=false;
 TempAspects->CommandText="Select Аспекты.[Номер аспекта], Аспекты.[Деятельность] From Аспекты Where [Вид территории]="+IntToStr(Num);
 TempAspects->Active=true;
 TempAspects->First();
 for(int j=0;j<TempAspects->RecordCount;j++)
 {
  TempAspects->Edit();
  TempAspects->FieldByName("Деятельность")->Value=0;
  TempAspects->Post();
  TempAspects->Next();
 }
Deyatelnost->Delete();
}
else
{
 AnsiString Text=Branches->FieldByName("Название")->Value;
 Deyatelnost->Edit();
 Deyatelnost->FieldByName("Наименование деятельности")->Value=Text;
 Deyatelnost->Post();
}
 Deyatelnost->Next();
}
Branches->Active=false;
Branches->CommandText="Select * From Ветви_6 Where Показ=true";
Branches->Active=true;
Branches->First();
for(int i=0;i<Branches->RecordCount;i++)
{
int Num=Branches->FieldByName("Номер ветви")->Value;
 Deyatelnost->Active=false;
 Deyatelnost->CommandText="Select * From Деятельность Where [Номер деятельности]="+IntToStr(Num);
 Deyatelnost->Active=true;
 if (Deyatelnost->RecordCount==0)
 {
  int Num=Branches->FieldByName("Номер ветви")->Value;
  AnsiString Text=Branches->FieldByName("Название")->Value;
  Deyatelnost->Append();
  Deyatelnost->FieldByName("Номер деятельности")->Value=Num;
  Deyatelnost->FieldByName("Наименование деятельности")->Value=Text;
  Deyatelnost->FieldByName("Показ")->Value=true;
  Deyatelnost->Post();
 }
Branches->Next();
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка синхронизации деятельностей","");
}
}
//---------------------------------------------------------------------------
void TForm1::SynhronAspect()
{
try
{
Aspect->Active=false;
Aspect->CommandText="Select * From Аспект Where Показ=true";
Aspect->Active=true;
Aspect->First();
for(int i=0;i<Aspect->RecordCount;i++)
{
int Num=Aspect->FieldByName("Номер аспекта")->Value;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_7 Where Показ=true and [Номер ветви]="+IntToStr(Num);
Branches->Active=true;

if (Branches->RecordCount==0)
{
 TempAspects->Active=false;
 TempAspects->CommandText="Select * From Аспекты Where [Аспект]="+IntToStr(Num);
 TempAspects->Active=true;
 TempAspects->First();
 for(int j=0;j<TempAspects->RecordCount;j++)
 {
  TempAspects->Edit();
  TempAspects->FieldByName("Аспект")->Value=0;
  TempAspects->Post();
  TempAspects->Next();
 }
Aspect->Delete();
}
else
{
 AnsiString Text=Branches->FieldByName("Название")->Value;
Aspect->Edit();
Aspect->FieldByName("Наименование аспекта")->Value=Text;
Aspect->Post();
}
Aspect->Next();
}
Branches->Active=false;
Branches->CommandText="Select * From Ветви_7 Where Показ=true";
Branches->Active=true;
Branches->First();
for(int i=0;i<Branches->RecordCount;i++)
{
int Num=Branches->FieldByName("Номер ветви")->Value;
Aspect->Active=false;
Aspect->CommandText="Select * From Аспект Where [Номер аспекта]="+IntToStr(Num);
Aspect->Active=true;
 if (Aspect->RecordCount==0)
 {
  int Num=Branches->FieldByName("Номер ветви")->Value;
  AnsiString Text=Branches->FieldByName("Название")->Value;
Aspect->Append();
Aspect->FieldByName("Номер аспекта")->Value=Num;
Aspect->FieldByName("Наименование аспекта")->Value=Text;
Aspect->FieldByName("Показ")->Value=true;
Aspect->Post();
 }
Branches->Next();
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка синхронизации экоаспектов","");
}
}
//---------------------------------------------------------------------------
void TForm1::SynhronVozd()
{
try
{
Vozdeystvie->Active=false;
Vozdeystvie->CommandText="Select * From Воздействия Where Показ=true";
Vozdeystvie->Active=true;
Vozdeystvie->First();
for(int i=0;i<Vozdeystvie->RecordCount;i++)
{
int Num=Vozdeystvie->FieldByName("Номер воздействия")->Value;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_3 Where Показ=true and [Номер ветви]="+IntToStr(Num);
Branches->Active=true;

if (Branches->RecordCount==0)
{
 TempAspects->Active=false;
 TempAspects->CommandText="Select Аспекты.[Номер аспекта], Аспекты.[Деятельность] From Аспекты Where [Вид территории]="+IntToStr(Num);
 TempAspects->Active=true;
 TempAspects->First();
 for(int j=0;j<TempAspects->RecordCount;j++)
 {
  TempAspects->Edit();
  TempAspects->FieldByName("Воздействие")->Value=0;
  TempAspects->Post();
  TempAspects->Next();
 }
Vozdeystvie->Delete();
}
else
{
 AnsiString Text=Branches->FieldByName("Название")->Value;
Vozdeystvie->Edit();
Vozdeystvie->FieldByName("Наименование воздействия")->Value=Text;
Vozdeystvie->Post();
}
Vozdeystvie->Next();
}
Branches->Active=false;
Branches->CommandText="Select * From Ветви_3 Where Показ=true";
Branches->Active=true;
Branches->First();
for(int i=0;i<Branches->RecordCount;i++)
{
int Num=Branches->FieldByName("Номер ветви")->Value;
Vozdeystvie->Active=false;
Vozdeystvie->CommandText="Select * From Воздействия Where [Номер воздействия]="+IntToStr(Num);
Vozdeystvie->Active=true;
 if (Vozdeystvie->RecordCount==0)
 {
  int Num=Branches->FieldByName("Номер ветви")->Value;
  AnsiString Text=Branches->FieldByName("Название")->Value;
Vozdeystvie->Append();
Vozdeystvie->FieldByName("Номер воздействия")->Value=Num;
Vozdeystvie->FieldByName("Наименование воздействия")->Value=Text;
Vozdeystvie->FieldByName("Показ")->Value=true;
Vozdeystvie->Post();
 }
Branches->Next();
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка синхронизации воздействий","");
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPodrazdelClick(TObject *Sender)
{
int N;
Podrazdel->First();
Podrazdel->MoveBy(CPodrazdel->ItemIndex);
N=Podrazdel->FieldByName("Номер подразделения")->Value;
Aspects->Edit();
Aspects->FieldByName("Подразделение")->Value=N;
Aspects->Post();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CSitClick(TObject *Sender)
{
int N;
Situaciya->First();
Situaciya->MoveBy(CSit->ItemIndex);
N=Situaciya->FieldByName("Номер ситуации")->Value;
Aspects->Edit();
Aspects->FieldByName("Ситуация")->Value=N;
Aspects->Post();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CProyavClick(TObject *Sender)
{
int N;
Posledstvie->First();
Posledstvie->MoveBy(CProyav->ItemIndex);
N=Posledstvie->FieldByName("Номер проявления")->Value;
Aspects->Edit();
Aspects->FieldByName("Проявление воздействия")->Value=N;
Aspects->Post();
Aspects->UpdateBatch();
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPoslClick(TObject *Sender)
{
int N;
Tiagest->First();
Tiagest->MoveBy(CPosl->ItemIndex);
N=Tiagest->FieldByName("Номер последствия")->Value;
Aspects->Edit();
Aspects->FieldByName("Тяжесть последствий")->Value=N;
Aspects->Post();
Aspects->UpdateBatch();
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPriorClick(TObject *Sender)
{
int N;
Prioritet->First();
Prioritet->MoveBy(CPrior->ItemIndex);
N=Prioritet->FieldByName("Номер приоритетности")->Value;
Aspects->Edit();
Aspects->FieldByName("Приоритетность")->Value=N;
Aspects->Post();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn5Click(TObject *Sender)
{


NewRecord();
C1=false;
C2=false;
C3=false;
C4=false;
C5=false;
C6=false;
C7=false;
InitCombo();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn1Click(TObject *Sender)
{
Aspects->First();
InitCombo();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn4Click(TObject *Sender)
{
Aspects->Last();
InitCombo();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn2Click(TObject *Sender)
{

Aspects->Prior();
InitCombo();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn3Click(TObject *Sender)
{

Aspects->Next();
InitCombo();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn6Click(TObject *Sender)
{
//if(Aspects->RecordCount>1)
{
Aspects->Delete();
InitCombo();
}
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button1Click(TObject *Sender)
{
/*
InputDocs->IForm=1;
InputDocs->Mode=1;
InputDocs->ShowModal();
*/
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button2Click(TObject *Sender)
{
/*
InputDocs->IForm=1;
InputDocs->Mode=2;
InputDocs->ShowModal();
*/
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button3Click(TObject *Sender)
{
/*
InputDocs->IForm=1;
InputDocs->Mode=3;
InputDocs->ShowModal();
*/
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button4Click(TObject *Sender)
{
/*
InputDocs->IForm=1;
InputDocs->Mode=4;
InputDocs->ShowModal();
*/
}
//---------------------------------------------------------------------------
void TForm1::InpTer()
{

Edit1->Text=InputDocs->TextBr;
Aspects->Edit();
Aspects->FieldByName("Вид территории")->Value=InputDocs->NumBr;
Aspects->Post();

}
//-------------------------------------------
void TForm1::InpDeyat()
{
/*
Edit2->Text=InputDocs->TextBr;
Aspects->Edit();
Aspects->FieldByName("Деятельность")->Value=InputDocs->NumBr;
Aspects->Post();
*/
}
//-------------------------------------------
void TForm1::InpAsp()
{
/*
Edit4->Text=InputDocs->TextBr;
Aspects->Edit();
Aspects->FieldByName("Аспект")->Value=InputDocs->NumBr;
Aspects->Post();
*/
}
//-------------------------------------------
void TForm1::InpVozd()
{
/*
Edit5->Text=InputDocs->TextBr;
Aspects->Edit();
Aspects->FieldByName("Воздействие")->Value=InputDocs->NumBr;
Aspects->Post();
*/
}
//-------------------------------------------
void TForm1::InpMeropr()
{
/*
DBMemo31->Clear();
DBMemo31->Lines->Add(InputDocs->TextBr);
Aspects->Edit();
Aspects->FieldByName("Предлагаемые мероприятия")->Value=InputDocs->TextBr;
Aspects->Post();
 */
}
//-------------------------------------------
void TForm1::Calc()
{

int G,O,R,S,T,L,N;
int Z,C,F,P;
Aspects->Active=true;
if (Aspects->RecordCount!=0)
{
 if(C1==true)
 {
  G=2;
 }
 else
 {
  G=0;
 }
 if(C5==true)
 {
  O=1;
 }
 else
 {
  O=0;
 }
 if(C2==true)
 {
  R=1;
 }
 else
 {
  R=0;
 }
 if(C6==true)
 {
  S=2;
 }
 else
 {
  S=0;
 }
 if(C3==true)
 {
  T=2;
 }
 else
 {
  T=0;
 }
 if(C4==true)
 {
  L=1;
 }
 else
 {
  L=0;
 }
 if(C7==true)
 {
  N=1;
 }
 else
 {
  N=0;
 }
 Tiagest->Active=false;
 Tiagest->CommandText="Select * From [Тяжесть последствий] ORDER By [Номер последствия]";
 Tiagest->Active=true;
 Tiagest->First();
 Tiagest->MoveBy(CPosl->ItemIndex);
 P=Tiagest->FieldByName("Балл")->Value;

 Posledstvie->Active=false;
 Posledstvie->CommandText="Select * From [Проявление воздействия] Order by [Номер проявления]";
 Posledstvie->Active=true;
 Posledstvie->First();
 Posledstvie->MoveBy(CProyav->ItemIndex);
 F=Posledstvie->FieldByName("Балл")->Value;

 C=G+O+R+S+T+L+N;
 if(C==0)
 {
  C=1;
 }
 Z=C*F*P;

int N=Aspects->FieldByName("Номер аспекта")->Value;

TempAspects->Active=false;
TempAspects->CommandText="Select * From Аспекты Where [Номер аспекта]="+IntToStr(N);;
TempAspects->Active=true;
TempAspects->Edit();
TempAspects->FieldByName("Z")->Value=Z;
if(C1==true)
{
TempAspects->FieldByName("G")->Value=1;
}
else
{
TempAspects->FieldByName("G")->Value=0;
}

if(C2==true)
{
TempAspects->FieldByName("R")->Value=1;
}
else
{
TempAspects->FieldByName("R")->Value=0;
}

if(C3==true)
{
TempAspects->FieldByName("T")->Value=1;
}
else
{
TempAspects->FieldByName("T")->Value=0;
}

if(C4==true)
{
TempAspects->FieldByName("L")->Value=1;
}
else
{
TempAspects->FieldByName("L")->Value=0;
}

if(C5==true)
{
TempAspects->FieldByName("O")->Value=1;
}
else
{
TempAspects->FieldByName("O")->Value=0;
}

if(C6==true)
{
TempAspects->FieldByName("S")->Value=1;
}
else
{
TempAspects->FieldByName("S")->Value=0;
}

if(C7==true)
{
TempAspects->FieldByName("N")->Value=1;
}
else
{
TempAspects->FieldByName("N")->Value=0;
}
TempAspects->FieldByName("Z")->Value=Z;
TempAspects->Post();
//DBEdit7->Text=Z;
Edit3->Text=IntToStr(Z);



//-----------Статус-бар---------------------


Znachimost->Active=false;
Znachimost->CommandText="Select * From Значимость Order By [Макс граница]";
Znachimost->Active=true;
AnsiString NZ;
AnsiString M;
int NumZn=-1;
bool Zn, ZZ=false;
for(int i=0;i<Znachimost->RecordCount;i++)
{
bool BB=Znachimost->FieldByName("Критерий")->Value;
if(BB==true & ZZ==false)
{
ZZ=true;
Zmax->Text=Znachimost->FieldByName("Мин граница")->Value;
}
if (Z>=Znachimost->FieldByName("Мин граница")->Value & Z<Znachimost->FieldByName("Макс граница")->Value)
{
NumZn=Znachimost->FieldByName("Номер значимости")->Value;
NZ=Znachimost->FieldByName("Наименование значимости")->Value;
M=Znachimost->FieldByName("Необходимая мера")->Value;
Zn=Znachimost->FieldByName("Критерий")->Value;

}
Znachimost->Next();
}
if(ZZ==false)
{
Znachimost->Last();
Zmax->Text="---";
}

if(NumZn==-1)
{
Znachimost->Last();
NumZn=Znachimost->FieldByName("Номер значимости")->Value;
NZ=Znachimost->FieldByName("Наименование значимости")->Value;
M=Znachimost->FieldByName("Необходимая мера")->Value;
Zn=Znachimost->FieldByName("Критерий")->Value;

}

if (Zn==true)
{
 CPrior->Enabled=true;
 Edit6->Font->Color=clRed;
 Edit6->Font->Style=TFontStyles()<< fsBold;
 Edit3->Font->Color=clRed;
 Edit3->Font->Style=TFontStyles()<< fsBold;
 Edit6->Text="Значимый";

}
else
{
 CPrior->Enabled=false;
 CPrior->ItemIndex=1;
  Edit6->Font->Color=clBlack;
 Edit6->Font->Style= TFontStyles();
 Edit3->Font->Color=clBlack;
 Edit3->Font->Style=TFontStyles();
 Edit6->Text="Незначимый";

}
int N1=Aspects->FieldByName("Номер аспекта")->Value;

TempAspects->Active=false;
TempAspects->CommandText="Select * From Аспекты Where [Номер аспекта]="+IntToStr(N1);
TempAspects->Active=true;


TempAspects->Edit();
TempAspects->FieldByName("Значимость")->Value=Zn;
TempAspects->Post();
//DBCheckBox1->Checked=Zn;


StatusBar1->Panels->Items[0]->Text=NZ;
StatusBar1->Panels->Items[1]->Text=IntToStr(Z);
StatusBar1->Panels->Items[2]->Text=M;

}
Aspects->UpdateBatch();
}
//---------------------------------------------------------------
void __fastcall TForm1::DBCheckBox1Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBCheckBox2Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBCheckBox3Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBCheckBox4Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBCheckBox5Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBCheckBox6Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBCheckBox7Click(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPoslChange(TObject *Sender)
{
 Calc();
}
//---------------------------------------------------------------------------


void __fastcall TForm1::CheckBox1Click(TObject *Sender)
{
C1=CheckBox1->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox2Click(TObject *Sender)
{
C2=CheckBox2->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox3Click(TObject *Sender)
{
C3=CheckBox3->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox4Click(TObject *Sender)
{
C4=CheckBox4->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox5Click(TObject *Sender)
{
C5=CheckBox5->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox6Click(TObject *Sender)
{
C6=CheckBox6->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CheckBox7Click(TObject *Sender)
{
C7=CheckBox7->Checked;
Calc();
Aspects->UpdateBatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker1Change(TObject *Sender)
{
//DBEdit9->Text=DateTimePicker1->Date;
Aspects->Edit();
Aspects->FieldByName("Дата создания")->Value=DateTimePicker1->Date;
Aspects->Post();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker2Change(TObject *Sender)
{
//DBEdit10->Text=DateTimePicker2->Date;
Aspects->Edit();
Aspects->FieldByName("Начало действия")->Value=DateTimePicker2->Date;
Aspects->Post();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker3Change(TObject *Sender)
{
//DBEdit11->Text=DateTimePicker3->Date;
Aspects->Edit();
Aspects->FieldByName("Конец действия")->Value=DateTimePicker3->Date;
Aspects->Post();
}
//---------------------------------------------------------------------------


void __fastcall TForm1::Button6Click(TObject *Sender)
{
/*
InputDocs->IForm=1;

InputDocs->Mode=5;
InputDocs->ShowModal();
*/
}
//---------------------------------------------------------------------------




void __fastcall TForm1::Button5Click(TObject *Sender)
{
//Filter->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::BitBtn7Click(TObject *Sender)
{
/*
Report1->Flt=Filtr1;
Report1->FltName=LFiltr->Caption;
Report1->PodrComText=Podrazdel->CommandText;

Report1->NumRep=1;
Report1->RepBase=Zast->ADOUsrAspect;
Report1->Role=Role;
Report1->ShowModal();
*/
}
//---------------------------------------------------------------------------

void __fastcall TForm1::FormHide(TObject *Sender)
{
/*
CurrentRecord=Aspects->RecNo;
Aspects->UpdateBatch();
*/
}
//---------------------------------------------------------------------------

//--------------------------------------------------
void __fastcall TForm1::DBMemo1Exit(TObject *Sender)
{
Aspects->UpdateBatch();
}
//----------------------------------

void __fastcall TForm1::N9Click(TObject *Sender)
{

/*
//Чтение аспектов
try
{
Zast->MClient->Start();

Zast->MClient->WriteDiaryEvent("NetAspects","Обновление аспектов (ручное)","");

if(LoadAspects(NumLogin))
{
SetAspects(Login);
//ShowMessage("Обновление справочников");
//LoadSpravs();
//ShowMessage("Начало");

SetCombo();

Aspects->Active=true;
if(Podrazdel->RecordCount!=0)
{

}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects","Для пользователя не назначено подразделений",Login);

ShowMessage("Для пользователя "+Login+" не назначено подразделений!\rЗавершение работы программы.");
Zast->Close();
}
InitCombo();
Zast->MClient->WriteDiaryEvent("NetAspects","Обновление аспектов (ручное) завершено","");
 ShowMessage("Завершено");
}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой обновления аспектов (ручное)","");
 ShowMessage("Ошибка загрузки аспектов");
}
Zast->MClient->Stop();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка обновления аспектов (ручное)"," Ошибка "+IntToStr(GetLastError()));
}
*/
}
//---------------------------------------------------------------------------
bool TForm1::LoadAspects(int NumLogin)
{
/*
bool Ret=true;
Zast->MClient->WriteDiaryEvent("NetAspects","Начало загрузки аспектов","");
try
{

MP<TADODataSet>LPodr(this);
LPodr->Connection=Zast->ADOUsrAspect;
LPodr->CommandText="SELECT Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Подразделения.ServerNum FROM Logins INNER JOIN (Подразделения INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+"));";
LPodr->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="DELETE Logins.AdmNum, Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+"));";
Comm->Execute();

MP<TADODataSet>LTemp(this);
LTemp->Connection=Zast->ADOUsrAspect;
LTemp->CommandText="SELECT TempAspects.[Номер аспекта], TempAspects.Подразделение, TempAspects.Ситуация, TempAspects.[Вид территории], TempAspects.Деятельность, TempAspects.Специальность, TempAspects.Аспект, TempAspects.Воздействие, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.Значимость, TempAspects.[Проявление воздействия], TempAspects.[Тяжесть последствий], TempAspects.Приоритетность,   TempAspects.[Дата создания], TempAspects.[Начало действия], TempAspects.[Конец действия] FROM TempAspects;";


Table* RTemp=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

RTemp->SetCommandText("SELECT TempAspects.[Номер аспекта], TempAspects.Подразделение, TempAspects.Ситуация, TempAspects.[Вид территории], TempAspects.Деятельность, TempAspects.Специальность, TempAspects.Аспект, TempAspects.Воздействие, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.Значимость, TempAspects.[Проявление воздействия], TempAspects.[Тяжесть последствий], TempAspects.Приоритетность,  TempAspects.[Дата создания], TempAspects.[Начало действия], TempAspects.[Конец действия] FROM TempAspects;");
//Remote->SetCommandText("Select Ситуации.[Номер ситуации], Ситуации.[Название ситуации] From Ситуации Order by [Номер ситуации]");




for(LPodr->First();!LPodr->Eof;LPodr->Next())
{
Comm->CommandText="Delete * From TempAspects";
Comm->Execute();

int PodrLoad=LPodr->FieldByName("ServerNum")->Value;
Zast->MClient->WriteDiaryEvent("NetAspects","Загрузка аспектов подразделения","Подразделение: "+LPodr->FieldByName("Название подразделения")->AsString);

Zast->MClient->PrepareLoadAspects(PodrLoad ,DBMemo1->Width, DBMemo31->Width, DBMemo2->Width, DBMemo4->Width);
LTemp->Active=false;
LTemp->Active=true;

RTemp->Active(false);
RTemp->Active(true);
FProg->Label1->Caption="Загрузка аспектов "+LPodr->FieldByName("Название подразделения")->AsString;
Zast->MClient->LoadTable(RTemp, LTemp, FProg->Label1, FProg->Progress);

if(Zast->MClient->VerifyTable(LTemp, RTemp)==0)
{
MP<TADODataSet>Memo1(this);
Memo1->Connection=Zast->ADOUsrAspect;
Memo1->CommandText="SELECT Memo1.Num, Memo1.NumStr, Memo1.Text FROM Memo1 ORDER BY Memo1.Num, Memo1.NumStr;";
Memo1->Active=true;

Table* LMemo1=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);
LMemo1->SetCommandText("SELECT Memo1.Num, Memo1.NumStr, Memo1.Text FROM Memo1 ORDER BY Memo1.Num, Memo1.NumStr;");
LMemo1->Active(true);

Zast->MClient->LoadTable(LMemo1, Memo1);
Memo1->Active=false;
Memo1->Active=true;
if(Zast->MClient->VerifyTable(LMemo1, Memo1)==0)
{

MP<TADODataSet>Memo2(this);
Memo2->Connection=Zast->ADOUsrAspect;
Memo2->CommandText="SELECT Memo2.Num, Memo2.NumStr, Memo2.Text FROM Memo2 ORDER BY Memo2.Num, Memo2.NumStr;";
Memo2->Active=true;

Table* LMemo2=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);
LMemo2->SetCommandText("SELECT Memo2.Num, Memo2.NumStr, Memo2.Text FROM Memo2 ORDER BY Memo2.Num, Memo2.NumStr;");
LMemo2->Active(true);

Zast->MClient->LoadTable(LMemo2, Memo2);

if(Zast->MClient->VerifyTable(LMemo2, Memo2)==0)
{

MP<TADODataSet>Memo3(this);
Memo3->Connection=Zast->ADOUsrAspect;
Memo3->CommandText="SELECT Memo3.Num, Memo3.NumStr, Memo3.Text FROM Memo3 ORDER BY Memo3.Num, Memo3.NumStr;";
Memo3->Active=true;

Table* LMemo3=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);
LMemo3->SetCommandText("SELECT Memo3.Num, Memo3.NumStr, Memo3.Text FROM Memo3 ORDER BY Memo3.Num, Memo3.NumStr;");
LMemo3->Active(true);

Zast->MClient->LoadTable(LMemo3, Memo3);

if(Zast->MClient->VerifyTable(LMemo3, Memo3)==0)
{

MP<TADODataSet>Memo4(this);
Memo4->Connection=Zast->ADOUsrAspect;
Memo4->CommandText="SELECT Memo4.Num, Memo4.NumStr, Memo4.Text FROM Memo4 ORDER BY Memo4.Num, Memo4.NumStr;";
Memo4->Active=true;

Table* LMemo4=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);
LMemo4->SetCommandText("SELECT Memo4.Num, Memo4.NumStr, Memo4.Text FROM Memo4 ORDER BY Memo4.Num, Memo4.NumStr;");
LMemo4->Active(true);

Zast->MClient->LoadTable(LMemo4, Memo4);

if(Zast->MClient->VerifyTable(LMemo4, Memo4)==0)
{
PrepareMergeAspects();
MergeAspects(NumLogin);
Zast->MClient->WriteDiaryEvent("NetAspects","Загрузка аспектов завершена","");

}
else
{
Ret=Ret & false;
break;
}
Zast->MClient->DeleteTable(this, LMemo4);

}
else
{
Ret=Ret & false;
break;
}
Zast->MClient->DeleteTable(this, LMemo3);

}
else
{
Ret=Ret & false;
break;
}
Zast->MClient->DeleteTable(this, LMemo2);

}
else
{
Ret=Ret & false;
break;
}
Zast->MClient->DeleteTable(this, LMemo1);
}
else
{
Ret=Ret & false;
break;
}
}
Zast->MClient->DeleteTable(this, RTemp);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой обновления аспектов","");

}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка обновления справочников (ручное)"," Ошибка "+IntToStr(GetLastError()));

}
return Ret;
*/


}
//------------------------------------------------------------------------
void TForm1::PrepareMergeAspects()
{
Zast->MClient->WriteDiaryEvent("NetAspects","Подготовка к загрузке аспектов","");
try
{

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="Select * From TempAspects order by [Номер аспекта]";
Temp->Active=true;

MP<TADODataSet>Memo(this);
Memo->Connection=Zast->ADOUsrAspect;


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
Zast->MClient->WriteDiaryEvent("NetAspects","Подготовка к загрузке аспектов завершена","");

}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка подготовки к загрузке аспектов"," Ошибка "+IntToStr(GetLastError()));

}
}
//------------------------------------------------------------------------
void  TForm1::MergeAspects(int NumLogin)
{
Zast->MClient->WriteDiaryEvent("NetAspects","Обновление аспектов","");
try
{
MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOUsrAspect;
Podr->CommandText="Select * From Подразделения";
Podr->Active=true;

MP<TADODataSet>TempAsp(this);
TempAsp->Connection=Zast->ADOUsrAspect;
TempAsp->CommandText="Select * From TempAspects";
TempAsp->Active=true;

for(TempAsp->First();!TempAsp->Eof;TempAsp->Next())
{
 int N=TempAsp->FieldByName("Подразделение")->Value;

 if(Podr->Locate("ServerNum",N,SO))
 {
  int Num=Podr->FieldByName("Номер подразделения")->Value;

  TempAsp->Edit();
  TempAsp->FieldByName("Подразделение")->Value=Num;
  TempAsp->Post();
 }
 else
 {
 Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой обновления аспектов","");
  ShowMessage("Ошибка объединения аспектов");
 }
}

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
String ST="INSERT INTO Аспекты ( Подразделение, Ситуация, [Вид территории], Деятельность, Специальность, Аспект, Воздействие, G, O, R, S, T, L, N, Z, Значимость, [Проявление воздействия], [Тяжесть последствий], Приоритетность, [Выполняющиеся мероприятия], [Предлагаемые мероприятия], [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Исполнитель, [Дата создания], [Начало действия], [Конец действия], ServerNum ) ";
ST=ST+" SELECT TempAspects.Подразделение, TempAspects.Ситуация, TempAspects.[Вид территории], TempAspects.Деятельность, TempAspects.Специальность, TempAspects.Аспект, TempAspects.Воздействие, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.Значимость, TempAspects.[Проявление воздействия], TempAspects.[Тяжесть последствий], TempAspects.Приоритетность, TempAspects.[Выполняющиеся мероприятия], TempAspects.[Предлагаемые мероприятия], TempAspects.[Мониторинг и контроль], TempAspects.[Предлагаемый мониторинг и контроль], TempAspects.Исполнитель, TempAspects.[Дата создания], TempAspects.[Начало действия], TempAspects.[Конец действия], TempAspects.[Номер аспекта] ";
ST=ST+" FROM TempAspects; ";
Comm->CommandText=ST;
Comm->Execute();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка обновления аспектов"," Ошибка "+IntToStr(GetLastError()));

}
}
//----------------------------------------------------------------------



void __fastcall TForm1::FormActivate(TObject *Sender)
{
/*
this->Top=0;
this->Left=0;
*/

this->Width=1024;
this->Height=742;

}
//---------------------------------------------------------------------------



void __fastcall TForm1::Cdjlysqhttcnhfcgtrnjd1Click(TObject *Sender)
{
/*
FSvod->Visible=true;
this->Visible=false;
*/
}
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------
void __fastcall TForm1::BitBtn9Click(TObject *Sender)
{
/*
Report1->Flt=Filtr2;
Report1->FltName=LFiltr->Caption;
Report1->PodrComText=Podrazdel->CommandText;

Report1->NumRep=2;
Report1->RepBase=Zast->ADOUsrAspect;
Report1->Role=Role;
Report1->ShowModal();
*/
}
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------

void __fastcall TForm1::DBMemo2Exit(TObject *Sender)
{
Aspects->UpdateBatch();        
}
//---------------------------------------------------------------------------


void __fastcall TForm1::DBMemo31Exit(TObject *Sender)
{
Aspects->UpdateBatch();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DBMemo4Exit(TObject *Sender)
{
Aspects->UpdateBatch();        
}
//---------------------------------------------------------------------------











void __fastcall TForm1::MetodikaN19Click(TObject *Sender)
{
// CreateNew();
//NameForm(this, "Оценка значимости аспектов");
Initialize();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N14Click(TObject *Sender)
{
//Connect();
//NameForm(this, "Оценка значимости аспектов");
Initialize();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N15Click(TObject *Sender)
{
//Rename();
//NameForm(this, "Оценка значимости аспектов");
Initialize();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N16Click(TObject *Sender)
{
//SaveAs();
//NameForm(this, "Оценка значимости аспектов");
Initialize();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N17Click(TObject *Sender)
{
//FullPatch();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CProyavKeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPoslKeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPriorKeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CPodrazdelKeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::CSitKeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker1KeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker2KeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker3KeyPress(TObject *Sender, char &Key)
{
Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N1Click(TObject *Sender)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_СТАРТ_ПОЛЬЗОВАТЕЛЬ");
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N18Click(TObject *Sender)
{
/*
if(Zast->MEcolog!=1)
{
ShowMessage("Доступ только с правами Администратора");
}
else
{
SZn->ShowModal();
}
*/
}
//---------------------------------------------------------------------------
bool TForm1::VerifySit()
{
/*
bool Ret=false;
try
{
MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOUsrAspect;
Local->CommandText="Select Ситуации.[Номер ситуации], Ситуации.[Название ситуации] From Ситуации  Where [Номер ситуации]<>0 Order by [Номер ситуации]";
Local->Active=true;

Table* Remote=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

Remote->SetCommandText("Select Ситуации.[Номер ситуации], Ситуации.[Название ситуации] From Ситуации Order by [Номер ситуации]");
Remote->Active(true);

Ret=Zast->MClient->VerifyTable(Local, Remote)!=0;
Zast->MClient->DeleteTable(this, Remote);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка проверки ситуаций"," Ошибка "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//--------------------------------------------
bool TForm1::VerifyTerr()
{
/*
bool Ret=false;
try
{
MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOUsrAspect;
Local->CommandText="Select Территория.[Номер территории], Территория.[Наименование территории] From Территория Order by [Номер территории]";
Local->Active=true;

Table* Remote=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

Remote->SetCommandText("Select Территория.[Номер территории], Территория.[Наименование территории] From Территория Order by [Номер территории]");
Remote->Active(true);

Ret=Zast->MClient->VerifyTable(Local, Remote)!=0;
Zast->MClient->DeleteTable(this, Remote);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка проверки территорий"," Ошибка "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//--------------------------------------------
bool TForm1::VerifyDeyat()
{
/*
bool Ret=false;
try
{
MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOUsrAspect;
Local->CommandText="Select Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности] From Деятельность Order by [Номер деятельности]";
Local->Active=true;

Table* Remote=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

Remote->SetCommandText("Select Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности] From Деятельность Order by [Номер деятельности]");
Remote->Active(true);

Ret=Zast->MClient->VerifyTable(Local, Remote)!=0;
Zast->MClient->DeleteTable(this, Remote);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка проверки деятельностей"," Ошибка "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//--------------------------------------------
bool TForm1::VerifyAsp()
{
/*
bool Ret=false;
try
{
MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOUsrAspect;
Local->CommandText="Select Аспект.[Номер аспекта], Аспект.[Наименование аспекта] From Аспект Order by [Номер аспекта]";
Local->Active=true;

Table* Remote=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

Remote->SetCommandText("Select Аспект.[Номер аспекта], Аспект.[Наименование аспекта] From Аспект Order by [Номер аспекта]");
Remote->Active(true);

Ret=Zast->MClient->VerifyTable(Local, Remote)!=0;
Zast->MClient->DeleteTable(this, Remote);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка проверки экоаспектов"," Ошибка "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//-------------------------------------------
bool TForm1::VerifyVozd()
{
/*
bool Ret=false;
try
{

MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOUsrAspect;
Local->CommandText="Select Воздействия.[Номер Воздействия], Воздействия.[Наименование Воздействия] From Воздействия Order by [Номер Воздействия]";
Local->Active=true;

Table* Remote=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

Remote->SetCommandText("Select Воздействия.[Номер Воздействия], Воздействия.[Наименование Воздействия] From Воздействия Order by [Номер Воздействия]");
Remote->Active(true);

Ret=Zast->MClient->VerifyTable(Local, Remote)!=0;
Zast->MClient->DeleteTable(this, Remote);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка проверки воздействий"," Ошибка "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//-------------------------------------------
bool TForm1::VerifyCryt()
{
/*
bool Ret=false;
try
{
MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOUsrAspect;
Local->CommandText="Select Значимость.[Номер значимости], Значимость.[Наименование значимости], Значимость.[Критерий], Значимость.[Критерий1], Значимость.[Мин граница], Значимость.[Макс граница], Значимость.[Необходимая мера] From Значимость Order by [Номер значимости]";
Local->Active=true;

Table* Remote=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

Remote->SetCommandText("Select Значимость.[Номер значимости], Значимость.[Наименование значимости], Значимость.[Критерий], Значимость.[Критерий1], Значимость.[Мин граница], Значимость.[Макс граница], Значимость.[Необходимая мера] From Значимость Order by [Номер значимости]");
Remote->Active(true);

Ret=Zast->MClient->VerifyTable(Local, Remote)!=0;
Zast->MClient->DeleteTable(this, Remote);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка проверки критериев"," Ошибка "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//--------------------------------------------
bool TForm1::LoadSit()
{
/*
bool Ret=false;
try
{
Zast->MClient->WriteDiaryEvent("NetAspects","Начало загрузки ситуаций","");
MP<TADODataSet>TempSit(this);
TempSit->Connection=Zast->ADOUsrAspect;
TempSit->CommandText="Select TempSit.[Номер ситуации], TempSit.[Название ситуации] From TempSit Order by [Номер ситуации]";

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempSit";
Comm->Execute();

Table* Sit=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

Sit->SetCommandText("Select Ситуации.[Номер ситуации], Ситуации.[Название ситуации] From Ситуации Where Показ=True  Order by [Номер ситуации]");
Sit->Active(true);
TempSit->Active=true;
Zast->MClient->LoadTable(Sit, TempSit);


if(Zast->MClient->VerifyTable(Sit, TempSit)==0)
{
MergeSit();
Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки ситуаций","");
Ret=true;
}
Zast->MClient->DeleteTable(this, Sit);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects","Сбой загрузки ситуаций","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка загрузки ситуаций"," Ошибка "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//-------------------------------------------
void TForm1::MergeSit()
{
/*
try
{
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="UPDATE Ситуации SET Ситуации.Del = True WHERE (((Ситуации.Показ)=True));";
Comm->Execute();

Zast->MClient->WriteDiaryEvent("NetAspects","Начало объединения ситуаций","");
MP<TADODataSet>Sit(this);
Sit->Connection=Zast->ADOUsrAspect;
Sit->CommandText="Select * From Ситуации";
Sit->Active=true;

MP<TADODataSet>TSit(this);
TSit->Connection=Zast->ADOUsrAspect;
TSit->CommandText="Select * From TempSit";
TSit->Active=true;



for(Sit->First();!Sit->Eof;Sit->Next())
{
 int N=Sit->FieldByName("Номер ситуации")->Value;
 if(TSit->Locate("Номер ситуации", N, SO))
 {
  Sit->Edit();
  Sit->FieldByName("Название ситуации")->Value=TSit->FieldByName("Название ситуации")->Value;
  Sit->FieldByName("Показ")->Value=true;
  Sit->FieldByName("Del")->Value=false;
  Sit->Post();

  TSit->Delete();
 }
}

Comm->CommandText="UPDATE Ситуации INNER JOIN Аспекты ON Ситуации.[Номер ситуации] = Аспекты.Ситуация SET Аспекты.Ситуация = 0 WHERE (((Ситуации.Del)=True));";
Comm->Execute();

Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * from Ситуации Where Del=true";
Comm->Execute();

Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="INSERT INTO Ситуации ( [Номер ситуации], [Название ситуации], Показ ) SELECT TempSit.[Номер ситуации], TempSit.[Название ситуации], True AS Показ FROM TempSit;";
Comm->Execute();

Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * from TempSit";
Comm->Execute();


}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка объединения ситуаций"," Ошибка "+IntToStr(GetLastError()));
}
*/
}
//-----------------------------------------
bool TForm1::LoadTerr()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","Начало загрузки территорий","");
try
{
MP<TADODataSet>TempTerr(this);
TempTerr->Connection=Zast->ADOUsrAspect;
TempTerr->CommandText="Select TempTerr.[Номер территории], TempTerr.[Наименование территории], TempTerr.[Показ] From TempTerr Order by [Номер территории]";

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempTerr";
Comm->Execute();

Table* Terr=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

Terr->SetCommandText("Select Территория.[Номер территории], Территория.[Наименование территории], Территория.[Показ] From Территория Order by [Номер территории]");
Terr->Active(true);
TempTerr->Active=true;
Zast->MClient->LoadTable(Terr, TempTerr);


if(Zast->MClient->VerifyTable(Terr, TempTerr)==0)
{
MergeTerr();
RefreshLocalReference("Узлы_5","Ветви_5");

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки территорий","");
Ret=true;
}
Zast->MClient->DeleteTable(this, Terr);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой загрузки территорий","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка загрузки территорий"," Ошибка "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//--------------------------------------------
void TForm1::MergeTerr()
{
/*
Zast->MClient->WriteDiaryEvent("NetAspects","Начало объединения территорий","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="UPDATE Территория SET Территория.Del = false;";
Comm->Execute();

MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
Tab->CommandText="select * from Территория";
Tab->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="Select * From TempTerr";
Temp->Active=true;

for(Tab->First();!Tab->Eof;Tab->Next())
{
 int Num=Tab->FieldByName("Номер территории")->Value;

 if(Temp->Locate("Номер территории",Num,SO))
 {
  Tab->Edit();
  Tab->FieldByName("Наименование территории")->Value=Temp->FieldByName("Наименование территории")->Value;
  Tab->FieldByName("Показ")->Value=Temp->FieldByName("Показ")->Value;
  Tab->Post();

  Temp->Delete();
 }
 else
 {
  Tab->Edit();
  Tab->FieldByName("Del")->Value=true;
  Tab->Post();
 }
}

Comm->CommandText="UPDATE Территория INNER JOIN Аспекты ON Территория.[Номер территории] = Аспекты.[Вид территории] SET Аспекты.[Вид территории] = 0 WHERE (((Территория.Del)=True));";
Comm->Execute();

Comm->CommandText="DELETE Территория.Del FROM Территория WHERE (((Территория.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO Территория ( [Номер территории], [Наименование территории], [Показ] ) SELECT TempTerr.[Номер территории], TempTerr.[Наименование территории], TempTerr.[Показ] FROM TempTerr;";
Comm->Execute();

Comm->CommandText="DELETE TempTerr.* FROM TempTerr;";
Comm->Execute();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка объединения территорий"," Ошибка "+IntToStr(GetLastError()));
}
*/
}
//---------------------------------------------

bool TForm1::LoadDeyat()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","Начало загрузки деятельностей","");
try
{
MP<TADODataSet>TempDeyat(this);
TempDeyat->Connection=Zast->ADOUsrAspect;
TempDeyat->CommandText="Select TempDeyat.[Номер деятельности], TempDeyat.[Наименование деятельности], TempDeyat.[Показ] From TempDeyat Order by [Номер деятельности]";

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempDeyat";
Comm->Execute();

Table* Deyat=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

Deyat->SetCommandText("Select Деятельность.[Номер деятельности], Деятельность.[Наименование деятельности], Деятельность.[Показ] From Деятельность Order by [Номер деятельности]");
Deyat->Active(true);
TempDeyat->Active=true;
Zast->MClient->LoadTable(Deyat, TempDeyat);


if(Zast->MClient->VerifyTable(Deyat, TempDeyat)==0)
{
MergeDeyat();
RefreshLocalReference("Узлы_6","Ветви_6");
Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки деятельностей","");

Ret=true;
}
Zast->MClient->DeleteTable(this, Deyat);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой загрузки деятельностей","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка загрузки деятельностей"," Ошибка "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//---------------------------------------------
void TForm1::MergeDeyat()
{
/*
Zast->MClient->WriteDiaryEvent("NetAspects","Начало объединения деятельностей","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="UPDATE Деятельность SET Деятельность.Del = false;";
Comm->Execute();

MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
Tab->CommandText="select * from Деятельность";
Tab->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="Select * From TempDeyat";
Temp->Active=true;

for(Tab->First();!Tab->Eof;Tab->Next())
{
 int Num=Tab->FieldByName("Номер деятельности")->Value;

 if(Temp->Locate("Номер деятельности",Num,SO))
 {
  Tab->Edit();
  Tab->FieldByName("Наименование деятельности")->Value=Temp->FieldByName("Наименование деятельности")->Value;
  Tab->FieldByName("Показ")->Value=Temp->FieldByName("Показ")->Value;
  Tab->Post();

  Temp->Delete();
 }
 else
 {
  Tab->Edit();
  Tab->FieldByName("Del")->Value=true;
  Tab->Post();
 }
}

Comm->CommandText="UPDATE Деятельность INNER JOIN Аспекты ON Деятельность.[Номер деятельности] = Аспекты.[деятельность] SET Аспекты.[Деятельность] = 0 WHERE (((Деятельность.Del)=True));";
Comm->Execute();

Comm->CommandText="DELETE Деятельность.Del FROM Деятельность WHERE (((Деятельность.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO Деятельность ( [Номер деятельности], [Наименование деятельности], Показ ) SELECT TempDeyat.[Номер Деятельности], TempDeyat.[Наименование деятельности], TempDeyat.Показ FROM TempDeyat;";
Comm->Execute();

Comm->CommandText="DELETE TempDeyat.* FROM TempDeyat;";
Comm->Execute();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка объединения деятельностей"," Ошибка "+IntToStr(GetLastError()));
}
*/
}
//---------------------------------------------

bool TForm1::LoadAsp()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","Начало загрузки экоаспектов","");
try
{
MP<TADODataSet>TempAsp(this);
TempAsp->Connection=Zast->ADOUsrAspect;
TempAsp->CommandText="Select TempAspect.[Номер аспекта], TempAspect.[Наименование аспекта], TempAspect.[Показ] From TempAspect Order by [Номер аспекта]";

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempAspect";
Comm->Execute();

Table* Asp=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

Asp->SetCommandText("Select Аспект.[Номер аспекта], Аспект.[Наименование аспекта], Аспект.[Показ] From Аспект Order by [Номер аспекта]");
Asp->Active(true);
TempAsp->Active=true;
Zast->MClient->LoadTable(Asp, TempAsp);


if(Zast->MClient->VerifyTable(Asp, TempAsp)==0)
{
MergeAsp();
RefreshLocalReference("Узлы_7","Ветви_7");

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки экоаспектов","");
Ret=true;
}
Zast->MClient->DeleteTable(this, Asp);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой загрузки экоаспектов","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка загрузки экоаспектов"," Ошибка "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//---------------------------------------------
void TForm1::MergeAsp()
{
/*
Zast->MClient->WriteDiaryEvent("NetAspects","Начало объединения экоаспектов","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="UPDATE Аспект SET Аспект.Del = false;";
Comm->Execute();

MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
Tab->CommandText="select * from Аспект";
Tab->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="Select * From TempAspect";
Temp->Active=true;

for(Tab->First();!Tab->Eof;Tab->Next())
{
 int Num=Tab->FieldByName("Номер аспекта")->Value;

 if(Temp->Locate("Номер аспекта",Num,SO))
 {
  Tab->Edit();
  Tab->FieldByName("Наименование аспекта")->Value=Temp->FieldByName("Наименование аспекта")->Value;
  Tab->FieldByName("Показ")->Value=Temp->FieldByName("Показ")->Value;
  Tab->Post();

  Temp->Delete();
 }
 else
 {
  Tab->Edit();
  Tab->FieldByName("Del")->Value=true;
  Tab->Post();
 }
}

Comm->CommandText="UPDATE Аспект INNER JOIN Аспекты ON Аспект.[Номер аспекта] = Аспекты.[Аспект] SET Аспекты.[Аспект] = 0 WHERE (((Аспект.Del)=True));";
Comm->Execute();

Comm->CommandText="DELETE Аспект.Del FROM Аспект WHERE (((Аспект.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO Аспект ( [Номер аспекта], [Наименование аспекта], Показ ) SELECT TempAspect.[Номер аспекта], TempAspect.[Наименование аспекта], TempAspect.[Показ] FROM TempAspect;";
Comm->Execute();

Comm->CommandText="DELETE TempAspect.* FROM TempAspect;";
Comm->Execute();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка объединения экоаспектов"," Ошибка "+IntToStr(GetLastError()));

}
*/
}
//----------------------------------------------

bool TForm1::LoadVozd()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","Начало загрузки воздействий","");
try
{
MP<TADODataSet>TempVozd(this);
TempVozd->Connection=Zast->ADOUsrAspect;
TempVozd->CommandText="Select TempVozd.[Номер воздействия], TempVozd.[Наименование воздействия], TempVozd.[Показ] From TempVozd Order by [Номер воздействия]";

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempVozd";
Comm->Execute();

Table* Vozd=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

Vozd->SetCommandText("Select Воздействия.[Номер воздействия], Воздействия.[Наименование воздействия], Воздействия.[Показ] From Воздействия Order by [Номер воздействия]");
Vozd->Active(true);
TempVozd->Active=true;
Zast->MClient->LoadTable(Vozd, TempVozd);


if(Zast->MClient->VerifyTable(Vozd, TempVozd)==0)
{
MergeVozd();

RefreshLocalReference("Узлы_3","Ветви_3");
Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки воздействий","");
Ret=true;
}
Zast->MClient->DeleteTable(this, Vozd);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой загрузки воздействий","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка загрузки воздействий"," Ошибка "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//---------------------------------------------
void TForm1::MergeVozd()
{
/*
Zast->MClient->WriteDiaryEvent("NetAspects","Начало объединения воздействий","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="UPDATE Воздействия SET Воздействия.Del = false;";
Comm->Execute();

MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOUsrAspect;
Tab->CommandText="select * from Воздействия";
Tab->Active=true;

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="Select * From TempVozd";
Temp->Active=true;

for(Tab->First();!Tab->Eof;Tab->Next())
{
 int Num=Tab->FieldByName("Номер воздействия")->Value;

 if(Temp->Locate("Номер воздействия",Num,SO))
 {
  Tab->Edit();
  Tab->FieldByName("Наименование воздействия")->Value=Temp->FieldByName("Наименование воздействия")->Value;
  Tab->FieldByName("Показ")->Value=Temp->FieldByName("Показ")->Value;
  Tab->Post();

  Temp->Delete();
 }
 else
 {
  Tab->Edit();
  Tab->FieldByName("Del")->Value=true;
  Tab->Post();
 }
}

Comm->CommandText="UPDATE Воздействия INNER JOIN Аспекты ON Воздействия.[Номер Воздействия] = Аспекты.[Воздействие] SET Аспекты.[Воздействие] = 0 WHERE (((Воздействия.Del)=True));";
Comm->Execute();

Comm->CommandText="DELETE Воздействия.Del FROM Воздействия WHERE (((Воздействия.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO Воздействия ( [Номер воздействия], [Наименование Воздействия], Показ ) SELECT TempVozd.[Номер воздействия], TempVozd.[Наименование Воздействия], TempVozd.[Показ] FROM TempVozd;";
Comm->Execute();

Comm->CommandText="DELETE TempVozd.* FROM TempVozd;";
Comm->Execute();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка объединения воздействий"," Ошибка "+IntToStr(GetLastError()));
}
*/
}
//---------------------------------------------
bool TForm1::LoadCryt()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","Начало загрузки критериев","");
try
{
MP<TADODataSet>TempZn(this);
TempZn->Connection=Zast->ADOUsrAspect;
TempZn->CommandText="Select TempZn.[Номер значимости], TempZn.[Наименование значимости], TempZn.[Критерий], TempZn.[Критерий1], TempZn.[Мин граница], TempZn.[Макс граница], TempZn.[Необходимая мера] From TempZn Order by [Номер значимости]";

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempZn";
Comm->Execute();

Table* Zn=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

Zn->SetCommandText("Select Значимость.[Номер значимости], Значимость.[Наименование значимости], Значимость.[Критерий], Значимость.[Критерий1], Значимость.[Мин граница], Значимость.[Макс граница], Значимость.[Необходимая мера] From Значимость Order by [Номер значимости]");
Zn->Active(true);

TempZn->Active=true;
Zast->MClient->LoadTable(Zn, TempZn);


if(Zast->MClient->VerifyTable(Zn, TempZn)==0)
{
MergeCryt();

Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки критериев","");
Ret=true;
}
Zast->MClient->DeleteTable(this, Zn);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой загрузки критериев","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка загрузки критериев"," Ошибка "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//---------------------------------------------
void TForm1::MergeCryt()
{
/*
Zast->MClient->WriteDiaryEvent("NetAspects","Начало объединения критериев","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="DELETE * FROM Значимость;";
Comm->Execute();

Comm->CommandText="INSERT INTO Значимость ( [Номер значимости], [Наименование значимости], Критерий, Критерий1, [Мин граница], [Макс граница], [Необходимая мера] ) SELECT TempZn.[Номер значимости], TempZn.[Наименование значимости], TempZn.Критерий, TempZn.Критерий1, TempZn.[Мин граница], TempZn.[Макс граница], TempZn.[Необходимая мера] FROM TempZn;";
Comm->Execute();
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка объединения критериев"," Ошибка "+IntToStr(GetLastError()));
}
*/
}
//---------------------------------------------
bool  TForm1::VerifyMeropr()
{
/*
bool Ret=false;
try
{
MP<TADODataSet>LNode(this);
LNode->Connection=Zast->ADOConn;
LNode->CommandText="Select Узлы_4.[Номер узла], Узлы_4.[Родитель], Узлы_4.[Название] From Узлы_4 Order by Узлы_4.[Номер узла], Узлы_4.[Родитель]";

Table* SNode=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Reference")].ServerDB);

SNode->SetCommandText("Select Узлы_4.[Номер узла], Узлы_4.[Родитель], Узлы_4.[Название] From Узлы_4 Order by Узлы_4.[Номер узла], Узлы_4.[Родитель]");
SNode->Active(true);


MP<TADODataSet>LBranch(this);
LBranch->Connection=Zast->ADOConn;
LBranch->CommandText="Select Узлы_4.[Номер узла], Узлы_4.[Родитель], Узлы_4.[Название] From Узлы_4 Order by Узлы_4.[Номер узла], Узлы_4.[Родитель]";

Table* SBranch=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Reference")].ServerDB);

SBranch->SetCommandText("Select Узлы_4.[Номер узла], Узлы_4.[Родитель], Узлы_4.[Название] From Узлы_4 Order by Узлы_4.[Номер узла], Узлы_4.[Родитель]");
SBranch->Active(true);

if(Zast->MClient->VerifyTable(SNode, LNode)!=0 | Zast->MClient->VerifyTable(SBranch, LBranch)!=0)
{
RefreshLocalReference("Узлы_4","Ветви_4");
Ret=true;
}
Zast->MClient->DeleteTable(this, SNode);
Zast->MClient->DeleteTable(this, SBranch);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка проверки мероприятий"," Ошибка "+IntToStr(GetLastError()));
}
return Ret;
*/
}
//----------------------------------------------
void TForm1::RefreshLocalReference(String NameNode, String NameBranch)
{
/*
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
Comm->CommandText="Delete * From TempNode";
Comm->Execute();

MP<TADODataSet>LNode(this);
LNode->Connection=Zast->ADOConn;
LNode->CommandText="Select TempNode.[Номер узла], TempNode.[Родитель], TempNode.[Название] From TempNode Order by TempNode.[Номер узла], TempNode.[Родитель]";
LNode->Active=true;

Table* SNode=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Reference")].ServerDB);

SNode->SetCommandText("Select "+NameNode+".[Номер узла], "+NameNode+".[Родитель], "+NameNode+".[Название] From "+NameNode+" Order by "+NameNode+".[Номер узла], "+NameNode+".[Родитель]");
SNode->Active(true);

Zast->MClient->LoadTable(SNode, LNode);


if(Zast->MClient->VerifyTable(SNode, LNode)==0)
{

Comm->CommandText="Delete * From TempBranch";
Comm->Execute();

MP<TADODataSet>LBranch(this);
LBranch->Connection=Zast->ADOConn;
LBranch->CommandText="Select TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.[Название], TempBranch.[Показ] From TempBranch Order by TempBranch.[Номер ветви], TempBranch.[Номер родителя]";
LBranch->Active=true;

Table* SBranch=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Reference")].ServerDB);

SBranch->SetCommandText("Select "+NameBranch+".[Номер ветви], "+NameBranch+".[Номер родителя], "+NameBranch+".[Название], "+NameBranch+".[Показ] From "+NameBranch+" Order by "+NameBranch+".[Номер ветви], "+NameBranch+".[Номер родителя]");
SBranch->Active(true);

Zast->MClient->LoadTable(SBranch, LBranch);


if(Zast->MClient->VerifyTable(SBranch, LBranch)==0)
{

Comm->CommandText="Delete * from "+NameNode;
Comm->Execute();

Comm->CommandText="INSERT INTO "+NameNode+" ( [Номер узла], Родитель, Название ) SELECT TempNode.[Номер узла], TempNode.Родитель, TempNode.Название FROM TempNode;";
Comm->Execute();

Comm->CommandText="INSERT INTO "+NameBranch+" ( [Номер ветви], [Номер родителя], Название, Показ ) SELECT TempBranch.[Номер ветви], TempBranch.[Номер родителя], TempBranch.Название, TempBranch.Показ FROM TempBranch;";
Comm->Execute();

}
Zast->MClient->DeleteTable(this, SBranch);

}
Zast->MClient->DeleteTable(this, SNode);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка локального обновления ссылок"," Ошибка "+IntToStr(GetLastError()));
}
*/
}
//---------------------------------------------

void __fastcall TForm1::Edit1DblClick(TObject *Sender)
{
Button1->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Edit2DblClick(TObject *Sender)
{
Button2->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Edit4DblClick(TObject *Sender)
{
Button3->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Edit5DblClick(TObject *Sender)
{
Button4->Click();        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::N10Click(TObject *Sender)
{
/*
Zast->MClient->Start();
Zast->MClient->WriteDiaryEvent("NetAspects","Начало обновления справочников (ручное)","");
try
{

if(LoadSpravs())
{
SetCombo();
InitCombo();
Zast->MClient->WriteDiaryEvent("NetAspects","Конец обновления справочников (ручное)","");
 ShowMessage("Завершено");
}
else
{
Zast->MClient->WriteDiaryEvent("Ошибка NetAspects","Сбой обновления справочников (ручное)","");

 ShowMessage("Ошибка загрузки справочников");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка обновления справочников (ручное)"," Ошибка "+IntToStr(GetLastError()));

}
Zast->MClient->Stop();
*/
}
//---------------------------------------------------------------------------




void __fastcall TForm1::N3Click(TObject *Sender)
{
/*
Zast->MClient->Start();
Zast->MClient->WriteDiaryEvent("NetAspects","Начало записи аспектов (ручное)","");
try
{
DataSetRefresh1->Execute();
DataSetPost1->Execute();
if(PrepareSaveAspects())
{

if(SaveAspects())
{
Zast->MClient->WriteDiaryEvent("NetAspects","Конец записи аспектов (ручное)","");
 ShowMessage("Завершено");
}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой записи аспектов (ручное)","");

 ShowMessage("Ошибка записи аспектов");
}
}
else
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Сбой подготовки к записи аспектов (ручное)","");

 ShowMessage("Ошибка подготовки записи аспектов");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка записи аспектов (ручное)"," Ошибка "+IntToStr(GetLastError()));

}
Zast->MClient->Stop();
*/
}
//---------------------------------------------------------------------------
bool  TForm1::PrepareSaveAspects()
{
/*
Zast->MClient->WriteDiaryEvent("NetAspects","Начало подготовки к записи аспектов","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From TempAspects";
Comm->Execute();

String CT="INSERT INTO TempAspects ( [Номер аспекта], Подразделение, Ситуация, [Вид территории], Деятельность, Специальность, Аспект, Воздействие, G, O, R, S, T, L, N, Z, Значимость, [Проявление воздействия], [Тяжесть последствий], Приоритетность, [Выполняющиеся мероприятия], [Предлагаемые мероприятия], [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Исполнитель, [Дата создания], [Начало действия], [Конец действия], ServerNum ) ";
CT=CT+" SELECT Аспекты.[Номер аспекта], Аспекты.Подразделение, Аспекты.Ситуация, Аспекты.[Вид территории], Аспекты.Деятельность, Аспекты.Специальность, Аспекты.Аспект, Аспекты.Воздействие, Аспекты.G, Аспекты.O, Аспекты.R, Аспекты.S, Аспекты.T, Аспекты.L, Аспекты.N, Аспекты.Z, Аспекты.Значимость, Аспекты.[Проявление воздействия], Аспекты.[Тяжесть последствий], Аспекты.Приоритетность, Аспекты.[Выполняющиеся мероприятия], Аспекты.[Предлагаемые мероприятия], Аспекты.[Мониторинг и контроль], Аспекты.[Предлагаемый мониторинг и контроль], Аспекты.Исполнитель, Аспекты.[Дата создания], Аспекты.[Начало действия], Аспекты.[Конец действия], Аспекты.[ServerNum] ";
CT=CT+" FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login ";
CT=CT+" WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+"));";

Comm->CommandText=CT;
Comm->Execute();

MP<TADODataSet>Temp(Form1);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="Select * From TempAspects";
Temp->Active=true;

for(Temp->First();!Temp->Eof;Temp->Next())
{
 int Num=Temp->FieldByName("Номер аспекта")->Value;

MP<TDataSource>DS(this);
DS->DataSet=Temp;
DS->Enabled=true;

MP<TDBMemo>TDBM(this);
TDBM->Visible=false;
TDBM->Parent=Zast;
TDBM->Width=DBMemo1->Width;
TDBM->DataSource=DS;
TDBM->DataField="Выполняющиеся мероприятия";

TStrings* TS=TDBM->Lines;
int N=TS->Count;

MP<TADODataSet>Memo1(this);
Memo1->Connection=Zast->ADOUsrAspect;
Memo1->CommandText="Select * From Memo1";
Memo1->Active=true;

for(int i=0;i<N;i++)
{
String S=TS->Strings[i];
Memo1->Insert();
Memo1->FieldByName("Num")->Value=Num;
Memo1->FieldByName("NumStr")->Value=i;
Memo1->FieldByName("Text")->Value=S;
Memo1->Post();
}

MP<TDBMemo>TDBM2(this);
TDBM2->Parent=Zast;
TDBM2->Visible=false;
TDBM2->Width=DBMemo31->Width;
TDBM2->DataSource=DS;
TDBM2->DataField="Предлагаемые мероприятия";

TStrings* TS2=TDBM2->Lines;
int N2=TS2->Count;

MP<TADODataSet>Memo2(this);
Memo2->Connection=Zast->ADOUsrAspect;
Memo2->CommandText="Select * From Memo2";
Memo2->Active=true;

for(int i=0;i<N2;i++)
{
String S=TS2->Strings[i];
Memo2->Insert();
Memo2->FieldByName("Num")->Value=Num;
Memo2->FieldByName("NumStr")->Value=i;
Memo2->FieldByName("Text")->Value=S;
Memo2->Post();
}

MP<TDBMemo>TDBM3(this);
TDBM3->Parent=Zast;
TDBM3->Visible=false;
TDBM3->Width=DBMemo2->Width;
TDBM3->DataSource=DS;
TDBM3->DataField="Мониторинг и контроль";

TStrings* TS3=TDBM3->Lines;
int N3=TS3->Count;

MP<TADODataSet>Memo3(this);
Memo3->Connection=Zast->ADOUsrAspect;
Memo3->CommandText="Select * From Memo3";
Memo3->Active=true;

for(int i=0;i<N3;i++)
{
String S=TS3->Strings[i];
Memo3->Insert();
Memo3->FieldByName("Num")->Value=Num;
Memo3->FieldByName("NumStr")->Value=i;
Memo3->FieldByName("Text")->Value=S;
Memo3->Post();
}

MP<TDBMemo>TDBM4(this);
TDBM4->Parent=Zast;
TDBM4->Visible=false;
TDBM4->Width=DBMemo4->Width;
TDBM4->DataSource=DS;
TDBM4->DataField="Предлагаемый мониторинг и контроль";

TStrings* TS4=TDBM4->Lines;
int N4=TS4->Count;

MP<TADODataSet>Memo4(this);
Memo4->Connection=Zast->ADOUsrAspect;
Memo4->CommandText="Select * From Memo4";
Memo4->Active=true;

for(int i=0;i<N4;i++)
{
String S=TS4->Strings[i];
Memo4->Insert();
Memo4->FieldByName("Num")->Value=Num;
Memo4->FieldByName("NumStr")->Value=i;
Memo4->FieldByName("Text")->Value=S;
Memo4->Post();
}
}
Zast->MClient->WriteDiaryEvent("NetAspects","Конец подготовки к записи аспектов","");

return true;
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка подготовки к записи аспектов"," Ошибка "+IntToStr(GetLastError()));

return false;
}
*/
}
//-----------------------------------------------------------
bool TForm1::SaveAspects()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","Начало записи аспектов","");
try
{
Zast->MClient->SetCommandText("Аспекты","Delete * From TempAspects");
Zast->MClient->CommandExec("Аспекты");

MP<TADODataSet>LTemp(this);
LTemp->Connection=Zast->ADOUsrAspect;
LTemp->CommandText="SELECT TempAspects.[Номер аспекта], TempAspects.Подразделение, TempAspects.Ситуация, TempAspects.[Вид территории], TempAspects.Деятельность, TempAspects.Специальность, TempAspects.Аспект, TempAspects.Воздействие, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.Значимость, TempAspects.[Проявление воздействия], TempAspects.[Тяжесть последствий], TempAspects.Приоритетность, TempAspects.[Дата создания], TempAspects.[Начало действия], TempAspects.[Конец действия], TempAspects.ServerNum FROM TempAspects ORDER BY TempAspects.[Номер аспекта];";
LTemp->Active=true;

Table* STemp=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

STemp->SetCommandText("SELECT TempAspects.[Номер аспекта], TempAspects.Подразделение, TempAspects.Ситуация, TempAspects.[Вид территории], TempAspects.Деятельность, TempAspects.Специальность, TempAspects.Аспект, TempAspects.Воздействие, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.Значимость, TempAspects.[Проявление воздействия], TempAspects.[Тяжесть последствий], TempAspects.Приоритетность, TempAspects.[Дата создания], TempAspects.[Начало действия], TempAspects.[Конец действия], TempAspects.ServerNum FROM TempAspects ORDER BY TempAspects.[Номер аспекта];");
STemp->Active(true);
FProg->Label1->Caption="Сохранение аспектов";
Zast->MClient->LoadTable(LTemp, STemp, FProg->Label1, FProg->Progress);

if(Zast->MClient->VerifyTable(LTemp, STemp)==0)
{

MP<TADODataSet>LMemo1(this);
LMemo1->Connection=Zast->ADOUsrAspect;
LMemo1->CommandText="SELECT Memo1.Num, Memo1.NumStr, Memo1.Text FROM Memo1 ORDER BY Memo1.Num;";
LMemo1->Active=true;

Table* SMemo1=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

SMemo1->SetCommandText("SELECT Memo1.Num, Memo1.NumStr, Memo1.Text FROM Memo1 ORDER BY Memo1.Num;");
SMemo1->Active(true);

Zast->MClient->LoadTable(LMemo1, SMemo1);

if(Zast->MClient->VerifyTable(LMemo1, SMemo1)==0)
{

MP<TADODataSet>LMemo2(this);
LMemo2->Connection=Zast->ADOUsrAspect;
LMemo2->CommandText="SELECT Memo2.Num, Memo2.NumStr, Memo2.Text FROM Memo2 ORDER BY Memo2.Num;";
LMemo2->Active=true;

Table* SMemo2=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

SMemo2->SetCommandText("SELECT Memo2.Num, Memo2.NumStr, Memo2.Text FROM Memo2 ORDER BY Memo2.Num;");
SMemo2->Active(true);

Zast->MClient->LoadTable(LMemo2, SMemo2);

if(Zast->MClient->VerifyTable(LMemo2, SMemo2)==0)
{

MP<TADODataSet>LMemo3(this);
LMemo3->Connection=Zast->ADOUsrAspect;
LMemo3->CommandText="SELECT Memo3.Num, Memo3.NumStr, Memo3.Text FROM Memo3 ORDER BY Memo3.Num;";
LMemo3->Active=true;

Table* SMemo3=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

SMemo3->SetCommandText("SELECT Memo3.Num, Memo3.NumStr, Memo3.Text FROM Memo3 ORDER BY Memo3.Num;");
SMemo3->Active(true);

Zast->MClient->LoadTable(LMemo3, SMemo3);

if(Zast->MClient->VerifyTable(LMemo3, SMemo3)==0)
{

MP<TADODataSet>LMemo4(this);
LMemo4->Connection=Zast->ADOUsrAspect;
LMemo4->CommandText="SELECT Memo4.Num, Memo4.NumStr, Memo4.Text FROM Memo4 ORDER BY Memo4.Num;";
LMemo4->Active=true;

Table* SMemo4=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

SMemo4->SetCommandText("SELECT Memo4.Num, Memo4.NumStr, Memo4.Text FROM Memo4 ORDER BY Memo4.Num;");
SMemo4->Active(true);

Zast->MClient->LoadTable(LMemo4, SMemo4);

if(Zast->MClient->VerifyTable(LMemo4, SMemo4)==0)
{
Zast->MClient->SetCommandText("Аспекты","Delete * From TempПодразделения");
Zast->MClient->CommandExec("Аспекты");

MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOUsrAspect;
Podr->CommandText="SELECT Подразделения.[Номер подразделения], Подразделения.[Название подразделения], Подразделения.ServerNum FROM Подразделения;";
Podr->Active=true;

Table* TempPodr=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

TempPodr->SetCommandText("SELECT TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения], TempПодразделения.ServerNum FROM TempПодразделения;");
TempPodr->Active(true);

Zast->MClient->LoadTable(Podr, TempPodr);

if(Zast->MClient->VerifyTable(Podr, TempPodr)==0)
{
Ret=true;
try
{
Zast->MClient->PrepareMergeAspects(NumLogin, DBMemo1->Width, DBMemo31->Width, DBMemo2->Width, DBMemo4->Width);

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * from TempAspects";
Comm->Execute();

MP<TADODataSet>Asp(this);
Asp->Connection=Zast->ADOUsrAspect;
Asp->CommandText="SELECT Аспекты.* FROM Logins INNER JOIN ((Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login WHERE (((Logins.ServerNum)="+IntToStr(NumLogin)+"));";
Asp->Active=true;

MP<TADODataSet>LT(this);
LT->Connection=Zast->ADOUsrAspect;
LT->CommandText="SELECT TempAspects.[Номер аспекта], TempAspects.ServerNum FROM TempAspects ORDER BY TempAspects.[Номер аспекта];";
LT->Active=true;

Table* ST=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

ST->SetCommandText("SELECT TempAspects.[Номер аспекта], TempAspects.ServerNum FROM TempAspects ORDER BY TempAspects.[Номер аспекта];");
ST->Active(true);

Zast->MClient->LoadTable(ST, LT);

if(Zast->MClient->VerifyTable(LT, ST)==0)
{
for(LT->First();!LT->Eof;LT->Next())
{
 int N=LT->FieldByName("Номер аспекта")->Value;

 if(Asp->Locate("Номер аспекта",N, SO))
 {
  Asp->Edit();
  Asp->FieldByName("ServerNum")->Value=LT->FieldByName("ServerNum")->Value;
  Asp->Post();
 }
 else
 {
 ShowMessage("Ошибка алгоритма");
 }
}


}
else
{
 ShowMessage("Ошибка передачи");
}
Zast->MClient->DeleteTable(this, ST);
}
catch(...)
{
return false;
}
}
Zast->MClient->DeleteTable(this, TempPodr);
}
Zast->MClient->DeleteTable(this, SMemo4);

}
Zast->MClient->DeleteTable(this, SMemo3);

}
Zast->MClient->DeleteTable(this, SMemo2);

}
Zast->MClient->DeleteTable(this, SMemo1);
}
Zast->MClient->DeleteTable(this, STemp);
Zast->MClient->WriteDiaryEvent("NetAspects","Конец записи аспектов","");
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка записи аспектов"," Ошибка "+IntToStr(GetLastError()));

}

return Ret;
*/
}
//------------------------------------------------------------
void __fastcall TForm1::N4Click(TObject *Sender)
{
/*
this->Hide();
Vvedenie->Show();
*/
}
//---------------------------------------------------------------------------


void __fastcall TForm1::N5Click(TObject *Sender)
{
/*
this->Hide();
Metodika->Show();
*/
}
//---------------------------------------------------------------------------

void __fastcall TForm1::WMSysCommand(TMessage & Msg)
{
  switch (Msg.WParam)
  {
/*
case SC_MINIMIZE:
{
 Shell_NotifyIcon(NIM_ADD, &NID);
  BCreate->Enabled = false;
 BDelete->Enabled = true;
 BHide->Enabled = true;
Form1->Visible=false;
//ShowMessage("SC_MINIMIZE");
                     break;
}
*/
case SC_MINIMIZE:
{

//  Shell_NotifyIcon(NIM_ADD, &NID);
//  BCreate->Enabled = false;
// BDelete->Enabled = true;
// BHide->Enabled = true;
//Form1->Visible=false;
//SendMessage(Handle,WM_SYSCOMMAND,SC_MINIMIZE,0);

//ShowMessage("SC_Close");
Application->Minimize();
break;
}

default:
DefWindowProc(Handle,Msg.Msg,Msg.WParam,Msg.LParam);

 }

}
//-------------------------------------------------------------------
bool TForm1::LoadPodr()
{
/*
bool Ret=false;
Zast->MClient->WriteDiaryEvent("NetAspects","Начало загрузки подразделений","");
try
{
MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOUsrAspect;
Temp->CommandText="SELECT TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения] FROM TempПодразделения order by [Номер подразделения];";
Temp->Active=true;

Table* SPodr=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

SPodr->SetCommandText("SELECT Подразделения.[Номер подразделения], Подразделения.[Название подразделения] FROM Подразделения order by [Номер подразделения];");
SPodr->Active(true);

Zast->MClient->LoadTable(Temp, SPodr);

if(Zast->MClient->VerifyTable(Temp, SPodr)==0)
{
MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOUsrAspect;
Podr->CommandText="SELECT * FROM Подразделения;";
Podr->Active=true;

for(Podr->First();!Podr->Eof;Podr->Next())
{
 int N=Podr->FieldByName("ServerNum")->Value;
 if(Temp->Locate("Номер подразделения", N, SO))
 {
  Podr->Edit();
  Podr->FieldByName("Название подразделения")->Value=Temp->FieldByName("Название подразделения")->Value;
  Podr->FieldByName("ServerNum")->Value=Temp->FieldByName("Номер подразделения")->Value;
  Podr->Post();

  Temp->Delete();
 }
 else
 {
  Podr->Edit();
  Podr->FieldByName("Del")->Value=true;
  Podr->Post();
 }
}

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * From Подразделения Whete Del=true";
Comm->Execute();

Comm->CommandText="INSERT INTO Подразделения ( ServerNum, [Название подразделения] ) SELECT TempПодразделения.[Номер подразделения], TempПодразделения.[Название подразделения] FROM TempПодразделения;";
Comm->Execute();

Comm->CommandText="Delete * From TempПодразделения";
Comm->Execute();
Zast->MClient->WriteDiaryEvent("NetAspects","Конец загрузки подразделений","");
Ret=true;
}
Zast->MClient->DeleteTable(this, SPodr);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("NetAspects","Сбой загрузки подразделений","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("NetAspects ошибка","Ошибка загрузки подразделений"," Ошибка "+IntToStr(GetLastError()));
return false;
}
return Ret;
*/
}
//--------------------------------------------------------------------
void __fastcall TForm1::N7Click(TObject *Sender)
{
//FAbout->ShowModal();
}
//---------------------------------------------------------------------------
void TForm1::EnabledForm(bool B)
{
Label2->Enabled=B;
CPodrazdel->Enabled=B;
Label3->Enabled=B;
CSit->Enabled=B;
Label4->Enabled=B;
Edit1->Enabled=B;
if (!B) Edit1->Text="";
Button1->Enabled=B;
Label5->Enabled=B;
Edit2->Enabled=B;
if (!B) Edit2->Text="";
Button2->Enabled=B;
Label6->Enabled=B;
DBEdit2->Enabled=B;
if (!B) DBEdit2->Text="";
Label7->Enabled=B;
Edit4->Enabled=B;
if (!B) Edit4->Text="";
Button3->Enabled=B;
Label8->Enabled=B;
Edit5->Enabled=B;
if (!B) Edit5->Text="";
Button4->Enabled=B;
Button5->Enabled=B;
LFiltr->Enabled=B;
Label9->Enabled=B;
Label10->Enabled=B;
DateTimePicker1->Enabled=B;
CheckBox1->Enabled=B;
CheckBox2->Enabled=B;
Label12->Enabled=B;
CheckBox3->Enabled=B;
CheckBox4->Enabled=B;
Label14->Enabled=B;
CheckBox5->Enabled=B;
Label11->Enabled=B;
CheckBox6->Enabled=B;
Label13->Enabled=B;
CheckBox7->Enabled=B;
Label15->Enabled=B;
Label16->Enabled=B;
CProyav->Enabled=B;
Label17->Enabled=B;
CPosl->Enabled=B;
Label18->Enabled=B;
Edit6->Enabled=B;
if (!B) Edit6->Text="";
Label19->Enabled=B;
Edit3->Enabled=B;
if (!B) Edit3->Text="";
Label26->Enabled=B;
Zmax->Enabled=B;
if (!B) Zmax->Text="";
Label20->Enabled=B;
Label21->Enabled=B;
CPrior->Enabled=B;
Label22->Enabled=B;
Label24->Enabled=B;
Label23->Enabled=B;
Label36->Enabled=B;
Label37->Enabled=B;
DBMemo1->Enabled=B;
Label25->Enabled=B;
DBMemo31->Enabled=B;
Button6->Enabled=B;
DBMemo2->Enabled=B;
DBMemo4->Enabled=B;
Label28->Enabled=B;
DateTimePicker2->Enabled=B;
Label29->Enabled=B;
DateTimePicker3->Enabled=B;
Label32->Enabled=B;
Label33->Enabled=B;
BitBtn7->Enabled=B;
Label34->Enabled=B;
Label35->Enabled=B;
BitBtn9->Enabled=B;

BitBtn1->Enabled=B;
BitBtn2->Enabled=B;
BitBtn3->Enabled=B;
BitBtn4->Enabled=B;
BitBtn6->Enabled=B;

StatusBar1->Visible=B;
}
//--------------------------------------------------------------------------
void __fastcall TForm1::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_ВВОД_РЕДАКТИРОВАНИЕ_И_УДАЛЕНИЕ_АСПЕКТОВ");
}
}
//---------------------------------------------------------------------------

