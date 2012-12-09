//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Main.h"
#include "MasterPointer.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::N1Click(TObject *Sender)
{
if(OpenDialog1->Execute())
{
 String Patch=OpenDialog1->FileName;
 String Name=ExtractFileName(Patch);

 try
 {
  Table->Insert();
  Table->FieldByName("Patch")->Value=Patch;
  Table->FieldByName("Result")->Value="Не обработан";
  Table->Post();


 }
 catch(...)
 {
  ShowMessage("Ошибка подключения к файлу\nВероятно ошибочный файл");
 }
}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormShow(TObject *Sender)
{
String Path=ExtractFilePath(Application->ExeName);
String DBPatch=Path+"Database\\DB.dtb";

Database->Connected=false;
Database->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+DBPatch+";Persist Security Info=False";
Database->LoginPrompt=false;
Database->Connected=true;

Table->Active=false;
Table->CommandText="Select * from Users Order by Num";
Table->Active=true;

for(Table->First();!Table->Eof;Table->Next())
{
 Table->Edit();
 Table->FieldByName("Result")->Value="Не обработан";
 Table->Post();
}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::N2Click(TObject *Sender)
{
Table->Delete();
}
//---------------------------------------------------------------------------


void __fastcall TForm1::Button1Click(TObject *Sender)
{
if(OpenDialog2->Execute())
{
 AspRefs->Text=OpenDialog2->FileName;
 Edit1->Text="Не обработан";
}        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button2Click(TObject *Sender)
{
MP<TADOCommand>Comm(this);
if(AspRefs->Text!="" & Table->RecordCount!=0)
{
//Копирование шаблонов
String Path=ExtractFilePath(Application->ExeName);
String PDiary=Path+"Templates\\Diary.dtb";
String PAspect=Path+"Templates\\Eco_Aspects.dtb";
String PRefs=Path+"Templates\\Reference.dtb";

String RDiary=Path+"Result\\Diary.dtb";
String RAspect=Path+"Result\\Eco_Aspects.dtb";
String RRefs=Path+"Result\\Reference.dtb";

DeleteFile(RDiary.c_str());
DeleteFile(RAspect.c_str());
DeleteFile(RRefs.c_str());

CopyFile(PDiary.c_str(), RDiary.c_str(), true);
CopyFile(PAspect.c_str(), RAspect.c_str(), true);
CopyFile(PRefs.c_str(), RRefs.c_str(), true);

try
{
//Копирование узлов и ветвей
AspectTo->Connected=false;
AspectTo->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+RAspect+";Persist Security Info=False";
AspectTo->LoginPrompt=false;
AspectTo->Connected=true;


TempFrom->Connected=false;
TempFrom->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+AspRefs->Text+";Persist Security Info=False";
TempFrom->LoginPrompt=false;
TempFrom->Connected=true;


TempTo->Connected=false;
TempTo->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+RRefs+";Persist Security Info=False";
TempTo->LoginPrompt=false;
TempTo->Connected=true;

CopyTable("Узлы_3","Ветви_3");
CopyTable("Узлы_4","Ветви_4");
CopyTable("Узлы_5","Ветви_5");
CopyTable("Узлы_6","Ветви_6");
CopyTable("Узлы_7","Ветви_7");

//Копирование методики
MP<TADODataSet>MetFrom(this);
MetFrom->Connection=TempFrom;
MetFrom->CommandText="select * from Методика";
MetFrom->Active=true;

MP<TADODataSet>MetTo(this);
MetTo->Connection=TempTo;
MetTo->CommandText="select * from Методика";
MetTo->Active=true;


Comm->Connection=TempTo;
Comm->CommandText="Delete * from Методика";
Comm->Execute();

MetTo->Insert();
MetTo->FieldByName("Номер")->Value=MetFrom->FieldByName("Номер")->Value;
MetTo->FieldByName("Методика")->Value=MetFrom->FieldByName("Методика")->Value;
MetTo->Post();

//Крпирование значимости
MP<TADODataSet>ZnFrom(this);
ZnFrom->Connection=TempFrom;
ZnFrom->CommandText="select * from Значимость";
ZnFrom->Active=true;

MP<TADODataSet>ZnTo(this);
ZnTo->Connection=TempTo;
ZnTo->CommandText="select * from Значимость";
ZnTo->Active=true;

MP<TADODataSet>ZnAsp(this);
ZnAsp->Connection=AspectTo;
ZnAsp->CommandText="select * from Значимость";
ZnAsp->Active=true;

Comm->Connection=TempTo;
Comm->CommandText="Delete * From Значимость";
Comm->Execute();

for(ZnFrom->First();!ZnFrom->Eof;ZnFrom->Next())
{
 ZnTo->Insert();
// ZnTo->FieldByName("Номер значимости")->Value=ZnFrom->FieldByName("Номер значимости")->Value;
 ZnTo->FieldByName("Наименование значимости")->Value=ZnFrom->FieldByName("Наименование значимости")->Value;
 ZnTo->FieldByName("Критерий1")->Value=ZnFrom->FieldByName("Критерий1")->Value;
 ZnTo->FieldByName("Критерий")->Value=ZnFrom->FieldByName("Критерий")->Value;
 ZnTo->FieldByName("Мин граница")->Value=ZnFrom->FieldByName("Мин граница")->Value;
 ZnTo->FieldByName("Макс граница")->Value=ZnFrom->FieldByName("Макс граница")->Value;
 ZnTo->FieldByName("Необходимая мера")->Value=ZnFrom->FieldByName("Необходимая мера")->Value;
 ZnTo->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From Значимость";
Comm->Execute();

for(ZnFrom->First();!ZnFrom->Eof;ZnFrom->Next())
{
 ZnAsp->Insert();
// ZnAsp->FieldByName("Номер значимости")->Value=ZnFrom->FieldByName("Номер значимости")->Value;
 ZnAsp->FieldByName("Номер значимости")->Value=ZnFrom->FieldByName("Номер значимости")->Value;
 ZnAsp->FieldByName("Наименование значимости")->Value=ZnFrom->FieldByName("Наименование значимости")->Value;
 ZnAsp->FieldByName("Критерий1")->Value=ZnFrom->FieldByName("Критерий1")->Value;
 ZnAsp->FieldByName("Критерий")->Value=ZnFrom->FieldByName("Критерий")->Value;
 ZnAsp->FieldByName("Мин граница")->Value=ZnFrom->FieldByName("Мин граница")->Value;
 ZnAsp->FieldByName("Макс граница")->Value=ZnFrom->FieldByName("Макс граница")->Value;
 ZnAsp->FieldByName("Необходимая мера")->Value=ZnFrom->FieldByName("Необходимая мера")->Value;
 ZnAsp->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From Аспект";
Comm->Execute();

MP<TADODataSet>AspectFrom(this);
AspectFrom->Connection=TempTo;
AspectFrom->CommandText="select * from Ветви_7";
AspectFrom->Active=true;

MP<TADODataSet>AspTo(this);
AspTo->Connection=AspectTo;
AspTo->CommandText="select * from Аспект";
AspTo->Active=true;


for(AspectFrom->First();!AspectFrom->Eof;AspectFrom->Next())
{
 AspTo->Insert();
 if(!AspectFrom->FieldByName("Показ")->AsBoolean)
 {
 AspTo->FieldByName("Номер аспекта")->Value=0; 
 }
 else
 {
 AspTo->FieldByName("Номер аспекта")->Value=AspectFrom->FieldByName("Номер ветви")->Value;
 }
 AspTo->FieldByName("Наименование аспекта")->Value=AspectFrom->FieldByName("Название")->Value;
 AspTo->FieldByName("Показ")->Value=AspectFrom->FieldByName("Показ")->Value;
 AspTo->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From Воздействия";
Comm->Execute();

MP<TADODataSet>VozdFrom(this);
VozdFrom->Connection=TempTo;
VozdFrom->CommandText="select * from Ветви_3";
VozdFrom->Active=true;

MP<TADODataSet>VozdTo(this);
VozdTo->Connection=AspectTo;
VozdTo->CommandText="select * from Воздействия";
VozdTo->Active=true;

for(VozdFrom->First();!VozdFrom->Eof;VozdFrom->Next())
{
 VozdTo->Insert();
 if(!VozdFrom->FieldByName("Показ")->AsBoolean)
 {
 VozdTo->FieldByName("Номер воздействия")->Value=0;
 }
 else
 {
 VozdTo->FieldByName("Номер воздействия")->Value=VozdFrom->FieldByName("Номер ветви")->Value;
 }
 VozdTo->FieldByName("Наименование воздействия")->Value=VozdFrom->FieldByName("Название")->Value;
 VozdTo->FieldByName("Показ")->Value=VozdFrom->FieldByName("Показ")->Value;
 VozdTo->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From Деятельность";
Comm->Execute();

MP<TADODataSet>DeyatFrom(this);
DeyatFrom->Connection=TempTo;
DeyatFrom->CommandText="select * from Ветви_6";
DeyatFrom->Active=true;

MP<TADODataSet>DeyatTo(this);
DeyatTo->Connection=AspectTo;
DeyatTo->CommandText="select * from Деятельность";
DeyatTo->Active=true;

for(DeyatFrom->First();!DeyatFrom->Eof;DeyatFrom->Next())
{
 DeyatTo->Insert();
 if(!DeyatFrom->FieldByName("Показ")->AsBoolean)
 {
  DeyatTo->FieldByName("Номер деятельности")->Value=0;
 }
 else
 {
 DeyatTo->FieldByName("Номер деятельности")->Value=DeyatFrom->FieldByName("Номер ветви")->Value;
 }
 DeyatTo->FieldByName("Наименование деятельности")->Value=DeyatFrom->FieldByName("Название")->Value;
 DeyatTo->FieldByName("Показ")->Value=DeyatFrom->FieldByName("Показ")->Value;
 DeyatTo->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From Значимость";
Comm->Execute();

MP<TADODataSet>ZFrom(this);
ZFrom->Connection=TempFrom;
ZFrom->CommandText="select * from Значимость";
ZFrom->Active=true;

MP<TADODataSet>ZTo(this);
ZTo->Connection=AspectTo;
ZTo->CommandText="select * from Значимость";
ZTo->Active=true;

for(ZFrom->First();!ZFrom->Eof;ZFrom->Next())
{
 ZTo->Insert();
 ZTo->FieldByName("Номер значимости")->Value=ZFrom->FieldByName("Номер значимости")->Value;
 ZTo->FieldByName("Наименование значимости")->Value=ZFrom->FieldByName("Наименование значимости")->Value;
 ZTo->FieldByName("Критерий")->Value=ZFrom->FieldByName("Критерий")->Value;
 ZTo->FieldByName("Критерий1")->Value=ZFrom->FieldByName("Критерий1")->Value;
 ZTo->FieldByName("Мин граница")->Value=ZFrom->FieldByName("Мин граница")->Value;
 ZTo->FieldByName("Макс граница")->Value=ZFrom->FieldByName("Макс граница")->Value;
 ZTo->FieldByName("Необходимая мера")->Value=ZFrom->FieldByName("Необходимая мера")->Value;
 ZTo->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From Территория";
Comm->Execute();

MP<TADODataSet>TerrFrom(this);
TerrFrom->Connection=TempTo;
TerrFrom->CommandText="select * from Ветви_5";
TerrFrom->Active=true;

MP<TADODataSet>TerrTo(this);
TerrTo->Connection=AspectTo;
TerrTo->CommandText="select * from Территория";
TerrTo->Active=true;

for(TerrFrom->First();!TerrFrom->Eof;TerrFrom->Next())
{
 TerrTo->Insert();
 if(!TerrFrom->FieldByName("Показ")->AsBoolean)
 {
 TerrTo->FieldByName("Номер территории")->Value=0;
 }
 else
 {
 TerrTo->FieldByName("Номер территории")->Value=TerrFrom->FieldByName("Номер ветви")->Value;
 }
 TerrTo->FieldByName("Наименование территории")->Value=TerrFrom->FieldByName("Название")->Value;
 TerrTo->FieldByName("Показ")->Value=TerrFrom->FieldByName("Показ")->Value;
 TerrTo->Post();
}


Edit1->Text="Обработан";

Table->First();
 String P1=Table->FieldByName("Patch")->Value;
MP<TADOConnection>AspConnect(this);
AspConnect->Connected=false;
AspConnect->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+P1+";Persist Security Info=False";
AspConnect->LoginPrompt=false;
AspConnect->Connected=true;

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From Ситуации Where показ=True";
Comm->Execute();

MP<TADODataSet>SitFrom(this);
SitFrom->Connection=AspConnect;
SitFrom->CommandText="select * from Ситуации";
SitFrom->Active=true;

MP<TADODataSet>SitTo(this);
SitTo->Connection=AspectTo;
SitTo->CommandText="select * from Ситуации";
SitTo->Active=true;

Comm->Connection=TempTo;
Comm->CommandText="Delete * From Ситуации";
Comm->Execute();

MP<TADODataSet>SitTo2(this);
SitTo2->Connection=TempTo;
SitTo2->CommandText="select * from Ситуации";
SitTo2->Active=true;

for(SitFrom->First();!SitFrom->Eof;SitFrom->Next())
{
 SitTo2->Insert();
 SitTo2->FieldByName("Номер ситуации")->Value=SitFrom->FieldByName("Номер ситуации")->AsInteger;
 SitTo2->FieldByName("Название ситуации")->Value=SitFrom->FieldByName("Название ситуации")->AsString;
 SitTo2->Post();

 SitTo->Insert();
 SitTo->FieldByName("Номер ситуации")->Value=SitFrom->FieldByName("Номер ситуации")->AsInteger;
 SitTo->FieldByName("Название ситуации")->Value=SitFrom->FieldByName("Название ситуации")->Value;
 SitTo->FieldByName("Показ")->Value=true;
 SitTo->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From Подразделения";
Comm->Execute();

for(Table->First();!Table->Eof;Table->Next())
{
try
{
 String P=Table->FieldByName("Patch")->Value;

AspConnect->Connected=false;
AspConnect->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+P+";Persist Security Info=False";
AspConnect->LoginPrompt=false;
AspConnect->Connected=true;

MP<TADODataSet>FromPodr(this);
FromPodr->Connection=AspConnect;
FromPodr->CommandText="Select * From Подразделения Order by [Номер подразделения]";
FromPodr->Active=true;



MP<TADODataSet>AspectsTo(this);
AspectsTo->Connection=AspectTo;
AspectsTo->CommandText="Select * from Аспекты";
AspectsTo->Active=true;




 SitTo2->Active=false;
 SitTo2->Active=true;
for(FromPodr->First();!FromPodr->Eof;FromPodr->Next())
{
 int NumPodrFrom=FromPodr->FieldByName("Номер подразделения")->Value;

 MP<TADODataSet>ToPodr(this);
 ToPodr->Connection=AspectTo;
 ToPodr->CommandText="Select * From Подразделения Order by [Номер подразделения]";
 ToPodr->Active=true;

 ToPodr->Insert();
 ToPodr->FieldByName("Название подразделения")->Value=FromPodr->FieldByName("Название подразделения")->Value;
 ToPodr->Post();

 ToPodr->Active=false;
 ToPodr->Active=true;
 ToPodr->Last();

 int NumPodrTo=ToPodr->FieldByName("Номер подразделения")->Value;

 MP<TADODataSet>AspectsFrom(this);
 AspectsFrom->Connection=AspConnect;
 AspectsFrom->CommandText="Select * from Аспекты Where Подразделение="+IntToStr(NumPodrFrom)+" Order by [Номер аспекта]";
 AspectsFrom->Active=true;
 int N;
 for(AspectsFrom->First();!AspectsFrom->Eof;AspectsFrom->Next())
 {

  N=0;
  AspectsTo->Insert();
  //AspectsTo->FieldByName("номер аспекта")->Value=AspectsFrom->FieldByName("Номер аспекта")->AsInteger;
  AspectsTo->FieldByName("Подразделение")->Value=NumPodrTo;
  N=AspectsFrom->FieldByName("Ситуация")->AsInteger;
  AspectsTo->FieldByName("Ситуация")->Value=N;

  N=EncodeBranch("Ветви_5", AspectsFrom->FieldByName("Вид территории")->AsInteger);
  AspectsTo->FieldByName("Вид территории")->Value=N;//AspectsFrom->FieldByName("Вид территории")->Value;
  N=EncodeBranch("Ветви_6", AspectsFrom->FieldByName("Деятельность")->AsInteger);
  AspectsTo->FieldByName("Деятельность")->Value=N;

  AspectsTo->FieldByName("Специальность")->Value=AspectsFrom->FieldByName("Специальность")->AsString+" ";

  //ShowMessage(AspectsFrom->FieldByName("Аспект")->Value);
  N=EncodeBranch("Ветви_7", AspectsFrom->FieldByName("Аспект")->AsInteger);
  AspectsTo->FieldByName("Аспект")->Value=N;
  N=EncodeBranch("Ветви_3", AspectsFrom->FieldByName("Воздействие")->AsInteger);
  AspectsTo->FieldByName("Воздействие")->Value=N;

  AspectsTo->FieldByName("G")->Value=AspectsFrom->FieldByName("G")->AsInteger;
  AspectsTo->FieldByName("O")->Value=AspectsFrom->FieldByName("O")->AsInteger;
  AspectsTo->FieldByName("R")->Value=AspectsFrom->FieldByName("R")->AsInteger;
  AspectsTo->FieldByName("S")->Value=AspectsFrom->FieldByName("S")->AsInteger;
  AspectsTo->FieldByName("T")->Value=AspectsFrom->FieldByName("T")->AsInteger;
  AspectsTo->FieldByName("L")->Value=AspectsFrom->FieldByName("L")->AsInteger;
  AspectsTo->FieldByName("N")->Value=AspectsFrom->FieldByName("N")->AsInteger;
  AspectsTo->FieldByName("Z")->Value=AspectsFrom->FieldByName("Z")->AsInteger;
  AspectsTo->FieldByName("Значимость")->Value=AspectsFrom->FieldByName("Значимость")->AsBoolean;

  N=AspectsFrom->FieldByName("Проявление воздействия")->AsInteger;
  AspectsTo->FieldByName("Проявление воздействия")->Value=N;
  N=AspectsFrom->FieldByName("Тяжесть последствий")->AsInteger;
  AspectsTo->FieldByName("Тяжесть последствий")->Value=N;
  N=AspectsFrom->FieldByName("Приоритетность")->AsInteger;
  AspectsTo->FieldByName("Приоритетность")->Value=N;

  AspectsTo->FieldByName("Выполняющиеся мероприятия")->Value=AspectsFrom->FieldByName("Выполняющиеся мероприятия")->Value;
  AspectsTo->FieldByName("Предлагаемые мероприятия")->Value=AspectsFrom->FieldByName("Предлагаемые мероприятия")->Value;
  AspectsTo->FieldByName("Мониторинг и контроль")->Value=AspectsFrom->FieldByName("Мониторинг и контроль")->Value;
  AspectsTo->FieldByName("Предлагаемый мониторинг и контроль")->Value=AspectsFrom->FieldByName("Предлагаемый мониторинг и контроль")->Value;
  AspectsTo->FieldByName("Исполнитель")->Value=AspectsFrom->FieldByName("Исполнитель")->Value;
  AspectsTo->FieldByName("Дата создания")->Value=AspectsFrom->FieldByName("Дата создания")->Value;
  AspectsTo->FieldByName("Начало действия")->Value=AspectsFrom->FieldByName("Начало действия")->Value;
  AspectsTo->FieldByName("Конец действия")->Value=AspectsFrom->FieldByName("Конец действия")->Value;

  AspectsTo->Post();
 }


}
Table->Edit();
Table->FieldByName("Result")->Value="Обработан";
Table->Post();
}
catch(...)
{
Table->Edit();
Table->FieldByName("Result")->Value="Ошибка";
Table->Post();
}

}
}
catch(...)
{
Edit1->Text="Ошибка";

}

//обработка Aspects

Comm->Connection=TempTo;
Comm->CommandText="UPDATE Ветви_3 SET Ветви_3.NumCopy = 0;";
Comm->Execute();

Comm->CommandText="UPDATE Ветви_4 SET Ветви_4.NumCopy = 0;";
Comm->Execute();

Comm->CommandText="UPDATE Ветви_5 SET Ветви_5.NumCopy = 0;";
Comm->Execute();

Comm->CommandText="UPDATE Ветви_6 SET Ветви_6.NumCopy = 0;";
Comm->Execute();

Comm->CommandText="UPDATE Ветви_7 SET Ветви_7.NumCopy = 0;";
Comm->Execute();

AspectTo->Connected=false;
TempFrom->Connected=false;
TempTo->Connected=false;


ShowMessage("Завершено");
}
else
{
 ShowMessage("Не указаны все файлы для преобразования");
}
}
//---------------------------------------------------------------------------
void TForm1::CopyTable(String NameNode, String NameBranch)
{

MP<TADODataSet>FromNode(this);
FromNode->Connection=TempFrom;
FromNode->CommandText="Select * From "+NameNode+" order by [Родитель], [Номер узла];";
FromNode->Active=true;

MP<TADODataSet>FromBranch(this);
FromBranch->Connection=TempFrom;
FromBranch->CommandText="Select * From "+NameBranch;
FromBranch->Active=true;

MP<TADODataSet>ToBranch(this);
ToBranch->Connection=TempTo;
ToBranch->CommandText="Select * From "+NameBranch;
ToBranch->Active=true;

MP<TADODataSet>ToNode(this);
ToNode->Connection=TempTo;
ToNode->CommandText="Select * From "+NameNode;
ToNode->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=TempTo;
Comm->CommandText="Delete * from "+NameNode;
Comm->Execute();

Comm->CommandText="Delete * from TempNode";
Comm->Execute();

Comm->CommandText="Delete * from TempBranch";
Comm->Execute();

MP<TADODataSet>TempNode(this);
TempNode->Connection=TempTo;
TempNode->CommandText="Select * From TempNode";
TempNode->Active=true;

MP<TADODataSet>TempBranch(this);
TempBranch->Connection=TempTo;
TempBranch->CommandText="Select * From TempBranch";
TempBranch->Active=true;

for(FromNode->First();!FromNode->Eof;FromNode->Next())
{
 TempNode->Insert();
 TempNode->FieldByName("Номер узла")->Value=FromNode->FieldByName("Номер узла")->Value;
 TempNode->FieldByName("Родитель")->Value=FromNode->FieldByName("Родитель")->Value;
 TempNode->FieldByName("Название")->Value=FromNode->FieldByName("Название")->Value;
 TempNode->Post();
}

for(TempBranch->First();!TempBranch->Eof;TempBranch->Next())
{
 TempBranch->Insert();
 TempBranch->FieldByName("Номер ветви")->Value=FromNode->FieldByName("Номер узла")->Value;
 TempBranch->FieldByName("Номер родителя")->Value=FromNode->FieldByName("Родитель")->Value;
 TempBranch->FieldByName("Название")->Value=FromNode->FieldByName("Название")->Value;
 TempBranch->FieldByName("Показ")->Value=FromNode->FieldByName("Показ")->Value;
 TempBranch->Post();
}

TempNode->First();
do
{
 int OldNumberNode=TempNode->FieldByName("Номер узла")->Value;

 ToNode->Insert();
 ToNode->FieldByName("Родитель")->Value=TempNode->FieldByName("Родитель")->Value;
 ToNode->FieldByName("Название")->Value=TempNode->FieldByName("Название")->Value;
 ToNode->Post();

 ToNode->Active=false;
 ToNode->Active=true;
 ToNode->Last();
 int NewNumberNode=ToNode->FieldByName("Номер узла")->Value;

 MP<TADODataSet>FromBranch(this);
 FromBranch->Connection=TempFrom;
 FromBranch->CommandText="Select * From "+NameBranch+" Where [Номер родителя]="+IntToStr(OldNumberNode);
 FromBranch->Active=true;

 for(FromBranch->First();!FromBranch->Eof;FromBranch->Next())
 {
  ToBranch->Insert();
  ToBranch->FieldByName("Номер родителя")->Value=NewNumberNode;
  ToBranch->FieldByName("Название")->Value=FromBranch->FieldByName("Название")->Value;
  ToBranch->FieldByName("Показ")->Value=FromBranch->FieldByName("Показ")->Value;
  ToBranch->FieldByName("NumCopy")->Value=FromBranch->FieldByName("Номер ветви")->Value;
  ToBranch->Post();
 }

 MP<TADODataSet>CorrectNode(this);
 CorrectNode->Connection=TempTo;
 CorrectNode->CommandText="Select * from TempNode Where Родитель="+IntToStr(OldNumberNode);
 CorrectNode->Active=true;

 for(CorrectNode->First();!CorrectNode->Eof;CorrectNode->Next())
 {
  CorrectNode->Edit();
  CorrectNode->FieldByName("Родитель")->Value=NewNumberNode;
  CorrectNode->Post();
 }

 Comm->CommandText="Delete * from TempNode Where [номер узла]="+IntToStr(OldNumberNode);
Comm->Execute();

 TempNode->Active=false;
 TempNode->Active=true;
}
while(TempNode->RecordCount!=0);

}
//--------------------------------------------------------------------------
int TForm1::EncodeBranch(String NameBranch, int NumCode)
{
MP<TADODataSet>Branch(this);
Branch->Connection=TempTo;
Branch->CommandText="Select * From "+NameBranch;
Branch->Active=true;

if(NumCode==0)
{
return 0;
}
else
{
Branch->Locate("NumCopy", NumCode, SO);
}
return Branch->FieldByName("Номер ветви")->AsInteger;
}
//-------------------------------------------------------------------------
