//---------------------------------------------------------------------------

#include <vcl.h>
#include "Reports.h"
#pragma hdrstop

#include "Reports.h"
#include "MasterPointer.h"
#include "Zastavka.h"

#include "MainForm.h"
#include "Rep1.h"
//---------------------------------------------------------------------------

#pragma package(smart_init)
void Reports::CreateReport1(int Podr, TDateTime Date1, TDateTime Date2, AnsiString Isp, String Filtr1, String LFiltr, bool OneList)
{
int Start=17;
int Number=0;
int i=0;
AnsiString T="Ф-001_1 ";
T=T+" Перечень "+IntToStr(Podr);
AnsiString PP1=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\Ф-001_1.xlt");
AnsiString PP2=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\"+T+".xlt");


Variant App =Variant::CreateObject("Excel.Application");
//Variant App1 =Variant::CreateObject("Excel.Application");
App.OlePropertySet("Visible",true);


Variant Book=App.OlePropertyGet("Workbooks").OleFunction("Add", PP1.c_str());
DeleteFile(PP2);

Variant Sheet=App.OlePropertyGet("ActiveSheet");

//Добавление листов
Variant Sheet1;

MP<TADODataSet>Report(Zast);
Report->Connection=Connect;

AnsiString G;


MP<TADODataSet>TempAspects(Zast);
TempAspects->Connection=Connect;

MP<TADODataSet>Podrazd(Zast);
Podrazd->Connection=this->Connect;

if(Podr==0)
{
if(OneList)
{
Podrazd->CommandText="Select * from TempПодразделения Order by [Название подразделения]";
}
else
{
Podrazd->CommandText="Select * from TempПодразделения Order by [Название подразделения]  DESC";
}
}
else
{
if(this->Role==4)
{
Podrazd->CommandText="Select * from TempПодразделения Where [Номер подразделения]="+IntToStr(Podr)+"  Order by [Название подразделения]  DESC";

}
else
{
Podrazd->CommandText="Select * from TempПодразделения Where ServerNum="+IntToStr(Podr)+"  Order by [Название подразделения]  DESC";
}
}
Podrazd->Active=true;
int ListCount=1;
int ListCount1=Podrazd->RecordCount;
for(Podrazd->First();!Podrazd->Eof;Podrazd->Next())
{
int NumPodrazd;
if(this->Role<4)
{
NumPodrazd=Podrazd->FieldByName("ServerNum")->AsInteger;
}
else
{
NumPodrazd=Podrazd->FieldByName("Номер подразделения")->AsInteger;
}

Report->Active=false;
if(Role==4)
{
if (Filtr1=="")
{
G="select TOP 2 * from TempAspects Where Подразделение="+IntToStr(NumPodrazd)+" Order By [Номер аспекта]";
}
else
{
G="SELECT TOP 2 TempAspects.*, Подразделения.ServerNum, Logins.ServerNum, Logins.Role FROM TempAspects INNER JOIN (Logins INNER JOIN (Подразделения INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login) ON TempAspects.Подразделение = Подразделения.[Номер подразделения] Where Подразделение="+IntToStr(NumPodrazd)+" AND "+Filtr1+" Order By [Номер аспекта]";
}
}
else
{
if (Filtr1=="")
{
G="select * from TempAspects Where Подразделение="+IntToStr(NumPodrazd)+" Order By [Номер аспекта]";
}
else
{
G="SELECT TempAspects.*, Logins.ServerNum, Подразделения.ServerNum FROM (Logins INNER JOIN (Подразделения INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login) INNER JOIN TempAspects ON Подразделения.ServerNum = TempAspects.Подразделение Where Подразделение="+IntToStr(NumPodrazd)+" AND "+Filtr1+" Order By [Номер аспекта]";
}
}
Report->CommandText=G;
Report->Active=true;

if(Report->RecordCount!=0)
{
if(!OneList)
{
Sheet1=Book.OlePropertyGet("Sheets",ListCount);
Sheet1.OleFunction("Copy",Book.OlePropertyGet("Sheets",1));
Sheet=Book.OlePropertyGet("Sheets",1);
//Sheet.OleProcedure("Select");

if(Podr!=0 | !OneList)
{
Sheet.OlePropertySet("Name",("Подразделение "+IntToStr(ListCount1)).c_str());
}
else
{
Sheet.OlePropertySet("Name","Все подразделения");
}
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value",Podrazd->FieldByName("Название подразделения")->AsString.c_str());
}
else
{
if(Podrazd->RecordCount>1)
{
Sheet.OlePropertySet("Name",("Все подразделения"));
}
else
{
Sheet.OlePropertySet("Name",("Подразделение "+IntToStr(ListCount1)).c_str());
}
}
//****************************
AnsiString NP;
NP=Podrazd->FieldByName("Название подразделения")->AsString;




TempAspects->Active=false;
TempAspects->CommandText="Select * From Подразделения Where [ServerNum]="+IntToStr(NumPodrazd);
TempAspects->Active=true;


AnsiString Text;
int Num;
T="Перечень экологических аспектов c "+Date1.DateString()+" по "+Date2.DateString();

T=" "+NP;
if(Podr==0 & OneList)
{
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value","Все подразделения");
}
else
{
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value",T.c_str());
}
T="Фильтр - "+LFiltr;
App.OlePropertyGet("Cells",12,1).OlePropertySet("Value",T.c_str());

//Report->First();


if(OneList & Podr==0)
{

App.OlePropertyGet("Range",(Address(Sheet,1,Start+i)+":"+Address(Sheet,17,Start+i)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Cells",Start+i,1).OlePropertySet("Value",Podrazd->FieldByName("Название подразделения")->AsString.c_str());
App.OlePropertyGet("Range",(Address(Sheet,1,Start+i)).c_str()).OlePropertySet("HorizontalAlignment",-4108);
i++;

}
else
{
i=0;
Number=0;
}

for(Report->First();!Report->Eof;Report->Next())
{
//Number=i;

App.OlePropertyGet("Cells",Start+i,1).OlePropertySet("Value",Number+1);
Num=Report->FieldByName("Деятельность")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From Деятельность Where [Номер деятельности]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("Наименование деятельности")->AsString+" ";
App.OlePropertyGet("Cells",Start+i,2).OlePropertySet("Value",Text.c_str());

Variant V=Report->FieldByName("Специальность")->Value;
if (V.IsNull()==true)
{
Text="";
}
else
{
Text=Report->FieldByName("Специальность")->Value;
}
App.OlePropertyGet("Cells",Start+i,3).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("Аспект")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From Аспект Where [Номер аспекта]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("Наименование аспекта")->Value;
App.OlePropertyGet("Cells",Start+i,4).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("Воздействие")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From Воздействия Where [Номер воздействия]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("Наименование воздействия")->Value;
App.OlePropertyGet("Cells",Start+i,5).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("Ситуация")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From Ситуации Where [Номер ситуации]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("Название ситуации")->AsString;
App.OlePropertyGet("Cells",Start+i,6).OlePropertySet("Value",Text.c_str());

if(Report->FieldByName("G")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,7).OlePropertySet("Value",2);
}
else
{
App.OlePropertyGet("Cells",Start+i,7).OlePropertySet("Value",0);
}

if(Report->FieldByName("O")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,8).OlePropertySet("Value",1);
}
else
{
App.OlePropertyGet("Cells",Start+i,8).OlePropertySet("Value",0);
}

if(Report->FieldByName("R")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,9).OlePropertySet("Value",1);
}
else
{
App.OlePropertyGet("Cells",Start+i,9).OlePropertySet("Value",0);
}

if(Report->FieldByName("S")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,10).OlePropertySet("Value",2);
}
else
{
App.OlePropertyGet("Cells",Start+i,10).OlePropertySet("Value",0);
}

if(Report->FieldByName("T")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,11).OlePropertySet("Value",2);
}
else
{
App.OlePropertyGet("Cells",Start+i,11).OlePropertySet("Value",0);
}

if(Report->FieldByName("L")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,12).OlePropertySet("Value",1);
}
else
{
App.OlePropertyGet("Cells",Start+i,12).OlePropertySet("Value",0);
}

if(Report->FieldByName("N")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,13).OlePropertySet("Value",1);
}
else
{
App.OlePropertyGet("Cells",Start+i,13).OlePropertySet("Value",0);
}


Num=Report->FieldByName("Тяжесть последствий")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From [Тяжесть последствий] Where [Номер последствия]="+IntToStr(Num);
TempAspects->Active=true;
Num=TempAspects->FieldByName("Балл")->Value;
App.OlePropertyGet("Cells",Start+i,14).OlePropertySet("Value",Num);

Num=Report->FieldByName("Проявление воздействия")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From [Проявление воздействия] Where [Номер проявления]="+IntToStr(Num);
TempAspects->Active=true;
Num=TempAspects->FieldByName("Балл")->Value;
App.OlePropertyGet("Cells",Start+i,15).OlePropertySet("Value",Num);

App.OlePropertyGet("Cells",Start+i,16).OlePropertySet("Value",Report->FieldByName("Z")->Value);


bool Zn=Report->FieldByName("Значимость")->Value;
if (Zn==true)
{

Text="Значимый";
}
else
{

Text="Незначимый";
}
App.OlePropertyGet("Cells",Start+i,17).OlePropertySet("Value",Text.c_str());


Number++;
i++;
}


//****************************

ListCount1--;
ListCount++;

App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertySet("VerticalAlignment",-4160);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertySet("WrapText",true);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertyGet("Borders",7).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertyGet("Borders",8).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertyGet("Borders",11).OlePropertySet("Weight",4);
if (Number>=1)
{
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertyGet("Borders",12).OlePropertySet("Weight",4);
}

if(!OneList)
{
T="Составил: _________________";
App.OlePropertyGet("Cells",Start+i+4,2).OlePropertySet("Value",T.c_str());
App.OlePropertyGet("Cells",Start+i+4,3).OlePropertySet("Value",Isp.c_str());
App.OlePropertyGet("Range",(Address(Sheet,3,Start+i+4)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+i+4)).c_str()).OlePropertySet("HorizontalAlignment",-4152);
App.OlePropertyGet("Cells",Start+i+4,4).OlePropertySet("Value","________________________");

App.OlePropertyGet("Cells",Start+i+5,2).OlePropertySet("Value","должность");
App.OlePropertyGet("Cells",Start+i+5,2).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,2,Start+i+5)).c_str()).OlePropertySet("IndentLevel",9);

App.OlePropertyGet("Cells",Start+i+5,3).OlePropertySet("Value","Ф.И.О.");
App.OlePropertyGet("Cells",Start+i+5,3).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,3,Start+i+5)).c_str()).OlePropertySet("IndentLevel",7);

App.OlePropertyGet("Cells",Start+i+5,4).OlePropertySet("Value","подпись");
App.OlePropertyGet("Cells",Start+i+5,4).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+i+5)).c_str()).OlePropertySet("IndentLevel",7);

AnsiString Patch=ExtractFilePath(Application->ExeName);
AnsiString NameReport=Patch+"\Reports\\Ф-001.1.xls";
}
}



}

if(!OneList)
{
//try
//{
App.OlePropertySet("DisplayAlerts",false);
Sheet1.OleFunction("Delete");
App.OlePropertySet("DisplayAlerts",true);
//}
//catch(...)
//{
//}
}
else
{
T="Составил: _________________";
App.OlePropertyGet("Cells",Start+i+4,2).OlePropertySet("Value",T.c_str());
App.OlePropertyGet("Cells",Start+i+4,3).OlePropertySet("Value",Isp.c_str());
App.OlePropertyGet("Range",(Address(Sheet,3,Start+i+4)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+i+4)).c_str()).OlePropertySet("HorizontalAlignment",-4152);
App.OlePropertyGet("Cells",Start+i+4,4).OlePropertySet("Value","________________________");

App.OlePropertyGet("Cells",Start+i+5,2).OlePropertySet("Value","должность");
App.OlePropertyGet("Cells",Start+i+5,2).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,2,Start+i+5)).c_str()).OlePropertySet("IndentLevel",9);

App.OlePropertyGet("Cells",Start+i+5,3).OlePropertySet("Value","Ф.И.О.");
App.OlePropertyGet("Cells",Start+i+5,3).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,3,Start+i+5)).c_str()).OlePropertySet("IndentLevel",7);

App.OlePropertyGet("Cells",Start+i+5,4).OlePropertySet("Value","подпись");
App.OlePropertyGet("Cells",Start+i+5,4).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+i+5)).c_str()).OlePropertySet("IndentLevel",7);

AnsiString Patch=ExtractFilePath(Application->ExeName);
AnsiString NameReport=Patch+"\Reports\\Ф-001.1.xls";
}

App.OlePropertySet("Visible",true);

App=NULL;
Book=NULL;
Sheet=NULL;
}
//------------------------------------------------------------------------------------------------------------------
void Reports::CreateReport2(int Podr, TDateTime Date1, TDateTime Date2, AnsiString Isp, String Filtr2, String LFiltr, bool OneList)
{
/*
int Start=17;
int Number=0;
int i=0;
AnsiString T="Ф-001_1 ";
T=T+" Перечень "+IntToStr(Podr);
AnsiString PP1=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\Ф-001_1.xlt");
AnsiString PP2=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\"+T+".xlt");


Variant App =Variant::CreateObject("Excel.Application");
//Variant App1 =Variant::CreateObject("Excel.Application");
App.OlePropertySet("Visible",true);


Variant Book=App.OlePropertyGet("Workbooks").OleFunction("Add", PP1.c_str());
DeleteFile(PP2);

Variant Sheet=App.OlePropertyGet("ActiveSheet");

//Добавление листов
Variant Sheet1;

MP<TADODataSet>Report(Zast);
Report->Connection=Connect;

AnsiString G;


MP<TADODataSet>TempAspects(Zast);
TempAspects->Connection=Connect;

MP<TADODataSet>Podrazd(Zast);
Podrazd->Connection=this->Connect;

if(Podr==0)
{
if(OneList)
{
Podrazd->CommandText="Select * from TempПодразделения Order by [Название подразделения]";
}
else
{
Podrazd->CommandText="Select * from TempПодразделения Order by [Название подразделения]  DESC";
}
}
else
{
if(this->Role==4)
{
Podrazd->CommandText="Select * from TempПодразделения Where [Номер подразделения]="+IntToStr(Podr)+"  Order by [Название подразделения]  DESC";

}
else
{
Podrazd->CommandText="Select * from TempПодразделения Where ServerNum="+IntToStr(Podr)+"  Order by [Название подразделения]  DESC";
}
}
Podrazd->Active=true;
int ListCount=1;
int ListCount1=Podrazd->RecordCount;
for(Podrazd->First();!Podrazd->Eof;Podrazd->Next())
{
int NumPodrazd;
if(this->Role<4)
{
NumPodrazd=Podrazd->FieldByName("ServerNum")->AsInteger;
}
else
{
NumPodrazd=Podrazd->FieldByName("Номер подразделения")->AsInteger;
}

Report->Active=false;
if(Role==4)
{
if (Filtr1=="")
{
G="select TOP 2 * from TempAspects Where Подразделение="+IntToStr(NumPodrazd)+" Order By [Номер аспекта]";
}
else
{
G="SELECT TOP 2 TempAspects.*, Подразделения.ServerNum, Logins.ServerNum, Logins.Role FROM TempAspects INNER JOIN (Logins INNER JOIN (Подразделения INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login) ON TempAspects.Подразделение = Подразделения.[Номер подразделения] Where Подразделение="+IntToStr(NumPodrazd)+" AND "+Filtr1+" Order By [Номер аспекта]";
}
}
else
{
if (Filtr1=="")
{
G="select * from TempAspects Where Подразделение="+IntToStr(NumPodrazd)+" Order By [Номер аспекта]";
}
else
{
G="SELECT TempAspects.*, Logins.ServerNum, Подразделения.ServerNum FROM (Logins INNER JOIN (Подразделения INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login) INNER JOIN TempAspects ON Подразделения.ServerNum = TempAspects.Подразделение Where Подразделение="+IntToStr(NumPodrazd)+" AND "+Filtr1+" Order By [Номер аспекта]";
}
}
Report->CommandText=G;
Report->Active=true;

if(Report->RecordCount!=0)
{
if(!OneList)
{
Sheet1=Book.OlePropertyGet("Sheets",ListCount);
Sheet1.OleFunction("Copy",Book.OlePropertyGet("Sheets",1));
Sheet=Book.OlePropertyGet("Sheets",1);
//Sheet.OleProcedure("Select");

if(Podr!=0 | !OneList)
{
Sheet.OlePropertySet("Name",("Подразделение "+IntToStr(ListCount1)).c_str());
}
else
{
Sheet.OlePropertySet("Name","Все подразделения");
}
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value",Podrazd->FieldByName("Название подразделения")->AsString.c_str());
}
else
{
if(Podrazd->RecordCount>1)
{
Sheet.OlePropertySet("Name",("Все подразделения"));
}
else
{
Sheet.OlePropertySet("Name",("Подразделение "+IntToStr(ListCount1)).c_str());
}
}
//****************************
AnsiString NP;
NP=Podrazd->FieldByName("Название подразделения")->AsString;




TempAspects->Active=false;
TempAspects->CommandText="Select * From Подразделения Where [ServerNum]="+IntToStr(NumPodrazd);
TempAspects->Active=true;


AnsiString Text;
int Num;
T="Перечень экологических аспектов c "+Date1.DateString()+" по "+Date2.DateString();

T=" "+NP;
if(Podr==0 & OneList)
{
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value","Все подразделения");
}
else
{
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value",T.c_str());
}
T="Фильтр - "+LFiltr;
App.OlePropertyGet("Cells",12,1).OlePropertySet("Value",T.c_str());

//Report->First();


if(OneList & Podr==0)
{

App.OlePropertyGet("Range",(Address(Sheet,1,Start+i)+":"+Address(Sheet,17,Start+i)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Cells",Start+i,1).OlePropertySet("Value",Podrazd->FieldByName("Название подразделения")->AsString.c_str());
App.OlePropertyGet("Range",(Address(Sheet,1,Start+i)).c_str()).OlePropertySet("HorizontalAlignment",-4108);
i++;

}
else
{
i=0;
Number=0;
}

for(Report->First();!Report->Eof;Report->Next())
{
//Number=i;

App.OlePropertyGet("Cells",Start+i,1).OlePropertySet("Value",Number+1);
Num=Report->FieldByName("Деятельность")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From Деятельность Where [Номер деятельности]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("Наименование деятельности")->AsString+" ";
App.OlePropertyGet("Cells",Start+i,2).OlePropertySet("Value",Text.c_str());

Variant V=Report->FieldByName("Специальность")->Value;
if (V.IsNull()==true)
{
Text="";
}
else
{
Text=Report->FieldByName("Специальность")->Value;
}
App.OlePropertyGet("Cells",Start+i,3).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("Аспект")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From Аспект Where [Номер аспекта]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("Наименование аспекта")->Value;
App.OlePropertyGet("Cells",Start+i,4).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("Воздействие")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From Воздействия Where [Номер воздействия]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("Наименование воздействия")->Value;
App.OlePropertyGet("Cells",Start+i,5).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("Ситуация")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From Ситуации Where [Номер ситуации]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("Название ситуации")->AsString;
App.OlePropertyGet("Cells",Start+i,6).OlePropertySet("Value",Text.c_str());

if(Report->FieldByName("G")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,7).OlePropertySet("Value",2);
}
else
{
App.OlePropertyGet("Cells",Start+i,7).OlePropertySet("Value",0);
}

if(Report->FieldByName("O")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,8).OlePropertySet("Value",1);
}
else
{
App.OlePropertyGet("Cells",Start+i,8).OlePropertySet("Value",0);
}

if(Report->FieldByName("R")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,9).OlePropertySet("Value",1);
}
else
{
App.OlePropertyGet("Cells",Start+i,9).OlePropertySet("Value",0);
}

if(Report->FieldByName("S")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,10).OlePropertySet("Value",2);
}
else
{
App.OlePropertyGet("Cells",Start+i,10).OlePropertySet("Value",0);
}

if(Report->FieldByName("T")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,11).OlePropertySet("Value",2);
}
else
{
App.OlePropertyGet("Cells",Start+i,11).OlePropertySet("Value",0);
}

if(Report->FieldByName("L")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,12).OlePropertySet("Value",1);
}
else
{
App.OlePropertyGet("Cells",Start+i,12).OlePropertySet("Value",0);
}

if(Report->FieldByName("N")->Value==1)
{
App.OlePropertyGet("Cells",Start+i,13).OlePropertySet("Value",1);
}
else
{
App.OlePropertyGet("Cells",Start+i,13).OlePropertySet("Value",0);
}


Num=Report->FieldByName("Тяжесть последствий")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From [Тяжесть последствий] Where [Номер последствия]="+IntToStr(Num);
TempAspects->Active=true;
Num=TempAspects->FieldByName("Балл")->Value;
App.OlePropertyGet("Cells",Start+i,14).OlePropertySet("Value",Num);

Num=Report->FieldByName("Проявление воздействия")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From [Проявление воздействия] Where [Номер проявления]="+IntToStr(Num);
TempAspects->Active=true;
Num=TempAspects->FieldByName("Балл")->Value;
App.OlePropertyGet("Cells",Start+i,15).OlePropertySet("Value",Num);

App.OlePropertyGet("Cells",Start+i,16).OlePropertySet("Value",Report->FieldByName("Z")->Value);


bool Zn=Report->FieldByName("Значимость")->Value;
if (Zn==true)
{

Text="Значимый";
}
else
{

Text="Незначимый";
}
App.OlePropertyGet("Cells",Start+i,17).OlePropertySet("Value",Text.c_str());


Number++;
i++;
}


//****************************

ListCount1--;
ListCount++;

App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertySet("VerticalAlignment",-4160);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertySet("WrapText",true);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertyGet("Borders",7).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertyGet("Borders",8).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertyGet("Borders",11).OlePropertySet("Weight",4);
if (Number>=1)
{
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,17,Start+i-1)).c_str()).OlePropertyGet("Borders",12).OlePropertySet("Weight",4);
}

if(!OneList)
{
T="Составил: _________________";
App.OlePropertyGet("Cells",Start+i+4,2).OlePropertySet("Value",T.c_str());
App.OlePropertyGet("Cells",Start+i+4,3).OlePropertySet("Value",Isp.c_str());
App.OlePropertyGet("Range",(Address(Sheet,3,Start+i+4)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+i+4)).c_str()).OlePropertySet("HorizontalAlignment",-4152);
App.OlePropertyGet("Cells",Start+i+4,4).OlePropertySet("Value","________________________");

App.OlePropertyGet("Cells",Start+i+5,2).OlePropertySet("Value","должность");
App.OlePropertyGet("Cells",Start+i+5,2).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,2,Start+i+5)).c_str()).OlePropertySet("IndentLevel",9);

App.OlePropertyGet("Cells",Start+i+5,3).OlePropertySet("Value","Ф.И.О.");
App.OlePropertyGet("Cells",Start+i+5,3).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,3,Start+i+5)).c_str()).OlePropertySet("IndentLevel",7);

App.OlePropertyGet("Cells",Start+i+5,4).OlePropertySet("Value","подпись");
App.OlePropertyGet("Cells",Start+i+5,4).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+i+5)).c_str()).OlePropertySet("IndentLevel",7);

AnsiString Patch=ExtractFilePath(Application->ExeName);
AnsiString NameReport=Patch+"\Reports\\Ф-001.1.xls";
}
}



}

if(!OneList)
{
//try
//{
App.OlePropertySet("DisplayAlerts",false);
Sheet1.OleFunction("Delete");
App.OlePropertySet("DisplayAlerts",true);
//}
//catch(...)
//{
//}
}
else
{
T="Составил: _________________";
App.OlePropertyGet("Cells",Start+i+4,2).OlePropertySet("Value",T.c_str());
App.OlePropertyGet("Cells",Start+i+4,3).OlePropertySet("Value",Isp.c_str());
App.OlePropertyGet("Range",(Address(Sheet,3,Start+i+4)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+i+4)).c_str()).OlePropertySet("HorizontalAlignment",-4152);
App.OlePropertyGet("Cells",Start+i+4,4).OlePropertySet("Value","________________________");

App.OlePropertyGet("Cells",Start+i+5,2).OlePropertySet("Value","должность");
App.OlePropertyGet("Cells",Start+i+5,2).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,2,Start+i+5)).c_str()).OlePropertySet("IndentLevel",9);

App.OlePropertyGet("Cells",Start+i+5,3).OlePropertySet("Value","Ф.И.О.");
App.OlePropertyGet("Cells",Start+i+5,3).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,3,Start+i+5)).c_str()).OlePropertySet("IndentLevel",7);

App.OlePropertyGet("Cells",Start+i+5,4).OlePropertySet("Value","подпись");
App.OlePropertyGet("Cells",Start+i+5,4).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+i+5)).c_str()).OlePropertySet("IndentLevel",7);

AnsiString Patch=ExtractFilePath(Application->ExeName);
AnsiString NameReport=Patch+"\Reports\\Ф-001.1.xls";
}

App.OlePropertySet("Visible",true);

App=NULL;
Book=NULL;
Sheet=NULL;
*/
int Start=16;

int Number=0;
int i=0;
AnsiString T="Ф-001_2 ";
T=T+" Перечень "+IntToStr(Podr);
AnsiString PP1=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\Ф-001_2.xlt");
AnsiString PP2=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\"+T+".xlt");


Variant App =Variant::CreateObject("Excel.Application");
//Variant App1 =Variant::CreateObject("Excel.Application");
App.OlePropertySet("Visible",true);


Variant Book=App.OlePropertyGet("Workbooks").OleFunction("Add", PP1.c_str());
DeleteFile(PP2);

Variant Sheet=App.OlePropertyGet("ActiveSheet");

//Добавление листов
Variant Sheet1;

MP<TADODataSet>Report(Zast);
Report->Connection=Connect;

AnsiString G;


MP<TADODataSet>TempAspects(Zast);
TempAspects->Connection=Connect;

MP<TADODataSet>Podrazd(Zast);
Podrazd->Connection=this->Connect;

if(Podr==0)
{
if(OneList)
{
Podrazd->CommandText="Select * from TempПодразделения Order by [Название подразделения]";
}
else
{
Podrazd->CommandText="Select * from TempПодразделения Order by [Название подразделения]  DESC";
}
}
else
{
if(this->Role==4)
{
Podrazd->CommandText="Select * from TempПодразделения Where [Номер подразделения]="+IntToStr(Podr)+"  Order by [Название подразделения]  DESC";

}
else
{
Podrazd->CommandText="Select * from TempПодразделения Where ServerNum="+IntToStr(Podr)+"  Order by [Название подразделения]  DESC";
}
}
Podrazd->Active=true;
int ListCount=1;
int ListCount1=Podrazd->RecordCount;
for(Podrazd->First();!Podrazd->Eof;Podrazd->Next())
{
int NumPodrazd;
if(this->Role<4)
{
NumPodrazd=Podrazd->FieldByName("ServerNum")->AsInteger;
}
else
{
NumPodrazd=Podrazd->FieldByName("Номер подразделения")->AsInteger;
}

Report->Active=false;
if(Role==4)
{
if (Filtr2=="")
{
G="select TOP 2 * from TempAspects Where Подразделение="+IntToStr(NumPodrazd)+" Order By [Номер аспекта]";
}
else
{
G="SELECT TOP 2 TempAspects.*, Подразделения.ServerNum, Logins.ServerNum, Logins.Role FROM TempAspects INNER JOIN (Logins INNER JOIN (Подразделения INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login) ON TempAspects.Подразделение = Подразделения.[Номер подразделения] Where Подразделение="+IntToStr(NumPodrazd)+" AND "+Filtr2+" Order By [Номер аспекта]";
}
}
else
{
if (Filtr2=="")
{
G="select * from TempAspects Where Подразделение="+IntToStr(NumPodrazd)+" Order By [Номер аспекта]";
}
else
{
G="SELECT TempAspects.*, Logins.ServerNum, Подразделения.ServerNum FROM (Logins INNER JOIN (Подразделения INNER JOIN ObslOtdel ON Подразделения.[Номер подразделения] = ObslOtdel.NumObslOtdel) ON Logins.Num = ObslOtdel.Login) INNER JOIN TempAspects ON Подразделения.ServerNum = TempAspects.Подразделение Where Подразделение="+IntToStr(NumPodrazd)+" AND "+Filtr2+" Order By [Номер аспекта]";
}
}
Report->CommandText=G;
Report->Active=true;

if(Report->RecordCount!=0)
{
if(!OneList)
{
Sheet1=Book.OlePropertyGet("Sheets",ListCount);
Sheet1.OleFunction("Copy",Book.OlePropertyGet("Sheets",1));
Sheet=Book.OlePropertyGet("Sheets",1);
//Sheet.OleProcedure("Select");

if(Podr!=0 | !OneList)
{
Sheet.OlePropertySet("Name",("Подразделение "+IntToStr(ListCount1)).c_str());
}
else
{
Sheet.OlePropertySet("Name","Все подразделения");
}
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value",Podrazd->FieldByName("Название подразделения")->AsString.c_str());
}
else
{
if(Podrazd->RecordCount>1)
{
Sheet.OlePropertySet("Name",("Все подразделения"));
}
else
{
Sheet.OlePropertySet("Name",("Подразделение "+IntToStr(ListCount1)).c_str());
}
}
//****************************
AnsiString NP;
NP=Podrazd->FieldByName("Название подразделения")->AsString;




TempAspects->Active=false;
TempAspects->CommandText="Select * From Подразделения Where [ServerNum]="+IntToStr(NumPodrazd);
TempAspects->Active=true;


AnsiString Text;
int Num;
T="Перечень экологических аспектов c "+Date1.DateString()+" по "+Date2.DateString();

T=" "+NP;
if(Podr==0 & OneList)
{
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value","Все подразделения");
}
else
{
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value",T.c_str());
}
T="Фильтр - "+LFiltr;
App.OlePropertyGet("Cells",12,1).OlePropertySet("Value",T.c_str());

//Report->First();


if(OneList & Podr==0)
{

App.OlePropertyGet("Range",(Address(Sheet,1,Start+i)+":"+Address(Sheet,17,Start+i)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Cells",Start+i,1).OlePropertySet("Value",Podrazd->FieldByName("Название подразделения")->AsString.c_str());
App.OlePropertyGet("Range",(Address(Sheet,1,Start+i)).c_str()).OlePropertySet("HorizontalAlignment",-4108);
i++;

}
else
{
i=0;
Number=0;
}


/*

AnsiString G;
if (Filtr2=="")
{
G="select * from TempAspects  Order By [Номер аспекта]";
}
else
{
G="select * from TempAspects Where "+Filtr2+" Order By [Номер аспекта]";
}
MP<TADODataSet>Report(Zast);
Report->Connection=Connect;

MP<TADODataSet>TempAspects(Zast);
TempAspects->Connection=Connect;

Report->Active=false;

Report->CommandText=G;
Report->Active=true;

AnsiString T="Ф-001.2 ";
TempAspects->Active=false;


AnsiString NP="Все подразделения";
if(Podr!=0)
{
TempAspects->CommandText="Select * From Подразделения Where [ServerNum]="+IntToStr(Podr);
TempAspects->Active=true;
NP=TempAspects->FieldByName("Название подразделения")->Value;
}


T=T+" Перечень "+IntToStr(Podr);

AnsiString P1=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\Ф-001_2.xlt");
AnsiString P2=WideString(ExtractFilePath(Application->ExeName)+"\Templates\\"+T+".xlt");
CopyFile(P1.c_str() ,P2.c_str() , false);

Variant App =Variant::CreateObject("Excel.Application");
App.OlePropertySet("Visible",false);
Variant Book=App.OlePropertyGet("Workbooks").OleFunction("Add", P2.c_str());
Variant Sheet=App.OlePropertyGet("ActiveSheet");
Sheet.OlePropertySet("Name","Ф-001.2");
App.OlePropertySet("Visible",false);

DeleteFile(P2);


AnsiString Text;
int Num;
T="Реестр значимых экологических аспектов с "+Date1.DateString()+" по "+Date2.DateString();
App.OlePropertyGet("Cells",9,1).OlePropertySet("Value",T.c_str());

T=" "+NP;
App.OlePropertyGet("Cells",10,1).OlePropertySet("Value",T.c_str());

T="Фильтр - "+LFiltr;
App.OlePropertyGet("Cells",12,1).OlePropertySet("Value",T.c_str());

Report->First();
int Number=0;
*/




//for(int i=0;i<Report->RecordCount;i++)
for(Report->First();!Report->Eof;Report->Next())
{
Number=i;
App.OlePropertyGet("Cells",Start+i,1).OlePropertySet("Value",i+1);

Num=Report->FieldByName("Деятельность")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From Деятельность Where [Номер деятельности]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("Наименование деятельности")->Value;
App.OlePropertyGet("Cells",Start+i,2).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("Аспект")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From Аспект Where [Номер аспекта]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("Наименование аспекта")->Value;
App.OlePropertyGet("Cells",Start+i,3).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("Воздействие")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From Воздействия Where [Номер воздействия]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("Наименование воздействия")->Value;
App.OlePropertyGet("Cells",Start+i,4).OlePropertySet("Value",Text.c_str());

Num=Report->FieldByName("Ситуация")->Value;
TempAspects->Active=false;
TempAspects->CommandText="Select * From Ситуации Where [Номер ситуации]="+IntToStr(Num);
TempAspects->Active=true;
Text=TempAspects->FieldByName("Название ситуации")->Value;
App.OlePropertyGet("Cells",Start+i,5).OlePropertySet("Value",Text.c_str());

int Z=Report->FieldByName("Z")->Value;
App.OlePropertyGet("Cells",Start+i,6).OlePropertySet("Value",Z);

AnsiString ZT=ClearString(Report->FieldByName("Мониторинг и контроль")->AsString);
App.OlePropertyGet("Cells",Start+i,7).OlePropertySet("Value",ZT.c_str());

ZT=ClearString(Report->FieldByName("Предлагаемый мониторинг и контроль")->AsString);

App.OlePropertyGet("Cells",Start+i,8).OlePropertySet("Value",ZT.c_str());

ZT=ClearString(Report->FieldByName("Предлагаемые мероприятия")->AsString);

App.OlePropertyGet("Cells",Start+i,9).OlePropertySet("Value",ZT.c_str());

Report->Next();
}
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertySet("VerticalAlignment",-4160);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertySet("WrapText",true);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertyGet("Borders",7).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertyGet("Borders",8).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertyGet("Borders",10).OlePropertySet("Weight",4);
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertyGet("Borders",11).OlePropertySet("Weight",4);
if (Number>=1)
{
App.OlePropertyGet("Range",(Address(Sheet,1,Start)+":"+Address(Sheet,9,Start+Number)).c_str()).OlePropertyGet("Borders",12).OlePropertySet("Weight",4);
}

App.OlePropertyGet("Cells",Start+Number+3,2).OlePropertySet("Value","Согласовано:");
App.OlePropertyGet("Range",(Address(Sheet,2,Start+Number+5)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+Number+6,2).OlePropertySet("Value","(должность)");
App.OlePropertyGet("Cells",Start+Number+6,2).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,2,Start+Number+6)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,2,Start+Number+8)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+Number+9,2).OlePropertySet("Value","(должность)");
App.OlePropertyGet("Cells",Start+Number+9,2).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,2,Start+Number+9)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,4,Start+Number+5)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+Number+6,4).OlePropertySet("Value","(подпись)");
App.OlePropertyGet("Cells",Start+Number+6,4).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+Number+6)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,4,Start+Number+8)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Cells",Start+Number+9,4).OlePropertySet("Value","(подпись)");
App.OlePropertyGet("Cells",Start+Number+9,4).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,4,Start+Number+9)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+5)+":"+Address(Sheet,7,Start+Number+5)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+5)+":"+Address(Sheet,7,Start+Number+5)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+6)+":"+Address(Sheet,7,Start+Number+6)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Cells",Start+Number+6,6).OlePropertySet("Value","(Ф.И.О.)");
App.OlePropertyGet("Cells",Start+Number+6,6).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+6)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+8)+":"+Address(Sheet,7,Start+Number+8)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+8)+":"+Address(Sheet,7,Start+Number+8)).c_str()).OlePropertyGet("Borders",9).OlePropertySet("Weight",2);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+9)+":"+Address(Sheet,7,Start+Number+9)).c_str()).OlePropertySet("MergeCells",true);
App.OlePropertyGet("Cells",Start+Number+9,6).OlePropertySet("Value","(Ф.И.О.)");
App.OlePropertyGet("Cells",Start+Number+9,6).OlePropertyGet("Font").OlePropertySet("Size",8);
App.OlePropertyGet("Range",(Address(Sheet,6,Start+Number+9)).c_str()).OlePropertySet("HorizontalAlignment",-4108);

i++;
Number++;
}
}
App.OlePropertySet("Visible",true);

App=NULL;
Book=NULL;
Sheet=NULL;
}

//------------------------------------------------------------------------------------------------------------------
AnsiString Reports::Address(Variant Sheet,int x,int y)
{
return Sheet.OlePropertyGet("Cells",y,x).OlePropertyGet("Address");
}
//------------------------------------------------------------------------------------------------------------------

void Reports::PrepareMergeAspects()
{
MP<TADODataSet>Temp(Form1);
Temp->Connection=Connect;
Temp->CommandText="Select * From TempAspects order by [Номер аспекта]";
Temp->Active=true;

MP<TADODataSet>Memo(Form1);
Memo->Connection=Connect;


for(Temp->First();!Temp->Eof;Temp->Next())
{
 int TempNum=Temp->FieldByName("Номер аспекта")->Value;

MP<TMemo>M(Form1);
M->Visible=false;
M->Parent=Form1;
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
//---------------------------------------------------------------------------
void Reports::PrepareDemoReport()
{
MP<TADOCommand>Comm(Form1);
Comm->Connection=Zast->ADOUsrAspect;
Comm->CommandText="Delete * from TempAspects";
Comm->Execute();

String S="INSERT INTO TempAspects ( [Номер аспекта], Подразделение, Ситуация, [Вид территории], Деятельность, Специальность, Аспект, Воздействие, G, O, R, S, T, L, N, Z, Значимость, [Проявление воздействия], [Тяжесть последствий], Приоритетность, [Выполняющиеся мероприятия], [Предлагаемые мероприятия], [Мониторинг и контроль], [Предлагаемый мониторинг и контроль], Исполнитель, [Дата создания], [Начало действия], [Конец действия], ServerNum ) ";
S=S+" SELECT Аспекты.[Номер аспекта], Аспекты.Подразделение, Аспекты.Ситуация, Аспекты.[Вид территории], Аспекты.Деятельность, Аспекты.Специальность, Аспекты.Аспект, Аспекты.Воздействие, Аспекты.G, Аспекты.O, Аспекты.R, Аспекты.S, Аспекты.T, Аспекты.L, Аспекты.N, Аспекты.Z, Аспекты.Значимость, Аспекты.[Проявление воздействия], Аспекты.[Тяжесть последствий], Аспекты.Приоритетность, Аспекты.[Выполняющиеся мероприятия], Аспекты.[Предлагаемые мероприятия], Аспекты.[Мониторинг и контроль], Аспекты.[Предлагаемый мониторинг и контроль], Аспекты.Исполнитель, Аспекты.[Дата создания], Аспекты.[Начало действия], Аспекты.[Конец действия], Аспекты.ServerNum ";
S=S+" FROM Аспекты;";
Comm->CommandText=S;
Comm->Execute();
}
//--------------------------------------------------------------------------
String Reports::ClearString(String ZT)
{
while (ZT.Pos("\r"))
{
 int N=ZT.Pos("\r");
 AnsiString Z=ZT.Delete(N,1);
 ZT=Z;
}

while (ZT.Pos("\n"))
{
 int N=ZT.Pos("\n");
 AnsiString Z=ZT.Delete(N,1);
 ZT=Z;
}
return ZT;
}
//----------------------------------------
