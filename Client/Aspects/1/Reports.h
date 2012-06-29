//---------------------------------------------------------------------------
#include <ADODB.hpp>
#include <DB.hpp>
#include "IdGlobal.hpp"
#ifndef ReportsH
#define ReportsH
//---------------------------------------------------------------------------

class Reports
{
public:
void CreateReport1(int Podr, TDateTime Date1, TDateTime Date2, AnsiString Isp, String Filtr1, String LFiltr, bool OneList);
void CreateReport2(int Podr, TDateTime Date1, TDateTime Date2, AnsiString Isp, String Filtr2, String LFiltr, bool OneList);
TADOConnection *Connect;
int Role;
private:
AnsiString Address(Variant Sheet,int x,int y);

void PrepareDemoReport();
void PrepareMergeAspects();
String ClearString(String ZT);
};
#endif
