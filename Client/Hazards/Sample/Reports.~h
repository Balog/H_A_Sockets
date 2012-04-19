//---------------------------------------------------------------------------
#include <ADODB.hpp>
#include <DB.hpp>
#ifndef ReportsH
#define ReportsH
//---------------------------------------------------------------------------

class Reports
{
public:
void CreateReport1(int Podr, TDateTime Date1, TDateTime Date2, AnsiString Isp, String Filtr1, String LFiltr);
void CreateReport2(int Podr, TDateTime Date1, TDateTime Date2, AnsiString Isp, String Filtr2, String LFiltr);
TADOConnection *Connect;
int Role;
private:
AnsiString Address(Variant Sheet,int x,int y);
void PrepareReport();
void PrepareDemoReport();
void PrepareMergeAspects();
String ClearString(String ZT);
};
#endif
