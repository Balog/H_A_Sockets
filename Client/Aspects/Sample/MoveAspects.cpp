//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "MoveAspects.h"
#include "MasterPointer.h"
#include "Zastavka.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TMoveAsp *MoveAsp;
//---------------------------------------------------------------------------
__fastcall TMoveAsp::TMoveAsp(TComponent* Owner)
        : TForm(Owner)
{

}
//---------------------------------------------------------------------------
void __fastcall TMoveAsp::FormShow(TObject *Sender)
{
MP<TADODataSet>Aspects(this);
Aspects->Connection=Zast->ADOAspect;
Aspects->CommandText="SELECT Аспекты.Подразделение, Подразделения.[Название подразделения], Деятельность.[Наименование деятельности], Аспект.[Наименование аспекта], Воздействия.[Наименование воздействия], Ситуации.[Название ситуации], Аспекты.Z FROM Ситуации INNER JOIN (Воздействия INNER JOIN (Аспект INNER JOIN (Деятельность INNER JOIN (Подразделения INNER JOIN Аспекты ON Подразделения.[Номер подразделения] = Аспекты.Подразделение) ON Деятельность.[Номер деятельности] = Аспекты.Деятельность) ON Аспект.[Номер аспекта] = Аспекты.Аспект) ON Воздействия.[Номер воздействия] = Аспекты.Воздействие) ON Ситуации.[Номер ситуации] = Аспекты.Ситуация ORDER BY Подразделения.[Название подразделения], Аспекты.[Номер аспекта];";
Aspects->Active=true;

TabAsp->Cells[0][0]="Подразделение";
TabAsp->Cells[1][0]="Деятельность";
TabAsp->Cells[2][0]="Аспект";
TabAsp->Cells[3][0]="Воздействие";
TabAsp->Cells[4][0]="Ситуация";
TabAsp->Cells[5][0]="Баллы (Z)";

int i=1;
for(Aspects->First();!Aspects->Eof;Aspects->Next())
{
TabAsp->Cells[0][i]=Aspects->FieldByName("Название подразделения")->Value;
TabAsp->Cells[1][i]=Aspects->FieldByName("Наименование деятельности")->Value;
TabAsp->Cells[2][i]=Aspects->FieldByName("Наименование аспекта")->Value;
TabAsp->Cells[3][i]=Aspects->FieldByName("Наименование воздействия")->Value;
TabAsp->Cells[4][i]=Aspects->FieldByName("Название ситуации")->Value;
TabAsp->Cells[5][i]=Aspects->FieldByName("Z")->Value;
i++;
TabAsp->RowCount++;
}
TabAsp->RowCount--;

Pan=new TPanel(this);
Pan->Visible=false;
Pan->Parent=TabAsp;
CPodr=new TComboBox(this);
CPodr->Visible=false;
CPodr->Parent=Pan;
CPodr->Align=alClient;

MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOConn;
Podr->CommandText="Select * From Подразделения order by [Номер подразделения]";
Podr->Active=true;

 CPodr->Clear();
for(Podr->First();!Podr->Eof;Podr->Next())
{
 CPodr->Items->Add(Podr->FieldByName("Название подразделения")->Value);
}
}
//---------------------------------------------------------------------------
void __fastcall TMoveAsp::FormClose(TObject *Sender, TCloseAction &Action)
{
delete CPodr;
}
//---------------------------------------------------------------------------
void __fastcall TMoveAsp::TabAspSelectCell(TObject *Sender, int ACol,
      int ARow, bool &CanSelect)
{
if(ACol==1)
{
 TRect RT=TabAsp->CellRect(1,ARow);
 Pan->Left=RT.Left;
 Pan->Top=RT.Top;
 Pan->Height=RT.Height();
 Pan->Width=TabAsp->ColWidths[1]+2;

 Pan->Visible=true;
 CPodr->Visible=true;
 CPodr->BringToFront();
 TabAsp->Repaint();
 Pan->Repaint();
 CPodr->Repaint();

}
else
{
 Pan->Visible=false;
 CPodr->Visible=false;
}
}
//---------------------------------------------------------------------------
