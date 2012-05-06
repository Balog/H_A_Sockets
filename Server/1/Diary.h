//---------------------------------------------------------------------------

#ifndef DiaryH
#define DiaryH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
#include <ADODB.hpp>
#include <CheckLst.hpp>
#include <DB.hpp>
#include <DBGrids.hpp>
#include <Grids.hpp>
//---------------------------------------------------------------------------
class TFDiary : public TForm
{
__published:	// IDE-managed Components
        TLabel *Label1;
        TLabel *Label2;
        TMonthCalendar *NDate;
        TMonthCalendar *KDate;
        TCheckListBox *Comps;
        TLabel *Label3;
        TCheckListBox *Logins;
        TLabel *Label4;
        TCheckListBox *Types;
        TLabel *Label5;
        TDBGrid *Diary;
        TCheckBox *EnNDate;
        TCheckBox *EnKDate;
        TDataSource *DataSource1;
        TADODataSet *Events;
        TADOConnection *ADODiary;
        TButton *Button1;
        void __fastcall FormShow(TObject *Sender);
        void __fastcall EnNDateClick(TObject *Sender);
        void __fastcall EnKDateClick(TObject *Sender);
        void __fastcall NDateClick(TObject *Sender);
        void __fastcall NTimeChange(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
private:	// User declarations
void Refresh();
void Initialize();
public:		// User declarations
        __fastcall TFDiary(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFDiary *FDiary;
//---------------------------------------------------------------------------
#endif