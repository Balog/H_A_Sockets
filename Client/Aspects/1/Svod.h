//---------------------------------------------------------------------------

#ifndef SvodH
#define SvodH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ADODB.hpp>
#include <DB.hpp>
#include <DBGrids.hpp>
#include <Grids.hpp>
#include <Menus.hpp>
#include <Dialogs.hpp>
//---------------------------------------------------------------------------
class TFSvod : public TForm
{
__published:	// IDE-managed Components
        TADODataSet *TSvod;
        TDataSource *DataSource1;
        TDBGrid *DBGrid1;
        TPopupMenu *PopupMenu1;
        TMenuItem *N13;
        TMenuItem *N14;
        TMenuItem *N15;
        TMenuItem *N16;
        TOpenDialog *OpenDialog1;
        TADOConnection *Temp;
        TADODataSet *TempTable;
        TButton *Button1;
        TEdit *MainPatch;
        TLabel *Label1;
        TButton *Button2;
        TLabel *Label2;
        void __fastcall FormActivate(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall N13Click(TObject *Sender);
        void __fastcall N14Click(TObject *Sender);
        void __fastcall N15Click(TObject *Sender);
        void __fastcall N16Click(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall N8Click(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
        void __fastcall FormKeyUp(TObject *Sender, WORD &Key,
          TShiftState Shift);
private:	// User declarations
void CreateMainSvod();


void PrepareMergeAspects();
bool Registered;
public:		// User declarations
        __fastcall TFSvod(TComponent* Owner);
AnsiString TFSvod::Address(Variant Sheet,int x,int y);
void TFSvod::Initialize();
void ContSvodReport();
void CreateRep(TADODataSet *TempTable, Variant App, Variant Book, Variant Sheet, int& Start, int& NN, int Number);
void EndSvod(Variant App, Variant Sheet, int Start);


};
//---------------------------------------------------------------------------
extern PACKAGE TFSvod *FSvod;
//---------------------------------------------------------------------------
#endif
