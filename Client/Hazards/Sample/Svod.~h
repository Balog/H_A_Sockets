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
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall N13Click(TObject *Sender);
        void __fastcall N14Click(TObject *Sender);
        void __fastcall N15Click(TObject *Sender);
        void __fastcall N16Click(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall N4Click(TObject *Sender);
        void __fastcall N5Click(TObject *Sender);
        void __fastcall N6Click(TObject *Sender);
        void __fastcall N2Click(TObject *Sender);
        void __fastcall N9Click(TObject *Sender);
        void __fastcall N11Click(TObject *Sender);
        void __fastcall N12Click(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
        void __fastcall DocumentsN29Click(TObject *Sender);
        void __fastcall N18Click(TObject *Sender);
        void __fastcall N19Click(TObject *Sender);
        void __fastcall N20Click(TObject *Sender);
        void __fastcall N21Click(TObject *Sender);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall N8Click(TObject *Sender);
        void __fastcall N22Click(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
        void __fastcall FormKeyUp(TObject *Sender, WORD &Key,
          TShiftState Shift);
private:	// User declarations
bool CreateMainSvod();
void AddMainSvod(TADODataSet *Table);
void CreateRep(TADODataSet *TempTable, Variant App, Variant Book, Variant Sheet, int& Start, int& NN, int Number);
void EndSvod(Variant App, Variant Sheet, int Start);

void PrepareMergeAspects();
bool Registered;
public:		// User declarations
        __fastcall TFSvod(TComponent* Owner);
//        void __fastcall MyException(TObject *Sender, Exception *E);
AnsiString TFSvod::Address(Variant Sheet,int x,int y);
void TFSvod::Initialize();
};
//---------------------------------------------------------------------------
extern PACKAGE TFSvod *FSvod;
//---------------------------------------------------------------------------
#endif
