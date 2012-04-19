//---------------------------------------------------------------------------

#ifndef SettingsH
#define SettingsH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
#include <DB.hpp>
#include <DBGrids.hpp>
#include <Grids.hpp>
#include "Zastavka.h"
//#include "F_Vvedenie.h"
//#include "Metod.h"
//#include "MainForm.h"
#include "Docs.h"
#include <ADODB.hpp>
#include <Menus.hpp>
//#include "Form_SendFile.h"
//---------------------------------------------------------------------------
class TSetting : public TForm
{
__published:	// IDE-managed Components
        TPageControl *PageControl1;
        TTabSheet *TabSheet1;
        TTabSheet *TabSheet2;
        TMainMenu *MainMenu1;
        TMenuItem *N4;
        TMenuItem *N5;
        TMenuItem *N6;
        TMenuItem *N2;
        TMenuItem *N3;
        TMenuItem *N7;
        TMenuItem *N1;
        TMenuItem *N8;
        TDBGrid *DBGrid1;
        TDataSource *DataSource1;
        TADODataSet *Podr;
        TPopupMenu *PopupMenu1;
        TMenuItem *N9;
        TMenuItem *N10;
        TADODataSet *Temp;
        TDBGrid *DBGrid2;
        TADODataSet *Sit;
        TDataSource *DataSource2;
        TPopupMenu *PopupMenu2;
        TMenuItem *N11;
        TMenuItem *N12;
        TMenuItem *N13;
        TMenuItem *N14;
        TMenuItem *N15;
        TMenuItem *N16;
        TMenuItem *N17;
        TMenuItem *N18;
        TMenuItem *N19;
        TMenuItem *N20;
        TMenuItem *N21;
        TMenuItem *N22;
        TMenuItem *N23;
        TMenuItem *N24;
        void __fastcall N6Click(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall N4Click(TObject *Sender);
        void __fastcall N5Click(TObject *Sender);
        void __fastcall N2Click(TObject *Sender);
        void __fastcall N8Click(TObject *Sender);
        void __fastcall N9Click(TObject *Sender);
        void __fastcall N10Click(TObject *Sender);
        void __fastcall N11Click(TObject *Sender);
        void __fastcall N12Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall N15Click(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
        void __fastcall N16Click(TObject *Sender);
        void __fastcall FormActivate(TObject *Sender);
        void __fastcall N17Click(TObject *Sender);
        void __fastcall N19Click(TObject *Sender);
        void __fastcall N20Click(TObject *Sender);
        void __fastcall N21Click(TObject *Sender);
        void __fastcall N22Click(TObject *Sender);
        void __fastcall N23Click(TObject *Sender);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall N24Click(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TSetting(TComponent* Owner);
void TSetting::Initialize();        
};
//---------------------------------------------------------------------------
extern PACKAGE TSetting *Setting;
//---------------------------------------------------------------------------
#endif
