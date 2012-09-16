//---------------------------------------------------------------------------

#ifndef MetodH
#define MetodH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Menus.hpp>
#include "F_Vvedenie.h"
#include <ADODB.hpp>
#include <ComCtrls.hpp>
#include <DB.hpp>
#include <ExtCtrls.hpp>
#include <Graphics.hpp>
#include <DBCtrls.hpp>
#include "Zastavka.h"
#include <Dialogs.hpp>
#include "MainForm.h"
#include <jpeg.hpp>
//---------------------------------------------------------------------------
class TMetodika : public TForm
{
__published:	// IDE-managed Components
        TMainMenu *MainMenu1;
        TMenuItem *N4;
        TMenuItem *N5;
        TMenuItem *N6;
        TMenuItem *N8;
        TStaticText *StaticText1;
        TLabel *Label1;
        TADODataSet *ADODataSet1;
        TDataSource *DataSource1;
        TDBMemo *DBMemo1;
        TImage *Image1;
        TMenuItem *N1;
        TImage *Image2;
        TADODataSet *ADODataSet2;
        TMenuItem *N2;
        void __fastcall N6Click(TObject *Sender);
        void __fastcall N4Click(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall N8Click(TObject *Sender);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall N2Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall FormActivate(TObject *Sender);
        void __fastcall FormKeyUp(TObject *Sender, WORD &Key,
          TShiftState Shift);
private:	// User declarations
void ReadMet();
bool Registered;
public:		// User declarations
        __fastcall TMetodika(TComponent* Owner);

        BEGIN_MESSAGE_MAP
          VCL_MESSAGE_HANDLER(WM_SYSCOMMAND, TMessage, WMSysCommand)
        END_MESSAGE_MAP(TComponent);
void __fastcall WMSysCommand(TMessage & Msg);        
};
//---------------------------------------------------------------------------
extern PACKAGE TMetodika *Metodika;
//---------------------------------------------------------------------------
#endif
