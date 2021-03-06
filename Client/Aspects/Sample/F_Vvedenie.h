//---------------------------------------------------------------------------

#ifndef F_VvedenieH
#define F_VvedenieH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
//#include "Monitor.h"
#include <Menus.hpp>
#include <ExtCtrls.hpp>
#include <Graphics.hpp>
#include "Metod.h"
//#include "About.h"
#include "Zastavka.h"
#include <Dialogs.hpp>
#include <ADODB.hpp>
#include <DB.hpp>
#include "MainForm.h"
#include "Docs.h"
#include "Settings.h"
//#include "Form_SendFile.h"
#include <jpeg.hpp>

//---------------------------------------------------------------------------
class TVvedenie : public TForm
{
__published:	// IDE-managed Components
        TImage *Image1;
        TStaticText *StaticText2;
        TButton *Button1;
        TMainMenu *MainMenu1;
        TMenuItem *N4;
        TMenuItem *N5;
        TMenuItem *N6;
        TMenuItem *N7;
        TMenuItem *N1;
        TMenuItem *N8;
        void __fastcall N6Click(TObject *Sender);
        void __fastcall N8Click(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall N5Click(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall N9Click(TObject *Sender);
        void __fastcall N2Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall N10Click(TObject *Sender);
        void __fastcall N11Click(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
        void __fastcall FormActivate(TObject *Sender);
        void __fastcall N12Click(TObject *Sender);
        void __fastcall N14Click(TObject *Sender);
        void __fastcall N17Click(TObject *Sender);
        void __fastcall N15Click(TObject *Sender);
        void __fastcall N16Click(TObject *Sender);
        void __fastcall N18Click(TObject *Sender);
        void __fastcall N19Click(TObject *Sender);
        void __fastcall FormKeyUp(TObject *Sender, WORD &Key,
          TShiftState Shift);
private:	// User declarations
public:		// User declarations
        __fastcall TVvedenie(TComponent* Owner);

        BEGIN_MESSAGE_MAP
          VCL_MESSAGE_HANDLER(WM_SYSCOMMAND, TMessage, WMSysCommand)
        END_MESSAGE_MAP(TComponent);
void __fastcall WMSysCommand(TMessage & Msg);
};
//---------------------------------------------------------------------------
extern PACKAGE TVvedenie *Vvedenie;
//---------------------------------------------------------------------------
#endif
