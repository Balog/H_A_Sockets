//---------------------------------------------------------------------------

#ifndef ProgressH
#define ProgressH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TFProg : public TForm
{
__published:	// IDE-managed Components
        TProgressBar *Progress;
        TPanel *Panel1;
        TLabel *Label1;
        TTimer *Timer1;
        void __fastcall FormShow(TObject *Sender);
        void __fastcall FormHide(TObject *Sender);
        void __fastcall Timer1Timer(TObject *Sender);
        void __fastcall FormActivate(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TFProg(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFProg *FProg;
//---------------------------------------------------------------------------
#endif
