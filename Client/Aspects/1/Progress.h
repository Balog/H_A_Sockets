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
class TProg : public TForm
{
__published:	// IDE-managed Components
        TPanel *Panel1;
        TLabel *Label1;
        TProgressBar *PB;
        void __fastcall FormCreate(TObject *Sender);
private:	// User declarations
public:		// User declarations
bool SignComplete;
        __fastcall TProg(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TProg *Prog;
//---------------------------------------------------------------------------
#endif
