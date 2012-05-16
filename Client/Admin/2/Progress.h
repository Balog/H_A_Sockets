//---------------------------------------------------------------------------

#ifndef ProgressH
#define ProgressH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
//---------------------------------------------------------------------------
class TProg : public TForm
{
__published:	// IDE-managed Components
        TProgressBar *PB;
        TLabel *Label1;
private:	// User declarations
public:		// User declarations
        __fastcall TProg(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TProg *Prog;
//---------------------------------------------------------------------------
#endif
