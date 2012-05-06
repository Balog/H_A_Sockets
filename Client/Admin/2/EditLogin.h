//---------------------------------------------------------------------------

#ifndef EditLoginH
#define EditLoginH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Mask.hpp>
//---------------------------------------------------------------------------
class TEditLogins : public TForm
{
__published:	// IDE-managed Components
        TLabel *Label1;
        TLabel *Label2;
        TEdit *Log;
        TMaskEdit *Pass1;
        TMaskEdit *Pass2;
        TButton *Button1;
private:	// User declarations
public:		// User declarations
        __fastcall TEditLogins(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TEditLogins *EditLogins;
//---------------------------------------------------------------------------
#endif
