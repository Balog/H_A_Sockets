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
        TEdit *Log;
        TMaskEdit *Pass1;
        TMaskEdit *Pass2;
        TButton *Button1;
        TLabel *Label1;
        TLabel *Label2;
        void __fastcall FormShow(TObject *Sender);
        void __fastcall Pass1Change(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall FormKeyUp(TObject *Sender, WORD &Key,
          TShiftState Shift);
private:	// User declarations
public:		// User declarations
        __fastcall TEditLogins(TComponent* Owner);
int Mode;
String Login;
String Pass;        
};
//---------------------------------------------------------------------------
extern PACKAGE TEditLogins *EditLogins;
//---------------------------------------------------------------------------
#endif
