//---------------------------------------------------------------------------

#ifndef PassFormH
#define PassFormH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
#include "MasterPointer.h"
//---------------------------------------------------------------------------
class TPass : public TForm
{
__published:	// IDE-managed Components
        TBevel *Bevel1;
        TLabel *Label1;
        TLabel *Label2;
        TComboBox *CbLogin;
        TEdit *EdPass;
        TButton *Button1;
        void __fastcall FormShow(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall EdPassKeyPress(TObject *Sender, char &Key);
        void __fastcall FormCreate(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TPass(TComponent* Owner);

void ViewLogins();        
};
//---------------------------------------------------------------------------
extern PACKAGE TPass *Pass;
//---------------------------------------------------------------------------
#endif
