//---------------------------------------------------------------------------

#ifndef MoveAspectsH
#define MoveAspectsH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Grids.hpp>
#include <dbcgrids.hpp>
//---------------------------------------------------------------------------
class TMoveAsp : public TForm
{
__published:	// IDE-managed Components
        TStringGrid *TabAsp;
        void __fastcall FormShow(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall TabAspSelectCell(TObject *Sender, int ACol,
          int ARow, bool &CanSelect);
private:	// User declarations
TPanel* Pan;
TComboBox* CPodr;
public:		// User declarations
        __fastcall TMoveAsp(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TMoveAsp *MoveAsp;
//---------------------------------------------------------------------------
#endif
