//---------------------------------------------------------------------------

#ifndef AboutH
#define AboutH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
#include <Graphics.hpp>
#include <Dialogs.hpp>
#include <jpeg.hpp>

//---------------------------------------------------------------------------
class TFAbout : public TForm
{
__published:	// IDE-managed Components
        TLabel *Label1;
        TImage *Image1;
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall FormClick(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall FormKeyUp(TObject *Sender, WORD &Key,
          TShiftState Shift);
private:	// User declarations
public:		// User declarations
        __fastcall TFAbout(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFAbout *FAbout;
//---------------------------------------------------------------------------
#endif
