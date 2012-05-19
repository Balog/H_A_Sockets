//---------------------------------------------------------------------------

#ifndef LGH
#define LGH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TLabel *Label1;
        TEdit *Edit1;
        TUpDown *UpDown1;
        TCheckBox *CheckBox1;
        TButton *Button1;
        TLabel *Label2;
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall CheckBox1Click(TObject *Sender);
private:	// User declarations
String Propis(int N);
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
