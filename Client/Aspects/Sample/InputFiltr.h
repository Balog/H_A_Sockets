//---------------------------------------------------------------------------

#ifndef InputFiltrH
#define InputFiltrH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "MainForm.h"
#include "Zastavka.h"
#include <ADODB.hpp>
#include <DB.hpp>
#include <DBCtrls.hpp>
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TFilter : public TForm
{
__published:	// IDE-managed Components
        TRadioGroup *RadioGroup1;
        TADODataSet *Podr;
        TDataSource *DataSource1;
        TEdit *Edit1;
        TEdit *Edit2;
        TEdit *Edit3;
        TEdit *Edit4;
        TButton *Button1;
        TButton *Button2;
        TButton *Button3;
        TButton *Button4;
        TButton *Button5;
        TComboBox *ComboBox1;
        TComboBox *ComboBox2;
        TComboBox *ComboBox3;
        void __fastcall Button6Click(TObject *Sender);
        void __fastcall RadioGroup1Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
        void __fastcall Button3Click(TObject *Sender);
        void __fastcall Button4Click(TObject *Sender);
        void __fastcall ComboBox1Click(TObject *Sender);
        void __fastcall Button5Click(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall ComboBox2Click(TObject *Sender);
        void __fastcall ComboBox3Click(TObject *Sender);
        void __fastcall ComboBox1KeyPress(TObject *Sender, char &Key);
        void __fastcall ComboBox2KeyPress(TObject *Sender, char &Key);
        void __fastcall ComboBox3KeyPress(TObject *Sender, char &Key);
        void __fastcall Edit1DblClick(TObject *Sender);
        void __fastcall Edit2DblClick(TObject *Sender);
        void __fastcall Edit3DblClick(TObject *Sender);
        void __fastcall Edit4DblClick(TObject *Sender);
        void __fastcall FormKeyUp(TObject *Sender, WORD &Key,
          TShiftState Shift);
private:	// User declarations
public:		// User declarations
        __fastcall TFilter(TComponent* Owner);

void InpTer();
void InpDeyat();
void InpAsp();
void InpVozd();

int Index;
AnsiString CText;
AnsiString Filtr1, SeFiltr1, Filtr2;
};
//---------------------------------------------------------------------------
extern PACKAGE TFilter *Filter;
//---------------------------------------------------------------------------
#endif
