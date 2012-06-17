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
        TButton *Button1;
        TButton *Button2;
        TButton *Button3;
        TButton *Button4;
        TButton *Button5;
        TComboBox *ComboBox1;
        TComboBox *ComboBox2;
        TComboBox *ComboBox3;
        TComboBox *ComboBox4;
        TComboBox *ComboBox5;
        TComboBox *ComboBox6;
        TComboBox *ComboBox7;
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
        void __fastcall ComboBox4Change(TObject *Sender);
        void __fastcall ComboBox4DropDown(TObject *Sender);
        void __fastcall ComboBox4KeyPress(TObject *Sender, char &Key);
        void __fastcall ComboBox4Select(TObject *Sender);
        void __fastcall ComboBox4Enter(TObject *Sender);
        void __fastcall ComboBox5Change(TObject *Sender);
        void __fastcall ComboBox5DropDown(TObject *Sender);
        void __fastcall ComboBox5KeyPress(TObject *Sender, char &Key);
        void __fastcall ComboBox5Select(TObject *Sender);
        void __fastcall ComboBox5Enter(TObject *Sender);
        void __fastcall ComboBox6Change(TObject *Sender);
        void __fastcall ComboBox6DropDown(TObject *Sender);
        void __fastcall ComboBox6KeyPress(TObject *Sender, char &Key);
        void __fastcall ComboBox6Select(TObject *Sender);
        void __fastcall ComboBox6Enter(TObject *Sender);
        void __fastcall ComboBox7Change(TObject *Sender);
        void __fastcall ComboBox7DropDown(TObject *Sender);
        void __fastcall ComboBox7KeyPress(TObject *Sender, char &Key);
        void __fastcall ComboBox7Select(TObject *Sender);
        void __fastcall ComboBox7Enter(TObject *Sender);
private:	// User declarations

int NumItemCombo;
public:		// User declarations
        __fastcall TFilter(TComponent* Owner);
int NumFiltr;
void InpTer();
void InpDeyat();
void InpAsp();
void InpVozd();

int Index;
AnsiString CText;
AnsiString Filtr1, SeFiltr1, Filtr2;
void SetDefFiltr();
};
//---------------------------------------------------------------------------
extern PACKAGE TFilter *Filter;
//---------------------------------------------------------------------------
#endif
