//---------------------------------------------------------------------------

#ifndef Rep1H
#define Rep1H
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
#include <ADODB.hpp>
#include <DB.hpp>
//---------------------------------------------------------------------------
class TReport1 : public TForm
{
__published:	// IDE-managed Components
        TLabel *Label2;
        TComboBox *CPodrazdel;
        TEdit *Edit1;
        TLabel *Label1;
        TDateTimePicker *Date1;
        TLabel *Label3;
        TDateTimePicker *Date2;
        TLabel *Label4;
        TButton *Button1;
        TADODataSet *Podr;
        TCheckBox *CheckBox1;
        void __fastcall FormShow(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall CPodrazdelKeyPress(TObject *Sender, char &Key);
        void __fastcall Date1KeyPress(TObject *Sender, char &Key);
        void __fastcall Date2KeyPress(TObject *Sender, char &Key);
        void __fastcall FormKeyUp(TObject *Sender, WORD &Key,
          TShiftState Shift);
private:	// User declarations
bool Registration;
public:		// User declarations
        __fastcall TReport1(TComponent* Owner);
int Role;
int NumRep;
String Flt;
String FltName;
String PodrComText;
TADOConnection *RepBase;
int NumLogin;
void CreateRep();
};
//---------------------------------------------------------------------------
extern PACKAGE TReport1 *Report1;
//---------------------------------------------------------------------------
#endif
