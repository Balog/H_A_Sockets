//---------------------------------------------------------------------------

#ifndef EditLoginH
#define EditLoginH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Mask.hpp>
#include <ADODB.hpp>
#include <DB.hpp>
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
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall Pass1Change(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TEditLogins(TComponent* Owner);
TLocateOptions SO;


bool Mode; //false - �������������� ������ � ������
           //true - �������������� ������ ������
int NumLogin; //����� �������������� ������
             //-1 - � ������ ���������� ������ ������

void NewLogin();
void EditLogin(int Num);
void EditPass(int Num);
};
//---------------------------------------------------------------------------
extern PACKAGE TEditLogins *EditLogins;
//---------------------------------------------------------------------------
#endif