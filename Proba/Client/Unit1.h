//---------------------------------------------------------------------------

#ifndef Unit1H
#define Unit1H
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ScktComp.hpp>
#include <ActnList.hpp>
#include <ActnMan.hpp>
#include <vector>
using namespace std;

//---------------------------------------------------------------------------
struct Parameters
{
int RespComm;
vector<String>SParam;
String ActTrue;
String ActFalse;
};
//---------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TClientSocket *ClientSocket1;
        TMemo *Memo1;
        TButton *Button1;
        TActionManager *ActionManager1;
        TAction *Command1;
        TAction *Command2;
        TAction *Command3;
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall ClientSocket1Read(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall Command1Execute(TObject *Sender);
        void __fastcall Command2Execute(TObject *Sender);
        void __fastcall Command3Execute(TObject *Sender);
private:	// User declarations
void RunAction(String);
vector<Parameters>Param;
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
