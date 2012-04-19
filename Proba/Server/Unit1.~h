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
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TServerSocket *ServerSocket1;
        TMemo *Memo1;
        TEdit *Edit1;
        TButton *Button1;
        TEdit *Edit2;
        TButton *Button2;
        TActionManager *ActionManager1;
        TAction *Comm1;
        TAction *Comm2;
        TAction *Comm3;
        void __fastcall ServerSocket1ClientConnect(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall ServerSocket1ClientDisconnect(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
        void __fastcall ServerSocket1ClientRead(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall Comm1Execute(TObject *Sender);
        void __fastcall Comm2Execute(TObject *Sender);
        void __fastcall Comm3Execute(TObject *Sender);
private:	// User declarations
TCustomWinSocket *WS;
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
