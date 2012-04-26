//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
//#include "ServerClass.h"
#include <ADODB.hpp>
#include <DB.hpp>
#include <ExtCtrls.hpp>
#include <Menus.hpp>
#include <ScktComp.hpp>
#include "ClientClass.h"
#include <ActnList.hpp>
#include <ActnMan.hpp>
#include <vector>
#include "MDBConnector.h"
#ifndef MainH
#define MainH
//---------------------------------------------------------------------------
using namespace std;


//---------------------------------------------------------------------------


class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TLabel *Label1;
        TLabel *Label2;
        TLabel *Label3;
        TLabel *Label4;
        TComboBox *Role;
        TListBox *Users;
        TListBox *Otdels;
        TComboBox *CBDatabase;
        TADOConnection *Database1;
        TPopupMenu *PopupMenu1;
        TMenuItem *N1;
        TMenuItem *N2;
        TMenuItem *N3;
        TMenuItem *N4;
        TPopupMenu *FreeOtdel;
        TMainMenu *MainMenu1;
        TMenuItem *N5;
        TMenuItem *N6;
        TMenuItem *N7;
        TMenuItem *N10;
        TMenuItem *N8;
        TMenuItem *N9;
        TTimer *DBTimer;
        TClientSocket *ClientSocket;
        TActionManager *ActionManager1;
        TAction *RegForm_Form1;
        TAction *RegForm_Form2;
        TAction *PostRegForm_Form1;
        TAction *PostRegForm_Form2;
        TAction *AspectsConnect;
        TAction *DiaryConnect;
        TAction *LoadLogins;
        TAction *ViewLogins;
        void __fastcall ClientSocketRead(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall ClientSocketWrite(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall RegForm_Form1Execute(TObject *Sender);
        void __fastcall RegForm_Form2Execute(TObject *Sender);
        void __fastcall PostRegForm_Form1Execute(TObject *Sender);
        void __fastcall PostRegForm_Form2Execute(TObject *Sender);
        void __fastcall AspectsConnectExecute(TObject *Sender);
        void __fastcall DiaryConnectExecute(TObject *Sender);
        void __fastcall LoadLoginsExecute(TObject *Sender);
        void __fastcall ViewLoginsExecute(TObject *Sender);
        void __fastcall FormCreate(TObject *Sender);
private:	// User declarations
String Path;
String AdminDatabase;
String DiaryDatabase;
int LicCount;

vector<String>Parameters;

String GetIP();

String ServerName;
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);


Client *MClient;
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
