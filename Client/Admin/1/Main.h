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
struct CDBItem
{
int Num;
String Name;
String ServerDB;
int NumDatabase;
};
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
        TAction *RegForm;
        void __fastcall ClientSocketRead(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall ClientSocketWrite(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall RegFormExecute(TObject *Sender);
private:	// User declarations
String Path;
String AdminDatabase;
int LicCount;

vector<String>Parameters;

String GetIP();
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);

MDBConnector* Database;

Client *MClient;
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
