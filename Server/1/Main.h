//---------------------------------------------------------------------------

#ifndef MainH
#define MainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ADODB.hpp>
#include <ComCtrls.hpp>
#include <DB.hpp>
#include <ImgList.hpp>
#include <Menus.hpp>
#include <ScktComp.hpp>
#include "ServerClass.h"
#include <vector>
using namespace std;
const MyTrayIcon = WM_USER+1;

//---------------------------------------------------------------------------
struct BaseItem
{
String Name;
String Patch;
bool MainSpec;
String FileName;
int LicCount;
};
//-----------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TPageControl *PageControl1;
        TTabSheet *TabSheet1;
        TLabel *Label2;
        TLabel *Label1;
        TLabel *Label3;
        TListBox *ListBox1;
        TListBox *ListBox2;
        TTabSheet *TabSheet2;
        TComboBox *Role;
        TListBox *Otdels;
        TListBox *Users;
        TButton *Button1;
        TEdit *Edit1;
        TComboBox *Base;
        TADOConnection *Database;
        TPopupMenu *PopupMenu1;
        TMenuItem *N1;
        TMenuItem *N2;
        TMenuItem *N3;
        TMenuItem *N4;
        TImageList *ImageList1;
        TPopupMenu *FreeOtdel;
        TPopupMenu *PopupMenu2;
        TMenuItem *N5;
        TServerSocket *ServerSocket;
        void __fastcall ServerSocketClientConnect(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall ServerSocketClientDisconnect(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall ServerSocketClientError(TObject *Sender,
          TCustomWinSocket *Socket, TErrorEvent ErrorEvent,
          int &ErrorCode);
        void __fastcall ServerSocketClientRead(TObject *Sender,
          TCustomWinSocket *Socket);
private:	// User declarations
int Block;
Clients* Cl;
String Path;
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
