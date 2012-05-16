//---------------------------------------------------------------------------


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
#include <stdio.h>
#include <ActnList.hpp>
#include <ActnMan.hpp>
#include <vector>
using namespace std;
const MyTrayIcon = WM_USER+1;

//---------------------------------------------------------------------------
#ifndef MainH
/*
struct BaseItem
{
String Name;
String Patch;
bool MainSpec;
String FileName;
int LicCount;
};
*/
#define MainH

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
        TServerSocket *ServerSocket;
        TActionManager *ActionManager1;
        TAction *ConnectDatabase;
        TPopupMenu *PopupMenu2;
        TMenuItem *N5;
        TMainMenu *MainMenu1;
        TMenuItem *N6;
        TMenuItem *N7;
        void __fastcall ServerSocketClientConnect(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall ServerSocketClientDisconnect(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall ServerSocketClientError(TObject *Sender,
          TCustomWinSocket *Socket, TErrorEvent ErrorEvent,
          int &ErrorCode);
        void __fastcall ServerSocketClientRead(TObject *Sender,
          TCustomWinSocket *Socket);
        void __fastcall FormDestroy(TObject *Sender);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall RoleClick(TObject *Sender);
        void __fastcall BaseClick(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall N2Click(TObject *Sender);
        void __fastcall N3Click(TObject *Sender);
        void __fastcall N4Click(TObject *Sender);
        void __fastcall UsersClick(TObject *Sender);
        void __fastcall UsersMouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall OtdelsMouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall N5Click(TObject *Sender);
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
        void __fastcall N6Click(TObject *Sender);
private:	// User declarations
int Block;
Clients* Cl;
String Path;
vector<String>Parameters;
//vector<BaseItem>VBases;
void __fastcall ChangeCountClient(TObject *Sender);
void __fastcall MTIcon(TMessage&);
void __fastcall WMSysCommand(TMessage & Msg);

int AnalizLic(String Text, String Pass);
String GetFileDatabase(String NameDatabase);
int ReadRes(String Key1, String Key2);
int Propis(String S);
int LicCount;
int EditLoginNumber;

void VerifyLicense();
void __fastcall TForm1::SelectOtdel(TObject *Sender);
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
long GetLicenseCount(String DBName);
void UpdateTempLogin();
void UpdateOtdel(int NumLogin);
void SaveCode(String Login, String Password);
TNotifyIconData NID;

        BEGIN_MESSAGE_MAP
          VCL_MESSAGE_HANDLER(MyTrayIcon,TMessage,MTIcon);
          VCL_MESSAGE_HANDLER(WM_SYSCOMMAND, TMessage, WMSysCommand)
        END_MESSAGE_MAP(TComponent);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
