//---------------------------------------------------------------------------

#ifndef UServerH
#define UServerH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "ServerClasses.h"
#include <ADODB.hpp>
#include <DB.hpp>
#include <DBGrids.hpp>
#include <Grids.hpp>
#include <DBCtrls.hpp>
#include <Menus.hpp>
#include <ComCtrls.hpp>
#include <AppEvnts.hpp>
#include <sysmac.h>
#include <ImgList.hpp>
#include <ExtCtrls.hpp>
#include <FileCtrl.hpp>
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
        TTabSheet *TabSheet2;
        TLabel *Label2;
        TLabel *Label1;
        TListBox *ListBox1;
        TListBox *ListBox2;
        TADOConnection *Database;
        TPopupMenu *PopupMenu1;
        TMenuItem *N1;
        TMenuItem *N2;
        TMenuItem *N3;
        TMenuItem *N4;
        TPopupMenu *FreeOtdel;
        TComboBox *Role;
        TListBox *Otdels;
        TListBox *Users;
        TButton *Button1;
        TEdit *Edit1;
        TPopupMenu *PopupMenu2;
        TMenuItem *N5;
        TImageList *ImageList1;
        TComboBox *Base;
        TMainMenu *MainMenu1;
        TMenuItem *N6;
        TLabel *Label3;
        TMenuItem *N7;
        void __fastcall FormDestroy(TObject *Sender);
        void __fastcall ListBox1Click(TObject *Sender);
        void __fastcall RoleClick(TObject *Sender);
        void __fastcall UsersClick(TObject *Sender);
        void __fastcall UsersMouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall N2Click(TObject *Sender);
        void __fastcall N4Click(TObject *Sender);
        void __fastcall N3Click(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall OtdelsMouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall N5Click(TObject *Sender);
        void __fastcall BaseClick(TObject *Sender);
        void __fastcall N6Click(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
private:	// User declarations

TLocateOptions SO;
void __fastcall ChangeCountClient(TObject *Sender);

void __fastcall MTIcon(TMessage&);
void __fastcall WMSysCommand(TMessage & Msg);

int OldRow;
String OldLog;
int EditLoginNumber;
vector<BaseItem>VBases;

int AnalizLic(String Text, String Pass);
void __fastcall TForm1::SelectOtdel(TObject *Sender);
int ReadRes(String Key1, String Key2);
int Propis(String S);
void VerifyLicense();
int LicCount;

public:		// User declarations
int Block;
String DatabasePatch;
mClients *MyClients;
Client *Client;
String Path;
String GetFileDatabase(String NameDatabase);
//String AdminDatabase;
bool MergeLogins();
long GetLicenseCount(String DBName);

void SaveCode(String Login, String Password);
        __fastcall TForm1(TComponent* Owner);

void UpdateTempLogin();
void UpdateOtdel(int NumLogin);
void Error();
bool Net_Error;

        BEGIN_MESSAGE_MAP
          VCL_MESSAGE_HANDLER(MyTrayIcon,TMessage,MTIcon);
          VCL_MESSAGE_HANDLER(WM_SYSCOMMAND, TMessage, WMSysCommand)
        END_MESSAGE_MAP(TComponent);

};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif

