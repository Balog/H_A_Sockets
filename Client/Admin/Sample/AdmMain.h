//---------------------------------------------------------------------------

#ifndef AdmMainH
#define AdmMainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ADODB.hpp>
#include <DB.hpp>
//#include "ServerClasses.h"
#include <Menus.hpp>
#include "MDBConnector.h"
#include "ClientClass.h"
#include <ExtCtrls.hpp>
#include <FileCtrl.hpp>
#include <vector>
using namespace std;

struct CDBItem
{
int Num;
String Name;
String ServerDB;
int NumDatabase;
};
//---------------------------------------------------------------------------
class TMain : public TForm
{
__published:	// IDE-managed Components
        TComboBox *Role;
        TListBox *Users;
        TListBox *Otdels;
        TADOConnection *Database1;
        TPopupMenu *PopupMenu1;
        TMenuItem *N1;
        TMenuItem *N2;
        TMenuItem *N3;
        TMenuItem *N4;
        TPopupMenu *FreeOtdel;
        TLabel *Label1;
        TLabel *Label2;
        TMainMenu *MainMenu1;
        TMenuItem *N5;
        TMenuItem *N6;
        TComboBox *CBDatabase;
        TLabel *Label3;
        TLabel *Label4;
        TTimer *DBTimer;
        TMenuItem *N7;
        TMenuItem *N8;
        TMenuItem *N9;
        TMenuItem *N10;
        void __fastcall RoleClick(TObject *Sender);
        void __fastcall UsersClick(TObject *Sender);
        void __fastcall UsersMouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall OtdelsMouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall N2Click(TObject *Sender);
        void __fastcall N3Click(TObject *Sender);
        void __fastcall N4Click(TObject *Sender);
        void __fastcall N5Click(TObject *Sender);
        void __fastcall N6Click(TObject *Sender);
        void __fastcall CBDatabaseClick(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall FormDestroy(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
        void __fastcall N8Click(TObject *Sender);
        void __fastcall N9Click(TObject *Sender);
        void __fastcall FormKeyUp(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall N10Click(TObject *Sender);
private:	// User declarations
TLocateOptions SO;
void UpdateTempLogin();
void UpdateOtdel(int NumLogin);
//mClients *MyClients;
String DatabasePatch;

String Path;
String AdminDatabase;
int LicCount;
void VerifyLicense();

//Form *MForm;
//long IDDB;
long MyTableID;
int EditLoginNumber;


void __fastcall TMain::SelectOtdel(TObject *Sender);

void LoadTable(Table* FromCopy, TADODataSet* ToCopy);
void LoadTable(TADODataSet* FromCopy, Table* ToCopy);

//0 - ��� ���������
//1 - �� ��������� ���������� �����
//2 - �� �������� ����� �������
//3 - �� ��������� ����� �����
int VerifyTable(TADODataSet* LTable, Table* RTable);
int VerifyTable(Table* LTable, TADODataSet* RTable);


void MergeLogins();

void MergeOtdels();

String UpLoadLogins();
bool VerifyUpdates();
void UpdateOtdels();

public:		// User declarations
bool Reg;
MDBConnector* Database;
vector<CDBItem>VDB;
String ServerName;
Client *MClient;
        __fastcall TMain(TComponent* Owner);
void SaveCode(String Login, String Password);
String LoadLogins();
bool Start;

int GetIDDBName(String Name);

bool InputPass;
};
//---------------------------------------------------------------------------
extern PACKAGE TMain *Main;
//---------------------------------------------------------------------------
#endif
