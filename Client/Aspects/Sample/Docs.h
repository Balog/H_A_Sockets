//---------------------------------------------------------------------------

#ifndef DocsH
#define DocsH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Menus.hpp>
//#include "MainForm.h"
//#include "F_Vvedenie.h"
//#include "Metod.h"
#include "Zastavka.h"
#include <ComCtrls.hpp>
#include <ADODB.hpp>
#include <DB.hpp>
#include <DBGrids.hpp>
#include <Grids.hpp>
#include <DBCtrls.hpp>
#include <Mask.hpp>
#include <ExtCtrls.hpp>
#include <ActnList.hpp>
#include <ActnMan.hpp>
#include <DBActns.hpp>
#include <deque>
//#include "Form_SendFile.h"
typedef struct MyNode
{
//  AnsiString FName, LName;
int Number;
int Parent;
 bool Node;
} TMyNode;
typedef TMyNode* PMyNode;
using namespace std;
struct Node
{
 int Number;
 int Parent;

 TTreeNode *Nod;
};
//---------------------------------------------------------------------------
class TDocuments : public TForm
{
__published:	// IDE-managed Components
        TMainMenu *MainMenu1;
        TMenuItem *N1;
        TMenuItem *N8;
        TPageControl *PageControl1;
        TTabSheet *TabVozd;
        TTabSheet *TabMeropr;
        TTabSheet *TabTerr;
        TTabSheet *TabDeyat;
        TTabSheet *TabAspects;
        TTreeView *TreeView1;
        TTreeView *TreeView2;
        TTreeView *TreeView3;
        TTreeView *TreeView4;
        TTreeView *TreeView5;
        TADODataSet *Nodes;
        TADODataSet *Branches;
        TPopupMenu *PopupMenu1;
        TMenuItem *N3;
        TMenuItem *N9;
        TADOCommand *Comm;
        TPopupMenu *PopupMenu2;
        TPopupMenu *PopupMenu3;
        TPopupMenu *PopupMenu4;
        TPopupMenu *PopupMenu5;
        TMenuItem *N10;
        TMenuItem *N11;
        TMenuItem *N12;
        TMenuItem *N13;
        TMenuItem *N14;
        TMenuItem *N15;
        TMenuItem *N16;
        TMenuItem *N17;
        TTabSheet *TabCryt;
        TTabSheet *TabPodr;
        TDBGrid *DBGrid1;
        TDataSource *DataSource1;
        TADODataSet *ADODataSet1;
        TPopupMenu *PopupMenu6;
        TMenuItem *MenuItem1;
        TMenuItem *MenuItem2;
        TADODataSet *ADODataSet2;
        TDBGrid *DBGrid2;
        TTabSheet *TabSit;
        TDBGrid *DBGrid3;
        TADODataSet *Podr;
        TDataSource *DataSource2;
        TPopupMenu *PopupMenu7;
        TMenuItem *MenuItem3;
        TMenuItem *MenuItem4;
        TADODataSet *Sit;
        TDataSource *DataSource3;
        TPopupMenu *PopupMenu8;
        TMenuItem *MenuItem5;
        TMenuItem *MenuItem6;
        TADODataSet *Temp;
        TMenuItem *N4;
        TMenuItem *N5;
        TMenuItem *N6;
        TMenuItem *N7;
        TMenuItem *N20;
        TMenuItem *N24;
        TMenuItem *N25;
        TMenuItem *N26;
        TMenuItem *N27;
        TMenuItem *N28;
        TMenuItem *N29;
        TMenuItem *N31;
        TMenuItem *N32;
        TMenuItem *N33;
        TMenuItem *N34;
        TMenuItem *N35;
        TMenuItem *N36;
        TMenuItem *N37;
        TMenuItem *N38;
        TMenuItem *N39;
        TMenuItem *N40;
        TMenuItem *N41;
        TTabSheet *TabMetod;
        TMenuItem *N2;
        TMenuItem *N18;
        TDBMemo *DBMemo1;
        TADODataSet *Metod;
        TDataSource *DataSource5;
        TPanel *Panel1;
        TCheckBox *ReadMetod;
        TActionManager *ActionManager1;
        TDataSetPost *DataSetPost1;
        TDataSetPost *DataSetPost2;
        TDataSetPost *DataSetPost3;
        TDataSetPost *DataSetPost4;
        TButton *Button1;
        TPanel *Panel2;
        TButton *Button2;
        TPanel *Panel3;
        TButton *Button3;
        TPanel *Panel4;
        TButton *Button4;
        TPanel *Panel5;
        TButton *Button5;
        TPanel *Panel6;
        TButton *Button6;
        TPanel *Panel7;
        TButton *Button7;
        TPanel *Panel8;
        TButton *Button8;
        TPanel *Panel9;
        TButton *Button9;
        TMenuItem *N19;
        TMenuItem *N21;
        TMenuItem *N00111;
        TMenuItem *N00121;
        TMenuItem *N22;
        TMenuItem *N23;
        TDataSetRefresh *DataSetRefresh1;
        TDataSetRefresh *DataSetRefresh2;
        TDataSetRefresh *DataSetRefresh3;
        TDataSetRefresh *DataSetRefresh4;
        void __fastcall N8Click(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall TreeView1Edited(TObject *Sender, TTreeNode *Node,
          AnsiString &S);
        void __fastcall N3Click(TObject *Sender);
        void __fastcall N9Click(TObject *Sender);
        void __fastcall TreeView1MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall TreeView2Edited(TObject *Sender, TTreeNode *Node,
          AnsiString &S);
        void __fastcall TreeView2MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall N10Click(TObject *Sender);
        void __fastcall N11Click(TObject *Sender);
        void __fastcall TreeView3Edited(TObject *Sender, TTreeNode *Node,
          AnsiString &S);
        void __fastcall PageControl1MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall N12Click(TObject *Sender);
        void __fastcall TreeView3MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall N13Click(TObject *Sender);
        void __fastcall TreeView4Edited(TObject *Sender, TTreeNode *Node,
          AnsiString &S);
        void __fastcall TreeView4MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall N14Click(TObject *Sender);
        void __fastcall N15Click(TObject *Sender);
        void __fastcall TreeView5Edited(TObject *Sender, TTreeNode *Node,
          AnsiString &S);
        void __fastcall TreeView5MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall N16Click(TObject *Sender);
        void __fastcall N17Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
//        void __fastcall N22Click(TObject *Sender);
//        void __fastcall N23Click(TObject *Sender);
        void __fastcall FormActivate(TObject *Sender);
        void __fastcall ADODataSet1BeforePost(TDataSet *DataSet);
        void __fastcall ADODataSet1AfterPost(TDataSet *DataSet);
        void __fastcall MenuItem1Click(TObject *Sender);
        void __fastcall MenuItem2Click(TObject *Sender);
        void __fastcall PopupMenu6Popup(TObject *Sender);
        void __fastcall ADODataSet1AfterInsert(TDataSet *DataSet);
        void __fastcall MenuItem3Click(TObject *Sender);
        void __fastcall MenuItem4Click(TObject *Sender);
        void __fastcall MenuItem5Click(TObject *Sender);
        void __fastcall MenuItem6Click(TObject *Sender);
        void __fastcall N6Click(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
        void __fastcall N20Click(TObject *Sender);
        void __fastcall N24Click(TObject *Sender);
        void __fastcall N25Click(TObject *Sender);
        void __fastcall N26Click(TObject *Sender);
        void __fastcall N27Click(TObject *Sender);
        void __fastcall N28Click(TObject *Sender);
        void __fastcall DBMemo1Exit(TObject *Sender);
        void __fastcall DBGrid1Exit(TObject *Sender);
        void __fastcall ReadMetodClick(TObject *Sender);
        void __fastcall N2Click(TObject *Sender);
        void __fastcall N32Click(TObject *Sender);
        void __fastcall N33Click(TObject *Sender);
        void __fastcall N34Click(TObject *Sender);
        void __fastcall N35Click(TObject *Sender);
        void __fastcall N36Click(TObject *Sender);
        void __fastcall N37Click(TObject *Sender);
        void __fastcall N18Click(TObject *Sender);
        void __fastcall N38Click(TObject *Sender);
        void __fastcall N39Click(TObject *Sender);
        void __fastcall N31Click(TObject *Sender);
        void __fastcall N41Click(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
        void __fastcall Button3Click(TObject *Sender);
        void __fastcall Button4Click(TObject *Sender);
        void __fastcall Button5Click(TObject *Sender);
        void __fastcall Button6Click(TObject *Sender);
        void __fastcall Button7Click(TObject *Sender);
        void __fastcall Button8Click(TObject *Sender);
        void __fastcall Button9Click(TObject *Sender);
        void __fastcall N19Click(TObject *Sender);
        void __fastcall N00111Click(TObject *Sender);
        void __fastcall N00121Click(TObject *Sender);
        void __fastcall N22Click(TObject *Sender);
        void __fastcall N23Click(TObject *Sender);
        void __fastcall FormKeyUp(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall N1Click(TObject *Sender);
private:	// User declarations

TLocateOptions SO;
void LoadTab1();
void LoadTab2();
void LoadTab3();
void LoadTab4();
void LoadTab5();

bool LoadNodeBranch(String NodeTable, String BranchTable, bool Pr);
void MergeNode(String NodeTable);
void MergeBranch(String BranchTable);

bool LoadCrit(bool Pr);
bool LoadPodr(bool Pr);
bool LoadSit(bool Pr);
bool LoadMet(bool Pr);

bool UploadNodeBranch(String NodeTable, String BranchTable, bool Pr);
void PrepareMet();

String ReadVidVozd(bool Pr);
String ReadMeropr(bool Pr);
String ReadTerr(bool Pr);
String ReadDeyat(bool Pr);
String ReadAspect(bool Pr);
String ReadCrit(bool Pr);
String ReadPodr(bool Pr);
String ReadSit(bool Pr);
String ReadMet(bool Pr);

String WriteVozd(bool Pr);
String WriteMeropr(bool Pr);
String WriteTerr(bool Pr);
String WriteDeyat(bool Pr);
String WriteAspect(bool Pr);
String WriteCryt(bool Pr);
String WritePodr(bool Pr);
String WriteSit(bool Pr);
String WriteMet(bool Pr);

void LoadDirVozd();
void LoadDirTerr();
void LoadDirDeyat();
void LoadDirAspect();

void SaveAllTables(bool Full);


void MergeAspects();
void PrepareMergeAspects();
//bool VerifyAspectsContent(bool New);

int DeleteNode(String NameNode, String NameBranch, int Number);

public:		// User declarations
        __fastcall TDocuments(TComponent* Owner);
        deque<Node> NodeVector_3;
        deque<Node>::iterator It_Node_3;
        deque<Node> NodeVector_4;
        deque<Node>::iterator It_Node_4;
        deque<Node> NodeVector_5;
        deque<Node>::iterator It_Node_5;
        deque<Node> NodeVector_6;
        deque<Node>::iterator It_Node_6;
        deque<Node> NodeVector_7;
        deque<Node>::iterator It_Node_7;

//Критерии
void Crit(TDataSet *DataSet);
void Crit1(TDataSet *DataSet);
void Crit2(TDataSet *DataSet);
bool C;
bool Ins;
bool Kriteriy;
//////////
//***
void Initialize();
//***
String LoadAspects();

        BEGIN_MESSAGE_MAP
          VCL_MESSAGE_HANDLER(WM_SYSCOMMAND, TMessage, WMSysCommand)
        END_MESSAGE_MAP(TComponent);
void __fastcall WMSysCommand(TMessage & Msg);  
};
//---------------------------------------------------------------------------
extern PACKAGE TDocuments *Documents;
//---------------------------------------------------------------------------
#endif
