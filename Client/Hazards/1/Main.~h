//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ADODB.hpp>
#include <DB.hpp>
#include <ExtCtrls.hpp>
#include <Menus.hpp>
#include <ScktComp.hpp>
#include "ClientClass.h"
#include <ActnList.hpp>
#include <ActnMan.hpp>
#include <ComCtrls.hpp>
#include <DBActns.hpp>
#include <DBCtrls.hpp>
#include <DBGrids.hpp>
#include <Grids.hpp>
#include <vector>
#include <deque>
#include "MDBConnector.h"
#ifndef MainH
#define MainH
//---------------------------------------------------------------------------
typedef struct MyNode
{
int Number;
int Parent;
 bool Node;
} TMyNode;
typedef TMyNode* PMyNode;
//---------------------------------------------------------------------------
using namespace std;
//---------------------------------------------------------------------------
struct Node
{
 int Number;
 int Parent;

 TTreeNode *Nod;
};
//---------------------------------------------------------------------------
struct Str_RW
{
String NameAction;
String Text;
int Num;
};
//----------------------------------------------------------------------------

class TDocuments : public TForm
{
__published:	// IDE-managed Components
        TPageControl *PageControl1;
        TTabSheet *TabMetod;
        TDBMemo *DBMemo1;
        TPanel *Panel1;
        TCheckBox *ReadMetod;
        TButton *Button1;
        TTabSheet *TabPodr;
        TDBGrid *DBGrid2;
        TPanel *Panel2;
        TButton *Button2;
        TTabSheet *TabCryt;
        TDBGrid *DBGrid1;
        TPanel *Panel3;
        TButton *Button3;
        TTabSheet *TabSit;
        TDBGrid *DBGrid3;
        TPanel *Panel9;
        TButton *Button9;
        TTabSheet *TabVozd;
        TTreeView *TreeView1;
        TPanel *Panel4;
        TButton *Button4;
        TTabSheet *TabMeropr;
        TTreeView *TreeView2;
        TPanel *Panel5;
        TButton *Button5;
        TTabSheet *TabTerr;
        TTreeView *TreeView3;
        TPanel *Panel6;
        TButton *Button6;
        TTabSheet *TabDeyat;
        TTreeView *TreeView4;
        TPanel *Panel7;
        TButton *Button7;
        TTabSheet *TabAspects;
        TTreeView *TreeView5;
        TPanel *Panel8;
        TButton *Button8;
        TMainMenu *MainMenu1;
        TMenuItem *N4;
        TMenuItem *N2;
        TMenuItem *N27;
        TMenuItem *N26;
        TMenuItem *N28;
        TMenuItem *N6;
        TMenuItem *N20;
        TMenuItem *N24;
        TMenuItem *N25;
        TMenuItem *N29;
        TMenuItem *N31;
        TMenuItem *N5;
        TMenuItem *N18;
        TMenuItem *N38;
        TMenuItem *N37;
        TMenuItem *N39;
        TMenuItem *N32;
        TMenuItem *N34;
        TMenuItem *N35;
        TMenuItem *N36;
        TMenuItem *N40;
        TMenuItem *N41;
        TMenuItem *N19;
        TMenuItem *N21;
        TMenuItem *N00111;
        TMenuItem *N00121;
        TMenuItem *N22;
        TMenuItem *N1;
        TMenuItem *N23;
        TMenuItem *N8;
        TADODataSet *Nodes;
        TADODataSet *Branches;
        TADOCommand *Comm;
        TPopupMenu *PopupMenu5;
        TMenuItem *N16;
        TMenuItem *N17;
        TPopupMenu *PopupMenu4;
        TMenuItem *N14;
        TMenuItem *N15;
        TPopupMenu *PopupMenu3;
        TMenuItem *N12;
        TMenuItem *N13;
        TPopupMenu *PopupMenu2;
        TMenuItem *N10;
        TMenuItem *N11;
        TPopupMenu *PopupMenu1;
        TMenuItem *N3;
        TMenuItem *N9;
        TADODataSet *Podr;
        TDataSource *DataSource2;
        TPopupMenu *PopupMenu7;
        TMenuItem *MenuItem3;
        TMenuItem *MenuItem4;
        TDataSource *DataSource1;
        TADODataSet *ADODataSet1;
        TADODataSet *Sit;
        TDataSource *DataSource3;
        TADODataSet *Metod;
        TPopupMenu *PopupMenu6;
        TMenuItem *MenuItem1;
        TMenuItem *MenuItem2;
        TPopupMenu *PopupMenu8;
        TMenuItem *MenuItem5;
        TMenuItem *MenuItem6;
        TDataSource *DataSource5;
        TActionManager *ActionManager1;
        TDataSetPost *DataSetPost1;
        TDataSetPost *DataSetPost2;
        TDataSetPost *DataSetPost3;
        TDataSetPost *DataSetPost4;
        TDataSetRefresh *DataSetRefresh1;
        TDataSetRefresh *DataSetRefresh2;
        TDataSetRefresh *DataSetRefresh3;
        TDataSetRefresh *DataSetRefresh4;
        TADODataSet *Temp;
        TADODataSet *ADODataSet2;
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall N9Click(TObject *Sender);
        void __fastcall N3Click(TObject *Sender);
        void __fastcall N10Click(TObject *Sender);
        void __fastcall N11Click(TObject *Sender);
        void __fastcall N12Click(TObject *Sender);
        void __fastcall N13Click(TObject *Sender);
        void __fastcall N14Click(TObject *Sender);
        void __fastcall N15Click(TObject *Sender);
        void __fastcall N16Click(TObject *Sender);
        void __fastcall N17Click(TObject *Sender);
        void __fastcall MenuItem3Click(TObject *Sender);
        void __fastcall MenuItem4Click(TObject *Sender);
        void __fastcall MenuItem1Click(TObject *Sender);
        void __fastcall MenuItem2Click(TObject *Sender);
        void __fastcall MenuItem5Click(TObject *Sender);
        void __fastcall MenuItem6Click(TObject *Sender);
        void __fastcall ReadMetodClick(TObject *Sender);
        void __fastcall DBGrid1Exit(TObject *Sender);
        void __fastcall ADODataSet1AfterInsert(TDataSet *DataSet);
        void __fastcall ADODataSet1AfterPost(TDataSet *DataSet);
        void __fastcall ADODataSet1BeforePost(TDataSet *DataSet);
        void __fastcall TreeView1Edited(TObject *Sender, TTreeNode *Node,
          AnsiString &S);
        void __fastcall TreeView1MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall TreeView2Edited(TObject *Sender, TTreeNode *Node,
          AnsiString &S);
        void __fastcall TreeView2MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall TreeView3MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall TreeView4Edited(TObject *Sender, TTreeNode *Node,
          AnsiString &S);
        void __fastcall TreeView4MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall TreeView5Edited(TObject *Sender, TTreeNode *Node,
          AnsiString &S);
        void __fastcall TreeView5MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall TreeView3Edited(TObject *Sender, TTreeNode *Node,
          AnsiString &S);
        void __fastcall N8Click(TObject *Sender);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall N2Click(TObject *Sender);
        void __fastcall N27Click(TObject *Sender);
        void __fastcall N26Click(TObject *Sender);
        void __fastcall N28Click(TObject *Sender);
        void __fastcall N6Click(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
        void __fastcall N20Click(TObject *Sender);
        void __fastcall N24Click(TObject *Sender);
        void __fastcall N25Click(TObject *Sender);
        void __fastcall N31Click(TObject *Sender);
        void __fastcall N18Click(TObject *Sender);
        void __fastcall DBMemo1Exit(TObject *Sender);
        void __fastcall N38Click(TObject *Sender);
        void __fastcall DBGrid2Exit(TObject *Sender);
        void __fastcall DBGrid3Exit(TObject *Sender);
        void __fastcall N37Click(TObject *Sender);
        void __fastcall N39Click(TObject *Sender);
        void __fastcall N32Click(TObject *Sender);
        void __fastcall N33Click(TObject *Sender);
        void __fastcall N34Click(TObject *Sender);
        void __fastcall N35Click(TObject *Sender);
        void __fastcall N36Click(TObject *Sender);
        void __fastcall N41Click(TObject *Sender);
        void __fastcall N00111Click(TObject *Sender);
        void __fastcall N00121Click(TObject *Sender);
        void __fastcall N22Click(TObject *Sender);
        void __fastcall N19Click(TObject *Sender);
        void __fastcall N23Click(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
        void __fastcall Button3Click(TObject *Sender);
        void __fastcall Button9Click(TObject *Sender);
        void __fastcall Button4Click(TObject *Sender);
        void __fastcall Button5Click(TObject *Sender);
        void __fastcall Button6Click(TObject *Sender);
        void __fastcall Button7Click(TObject *Sender);
        void __fastcall Button8Click(TObject *Sender);
private:	// User declarations
String Path;
String AdminDatabase;
String DiaryDatabase;
int LicCount;

vector<String>Parameters;

String GetIP();

String ServerName;


int DeleteNode(String NameNode, String NameBranch, int Number);



public:		// User declarations
        __fastcall TDocuments(TComponent* Owner);
bool Reg;
void Initialize();
bool C;
bool Ins;
bool Kriteriy;
void Crit(TDataSet *DataSet);
void Crit1(TDataSet *DataSet);
void Crit2(TDataSet *DataSet);

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

vector<Str_RW>ReadWrite;
vector<Str_RW>::iterator RW;

void MergeNode(String NodeTable);
void MergeBranch(String BranchTable);

void LoadTab1();
void LoadTab2();
void LoadTab3();
void LoadTab4();
void LoadTab5();

        BEGIN_MESSAGE_MAP
          VCL_MESSAGE_HANDLER(WM_SYSCOMMAND, TMessage, WMSysCommand)
        END_MESSAGE_MAP(TComponent);
void __fastcall WMSysCommand(TMessage & Msg);        
};
//---------------------------------------------------------------------------
extern PACKAGE TDocuments *Documents;
//---------------------------------------------------------------------------
#endif
