//---------------------------------------------------------------------------

#ifndef InpDocsH
#define InpDocsH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
//#include "Docs.h"
#include "Main.h"
//#include "MainForm.h"
#include "Zastavka.h"
#include <ADODB.hpp>
#include <DB.hpp>
#include <deque>
typedef struct MyNod
{
//  AnsiString FName, LName;
int Number;
int Parent;
 bool Node;
} TMyNod;
typedef TMyNod* PMyNod;
using namespace std;
struct Nod
{
 int Number;
 int Parent;

 TTreeNode *Nod;
};
//---------------------------------------------------------------------------
class TInputDocs : public TForm
{
__published:	// IDE-managed Components
        TTreeView *TreeView1;
        TButton *Button1;
        TButton *Button2;
        TADODataSet *Nodes;
        TADODataSet *Branches;
        void __fastcall FormShow(TObject *Sender);
        void __fastcall TreeView1MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
        void __fastcall TreeView1DblClick(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall Button2Click(TObject *Sender);
        void __fastcall FormKeyUp(TObject *Sender, WORD &Key,
          TShiftState Shift);
private:	// User declarations
public:		// User declarations
        __fastcall TInputDocs(TComponent* Owner);
/*
Mode
1-����������
2-��� ������������, ��������
3-������������� ������
4-�����������
*/
int Mode;
        deque<Nod> NodeVector_3;
        deque<Nod>::iterator It_Node_3;

AnsiString TextBr;
int NumBr;
/*
IForm=1 (MainForm)
IForm=2 (InputFiltr)
*/
int IForm;
};
//---------------------------------------------------------------------------
extern PACKAGE TInputDocs *InputDocs;
//---------------------------------------------------------------------------
#endif