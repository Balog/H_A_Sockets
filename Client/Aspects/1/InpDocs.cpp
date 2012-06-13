//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "InpDocs.h"
#include "Zastavka.h"
#include "FMoveAsp.h"
#include "MainForm.h"
#include "InputFiltr.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TInputDocs *InputDocs;
TTreeNode *SelNode1;
//---------------------------------------------------------------------------
__fastcall TInputDocs::TInputDocs(TComponent* Owner)
        : TForm(Owner)
{

}
//---------------------------------------------------------------------------
void __fastcall TInputDocs::FormShow(TObject *Sender)
{

Button1->Enabled=false;
NodeVector_3.clear();

/*
Mode
1-����������
2-��� ������������, ��������
3-������������� ������
4-�����������
5-�����������
*/
//Mode=Form1->Mode;
AnsiString SNode,SBranch;

switch (Mode)
{
 case 1:
 {
 SNode="����_5";
 SBranch="�����_5";
 break;
 }
 case 2:
 {
 SNode="����_6";
 SBranch="�����_6";
 break;
 }
 case 3:
 {
 SNode="����_7";
 SBranch="�����_7";
 break;
 }
 case 4:
 {
 SNode="����_3";
 SBranch="�����_3";
 break;
 }
 case 5:
 {
 SNode="����_4";
 SBranch="�����_4";
 break;
 }

}

Nodes->Active=false;
Nodes->CommandText="Select * From "+SNode;
Nodes->Connection=Zast->ADOConn;
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
 int Parent=Nodes->FieldByName("��������")->Value;
 if (Parent!=0)
 {
 Branches->Active=false;
 Branches->CommandText="Select * From "+SNode+" Where [����� ����]="+IntToStr(Parent);
 Branches->Connection=Zast->ADOConn;
 Branches->Active=true;
 if (Branches->RecordCount==0)
 {
  Nodes->Delete();
 }
// Nodes->Next();
 }
 Nodes->Next();
}

PMyNode  MyNodePtr;
MyNodePtr = new TMyNode;


TTreeNode *N;
TTreeNode *N1;
TTreeNode *N2;
int NumParent;
TreeView1->Items->BeginUpdate();
TreeView1->Items->Clear();


Nodes->Active=true;
Nodes->First();
NumParent=Nodes->FieldByName("����� ����")->Value;
AnsiString Text=Nodes->FieldByName("��������")->Value;

MyNodePtr->Number=NumParent;
MyNodePtr->Parent=0;
MyNodePtr->Node=true;

Nodes->Edit();
//N1=TreeView1->Items->Add(0,Text);
//N1=TreeView1->Items->AddChildObject(NULL,Text+" "+IntToStr(MyNodePtr->Number)+" "+IntToStr(MyNodePtr->Parent)+" ����",MyNodePtr);
N1=TreeView1->Items->AddChildObject(NULL,Text,MyNodePtr);

Nodes->Post();

Branches->Active=false;
Branches->CommandText="Select * From "+SBranch+" Where [����� ��������]="+IntToStr(NumParent);
Branches->Active=true;
Branches->First();
for (int j=0; j<Branches->RecordCount;j++)
{
int NumBranches=Branches->FieldByName("����� �����")->Value;
int BranchesParent=Branches->FieldByName("����� ��������")->Value;
AnsiString TextBranches=Branches->FieldByName("��������")->Value;
bool View=Branches->FieldByName("�����")->Value;
MyNodePtr = new TMyNode;
MyNodePtr->Number=NumBranches;
MyNodePtr->Parent=BranchesParent;
MyNodePtr->Node=false;
//N2=TreeView1->Items->AddChildObject(N1,TextBranches+" "+IntToStr(MyNodePtr->Number)+" "+IntToStr(MyNodePtr->Parent)+" �����",MyNodePtr);
if (View==true)
{
TreeView1->Items->AddChildObject(N1,TextBranches,MyNodePtr);
}

Branches->Edit();

Branches->Post();
Branches->Next();
}



Nodes->Active=false;
Nodes->CommandText="Select * From "+SNode+" Where ([��������])>0";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
int NumNode=Nodes->FieldByName("����� ����")->Value;
NumParent=Nodes->FieldByName("��������")->Value;
AnsiString TextNode=Nodes->FieldByName("��������")->Value;
TTreeNode *ParNod;
ParNod=N1;
for (unsigned int k=0; k<NodeVector_3.size();k++)
{
 if (NodeVector_3[k].Number==NumParent)
 {
  ParNod=NodeVector_3[k].Nod;
 }
}

MyNodePtr = new TMyNode;
MyNodePtr->Number=NumNode;
MyNodePtr->Parent=NumParent;
MyNodePtr->Node=true;
//N=TreeView1->Items->AddChildObject(ParNod,TextNode+" "+IntToStr(MyNodePtr->Number)+" "+IntToStr(MyNodePtr->Parent)+" ����",MyNodePtr);
N=TreeView1->Items->AddChildObject(ParNod,TextNode,MyNodePtr);

//N=TreeView1->Items->AddChild(ParNod,TextNode);
//Nodes->Edit();

//Nodes->Post();
Nod NN;
NN.Number=NumNode;
NN.Parent=NumParent;
NN.Nod=N;
It_Node_3=NodeVector_3.end();
NodeVector_3.insert(It_Node_3,NN);


Branches->Active=false;
Branches->CommandText="Select * From "+SBranch+" Where [����� ��������]="+IntToStr(NumNode);
Branches->Active=true;
Branches->First();
for (int j=0; j<Branches->RecordCount;j++)
{
int NumBranches=Branches->FieldByName("����� �����")->Value;
int BranchesParent=Branches->FieldByName("����� ��������")->Value;
AnsiString TextBranches=Branches->FieldByName("��������")->Value;


MyNodePtr = new TMyNode;
MyNodePtr->Number=NumBranches;
MyNodePtr->Parent=BranchesParent;
MyNodePtr->Node=false;
//N2=TreeView1->Items->AddChildObject(N,TextBranches+" "+IntToStr(MyNodePtr->Number)+" "+IntToStr(MyNodePtr->Parent)+" �����",MyNodePtr);
TreeView1->Items->AddChildObject(N,TextBranches,MyNodePtr);

Branches->Edit();

Branches->Post();
Branches->Next();
}
Nodes->Next();
}
Nodes->Active=false;
Branches->Active=false;


TreeView1->Items->EndUpdate();
TreeView1->Items->Item[0]->Expand(false);
/*
for(int i=0;i<TreeView1->Items->Count;i++)
{
TreeView1->Items->Item[i]->Expand(true);
}

*/
}
//---------------------------------------------------------------------------


void __fastcall TInputDocs::TreeView1MouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
THitTests HT;
HT=TreeView1->GetHitTestInfoAt(X,Y);
if (HT.Contains(htOnItem))
{

SelNode1=TreeView1->Selected;
if(PMyNod(SelNode1->Data)->Node==false)
{
Button1->Enabled=true;
}
else
{
Button1->Enabled=false;
}
//ShowMessage(SelNode->Text);


//XX=X;
//YY=Y;

}
//PMyNode(SelNode->Data)->Node
//PMyNod(SelNode->Data)->Node
}
//---------------------------------------------------------------------------

void __fastcall TInputDocs::TreeView1DblClick(TObject *Sender)
{
try
{
Button1->Click();
}
catch(EAccessViolation&)
{
}
}
//---------------------------------------------------------------------------

void __fastcall TInputDocs::Button1Click(TObject *Sender)
{
if(PMyNod(SelNode1->Data)->Node==false)
{

TextBr=SelNode1->Text;
NumBr=PMyNod(SelNode1->Data)->Number;

if(IForm==1)
{
switch (Mode)
{
 case 1:
 {

 Form1->InpTer();

 break;
 }
 case 2:
 {
 Form1->InpDeyat();

 break;
 }
 case 3:
 {
 Form1->InpAsp();

 break;
 }
 case 4:
 {
 Form1->InpVozd();

 break;
 }
 case 5:
 {
 Form1->InpMeropr();

 break;
 }
}

}

if(IForm==2)
{
//-----
 MAsp->Index=InputDocs->NumBr;
switch (Mode)
{
 case 1:
 {
 MAsp->ComboBox4->Text=InputDocs->TextBr;

 MAsp->InpTer();
 break;
 }
 case 2:
 {
  MAsp->ComboBox5->Text=InputDocs->TextBr;
 MAsp->InpDeyat();
 break;
 }
 case 3:
 {
  MAsp->ComboBox6->Text=InputDocs->TextBr;
 MAsp->InpAsp();

 break;
 }
 case 4:
 {
  MAsp->ComboBox7->Text=InputDocs->TextBr;
 MAsp->InpVozd();
 break;
 }
}
}
if(IForm==3)
{
//-----
switch (Mode)
{
 case 1:
 {
 Filter->InpTer();
 break;
 }
 case 2:
 {
 Filter->InpDeyat();
 break;
 }
 case 3:
 {
 Filter->InpAsp();

 break;
 }
 case 4:
 {
 Filter->InpVozd();
 break;
 }
}
}
//-----

this->Close();
}

}
//---------------------------------------------------------------------------

void __fastcall TInputDocs::Button2Click(TObject *Sender)
{
this->Close();
}
//---------------------------------------------------------------------------


void __fastcall TInputDocs::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
Application->HelpJump("IDH_�����_������_������_��_�����������");
}
}
//---------------------------------------------------------------------------
