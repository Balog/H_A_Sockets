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
1-Территория
2-Вид деятельности, операция
3-Экологический аспект
4-Воздействие
5-Мероприятие
*/
//Mode=Form1->Mode;
AnsiString SNode,SBranch;

switch (Mode)
{
 case 1:
 {
 SNode="Узлы_5";
 SBranch="Ветви_5";
 break;
 }
 case 2:
 {
 SNode="Узлы_6";
 SBranch="Ветви_6";
 break;
 }
 case 3:
 {
 SNode="Узлы_7";
 SBranch="Ветви_7";
 break;
 }
 case 4:
 {
 SNode="Узлы_3";
 SBranch="Ветви_3";
 break;
 }
 case 5:
 {
 SNode="Узлы_4";
 SBranch="Ветви_4";
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
 int Parent=Nodes->FieldByName("Родитель")->Value;
 if (Parent!=0)
 {
 Branches->Active=false;
 Branches->CommandText="Select * From "+SNode+" Where [Номер узла]="+IntToStr(Parent);
 Branches->Connection=Zast->ADOConn;
 Branches->Active=true;
 if (Branches->RecordCount==0)
 {
  Nodes->Delete();
 }
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
NumParent=Nodes->FieldByName("Номер узла")->Value;
AnsiString Text=Nodes->FieldByName("Название")->Value;

MyNodePtr->Number=NumParent;
MyNodePtr->Parent=0;
MyNodePtr->Node=true;

Nodes->Edit();
N1=TreeView1->Items->AddChildObject(NULL,Text,MyNodePtr);

Nodes->Post();

Branches->Active=false;
Branches->CommandText="Select * From "+SBranch+" Where [Номер родителя]="+IntToStr(NumParent);
Branches->Active=true;
Branches->First();
for (int j=0; j<Branches->RecordCount;j++)
{
int NumBranches=Branches->FieldByName("Номер ветви")->Value;
int BranchesParent=Branches->FieldByName("Номер родителя")->Value;
AnsiString TextBranches=Branches->FieldByName("Название")->Value;
bool View=Branches->FieldByName("Показ")->Value;
MyNodePtr = new TMyNode;
MyNodePtr->Number=NumBranches;
MyNodePtr->Parent=BranchesParent;
MyNodePtr->Node=false;
if (View==true)
{
TreeView1->Items->AddChildObject(N1,TextBranches,MyNodePtr);
}

Branches->Edit();

Branches->Post();
Branches->Next();
}



Nodes->Active=false;
Nodes->CommandText="Select * From "+SNode+" Where ([Родитель])>0";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
int NumNode=Nodes->FieldByName("Номер узла")->Value;
NumParent=Nodes->FieldByName("Родитель")->Value;
AnsiString TextNode=Nodes->FieldByName("Название")->Value;
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
N=TreeView1->Items->AddChildObject(ParNod,TextNode,MyNodePtr);
Nod NN;
NN.Number=NumNode;
NN.Parent=NumParent;
NN.Nod=N;
It_Node_3=NodeVector_3.end();
NodeVector_3.insert(It_Node_3,NN);


Branches->Active=false;
Branches->CommandText="Select * From "+SBranch+" Where [Номер родителя]="+IntToStr(NumNode);
Branches->Active=true;
Branches->First();
for (int j=0; j<Branches->RecordCount;j++)
{
int NumBranches=Branches->FieldByName("Номер ветви")->Value;
int BranchesParent=Branches->FieldByName("Номер родителя")->Value;
AnsiString TextBranches=Branches->FieldByName("Название")->Value;


MyNodePtr = new TMyNode;
MyNodePtr->Number=NumBranches;
MyNodePtr->Parent=BranchesParent;
MyNodePtr->Node=false;
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
}
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
Application->HelpJump("IDH_ФОРМА_ВЫБОРА_ДАННЫХ_ИЗ_СПРАВОЧНИКА");
}
}
//---------------------------------------------------------------------------

