//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Docs.h"
#include "About.h"
//#include "LoadKeyFile.h"
//#include "Svod.h"
//#include "File_operations.cpp"
//#include "SetZn.h"
#include "Password.h"
#include "MasterPointer.h"
#include "Progress.h"
#include "MoveAspects.h"
#include "FMoveAsp.h"
#include "Rep1.h"
#include "Svod.h"
#include "About.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TDocuments *Documents;
int XX,YY;
TTreeNode *SelNode;
PMyNode  MyNodePtr;
//---------------------------------------------------------------------------
__fastcall TDocuments::TDocuments(TComponent* Owner)
        : TForm(Owner)
{
//TreeView1->RightClickSelect=true;
Nodes->Connection=Zast->ADOConn;
Branches->Connection=Zast->ADOConn;
Comm->Connection=Zast->ADOConn;
LoadTab1();
LoadTab2();
LoadTab3();
LoadTab4();
LoadTab5();
Zast->Timer1->Enabled=true;

///
C=false;
Ins=false;

ADODataSet1->Connection=Zast->ADOConn;
ADODataSet2->Connection=Zast->ADOConn;
///
//***
Podr->Connection=Zast->ADOAspect;
Sit->Connection=Zast->ADOConn;
Temp->Connection=Zast->ADOConn;
Metod->Connection=Zast->ADOConn;


Metod->CommandText="Select * From ��������";

Podr->Active=true;
Sit->Active=true;

//***
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::N8Click(TObject *Sender)
{
//Zast->Close();
this->Close();
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::FormClose(TObject *Sender,
      TCloseAction &Action)
{
DataSetRefresh1->Execute();
DataSetRefresh2->Execute();
DataSetRefresh3->Execute();
DataSetRefresh4->Execute();

DataSetPost1->Execute();
DataSetPost2->Execute();
DataSetPost3->Execute();
DataSetPost4->Execute();
//if(Application->MessageBox("��������� ���������?","����� �� ���������",MB_YESNO+MB_ICONQUESTION)==IDYES	)
/*
if(Application->MessageBoxA("��������� ���������?","����� �� ���������",MB_YESNOCANCEL	+MB_ICONQUESTION+MB_DEFBUTTON1)==IDYES)
{

}
*/

switch (Application->MessageBoxA("�������� ��� ��������� �� ������?","����� �� ���������",MB_YESNOCANCEL	+MB_ICONQUESTION+MB_DEFBUTTON1))
{
 case IDYES:
 {
 Zast->MClient->WriteDiaryEvent("Hazards","���������� ������ ������ ������","");
 SaveAllTables(true);
  Zast->Close();
 break;
 }
 case IDNO:
 {
 Zast->MClient->WriteDiaryEvent("Hazards","���������� ������ ����� �� ������","");
 Action=caFree;
 Zast->Close();
 break;
 }
 case IDCANCEL:
 {
 Action=caNone;
 break;
 }
}


}
//---------------------------------------------------------------------------
void TDocuments::LoadTab1()
{
Nodes->Active=false;
Nodes->CommandText="Select * From ����_3";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
 int Parent=Nodes->FieldByName("��������")->Value;
 if (Parent!=0)
 {
 Branches->Active=false;
 Branches->CommandText="Select * From ����_3 Where [����� ����]="+IntToStr(Parent);
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
//TTreeNode *N2;
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
N1=TreeView1->Items->AddChildObject(NULL,Text,MyNodePtr);

Nodes->Post();

Branches->Active=false;
Branches->CommandText="Select * From �����_3 Where [����� ��������]="+IntToStr(NumParent);
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
if (View==true)
{
TreeView1->Items->AddChildObject(N1,TextBranches,MyNodePtr);
}

Branches->Edit();

Branches->Post();
Branches->Next();
}



Nodes->Active=false;
Nodes->CommandText="Select * From ����_3 Where ([��������])>0";
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
N=TreeView1->Items->AddChildObject(ParNod,TextNode,MyNodePtr);


Nodes->Edit();

Nodes->Post();
Node NN;
NN.Number=NumNode;
NN.Parent=NumParent;
NN.Nod=N;
It_Node_3=NodeVector_3.end();
NodeVector_3.insert(It_Node_3,NN);


Branches->Active=false;
Branches->CommandText="Select * From �����_3 Where [����� ��������]="+IntToStr(NumNode);
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

//ShowMessage(TreeView1->Items->Item[0]->Text);
TreeView1->Items->Item[0]->Expand(false);
/*
for(int i=0;i<TreeView1->Items->Count;i++)
{
ShowMessage(TreeView1->Items->Item[i]->Text);
TreeView1->Items->Item[i]->Expand(false);
}
*/
}
//-------------------------------------------------------------------------------------------------
void TDocuments::LoadTab2()
{
Nodes->Active=false;
Nodes->CommandText="Select * From ����_4";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
 int Parent=Nodes->FieldByName("��������")->Value;
 if (Parent!=0)
 {
 Branches->Active=false;
 Branches->CommandText="Select * From ����_4 Where [����� ����]="+IntToStr(Parent);
 Branches->Active=true;
 if (Branches->RecordCount==0)
 {
  Nodes->Delete();
 }
// Nodes->Next();
 }
 Nodes->Next();
}
//--------------------------------------------------



PMyNode  MyNodePtr;
MyNodePtr = new TMyNode;


TTreeNode *N;
TTreeNode *N1;

int NumParent;
TreeView2->Items->BeginUpdate();
TreeView2->Items->Clear();


Nodes->Active=true;
Nodes->First();
NumParent=Nodes->FieldByName("����� ����")->Value;
AnsiString Text=Nodes->FieldByName("��������")->Value;

MyNodePtr->Number=NumParent;
MyNodePtr->Parent=0;
MyNodePtr->Node=true;

Nodes->Edit();
N1=TreeView2->Items->AddChildObject(NULL,Text,MyNodePtr);

Nodes->Post();

Branches->Active=false;
Branches->CommandText="Select * From �����_4 Where [����� ��������]="+IntToStr(NumParent);
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
if (View==true)
{
TreeView2->Items->AddChildObject(N1,TextBranches,MyNodePtr);
}

Branches->Edit();

Branches->Post();
Branches->Next();
}



Nodes->Active=false;
Nodes->CommandText="Select * From ����_4 Where ([��������])>0";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
int NumNode=Nodes->FieldByName("����� ����")->Value;
NumParent=Nodes->FieldByName("��������")->Value;
AnsiString TextNode=Nodes->FieldByName("��������")->Value;
TTreeNode *ParNod;
ParNod=N1;
for (unsigned int k=0; k<NodeVector_4.size();k++)
{
 if (NodeVector_4[k].Number==NumParent)
 {
  ParNod=NodeVector_4[k].Nod;
 }
}

MyNodePtr = new TMyNode;
MyNodePtr->Number=NumNode;
MyNodePtr->Parent=NumParent;
MyNodePtr->Node=true;
N=TreeView2->Items->AddChildObject(ParNod,TextNode,MyNodePtr);


Nodes->Edit();

Nodes->Post();
Node NN;
NN.Number=NumNode;
NN.Parent=NumParent;
NN.Nod=N;
It_Node_4=NodeVector_4.end();
NodeVector_4.insert(It_Node_4,NN);


Branches->Active=false;
Branches->CommandText="Select * From �����_4 Where [����� ��������]="+IntToStr(NumNode);
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
TreeView2->Items->AddChildObject(N,TextBranches,MyNodePtr);


Branches->Edit();

Branches->Post();
Branches->Next();
}
Nodes->Next();
}
Nodes->Active=false;
Branches->Active=false;


TreeView2->Items->EndUpdate();
TreeView2->Items->Item[0]->Expand(false);
/*
for(int i=0;i<TreeView2->Items->Count;i++)
{
TreeView2->Items->Item[i]->Expand(true);
}
*/
}
//---------------------------------------------------------------------------------
void TDocuments::LoadTab3()
{
Nodes->Active=false;
Nodes->CommandText="Select * From ����_5";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
 int Parent=Nodes->FieldByName("��������")->Value;
 if (Parent!=0)
 {
 Branches->Active=false;
 Branches->CommandText="Select * From ����_5 Where [����� ����]="+IntToStr(Parent);
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

int NumParent;
TreeView3->Items->BeginUpdate();
TreeView3->Items->Clear();


Nodes->Active=true;
Nodes->First();
NumParent=Nodes->FieldByName("����� ����")->Value;
AnsiString Text=Nodes->FieldByName("��������")->Value;

MyNodePtr->Number=NumParent;
MyNodePtr->Parent=0;
MyNodePtr->Node=true;

Nodes->Edit();
N1=TreeView3->Items->AddChildObject(NULL,Text,MyNodePtr);

Nodes->Post();

Branches->Active=false;
Branches->CommandText="Select * From �����_5 Where [����� ��������]="+IntToStr(NumParent);
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
if (View==true)
{
TreeView3->Items->AddChildObject(N1,TextBranches,MyNodePtr);
}

Branches->Edit();

Branches->Post();
Branches->Next();
}



Nodes->Active=false;
Nodes->CommandText="Select * From ����_5 Where ([��������])>0";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
int NumNode=Nodes->FieldByName("����� ����")->Value;
NumParent=Nodes->FieldByName("��������")->Value;
AnsiString TextNode=Nodes->FieldByName("��������")->Value;
TTreeNode *ParNod;
ParNod=N1;
for (unsigned int k=0; k<NodeVector_5.size();k++)
{
 if (NodeVector_5[k].Number==NumParent)
 {
  ParNod=NodeVector_5[k].Nod;
 }
}

MyNodePtr = new TMyNode;
MyNodePtr->Number=NumNode;
MyNodePtr->Parent=NumParent;
MyNodePtr->Node=true;
//N=TreeView1->Items->AddChildObject(ParNod,TextNode+" "+IntToStr(MyNodePtr->Number)+" "+IntToStr(MyNodePtr->Parent)+" ����",MyNodePtr);
N=TreeView3->Items->AddChildObject(ParNod,TextNode,MyNodePtr);

//N=TreeView1->Items->AddChild(ParNod,TextNode);
Nodes->Edit();

Nodes->Post();
Node NN;
NN.Number=NumNode;
NN.Parent=NumParent;
NN.Nod=N;
It_Node_5=NodeVector_5.end();
NodeVector_5.insert(It_Node_5,NN);


Branches->Active=false;
Branches->CommandText="Select * From �����_5 Where [����� ��������]="+IntToStr(NumNode);
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
TreeView3->Items->AddChildObject(N,TextBranches,MyNodePtr);


Branches->Edit();

Branches->Post();
Branches->Next();
}
Nodes->Next();
}
Nodes->Active=false;
Branches->Active=false;


TreeView3->Items->EndUpdate();
TreeView3->Items->Item[0]->Expand(false);
/*
for(int i=0;i<TreeView3->Items->Count;i++)
{
TreeView3->Items->Item[i]->Expand(true);
}
*/
}
//------------------------------------------------------------------------------
void TDocuments::LoadTab4()
{
Nodes->Active=false;
Nodes->CommandText="Select * From ����_6";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
 int Parent=Nodes->FieldByName("��������")->Value;
 if (Parent!=0)
 {
 Branches->Active=false;
 Branches->CommandText="Select * From ����_6 Where [����� ����]="+IntToStr(Parent);
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

int NumParent;
TreeView4->Items->BeginUpdate();
TreeView4->Items->Clear();


Nodes->Active=true;
Nodes->First();
NumParent=Nodes->FieldByName("����� ����")->Value;
AnsiString Text=Nodes->FieldByName("��������")->Value;

MyNodePtr->Number=NumParent;
MyNodePtr->Parent=0;
MyNodePtr->Node=true;

Nodes->Edit();
N1=TreeView4->Items->AddChildObject(NULL,Text,MyNodePtr);

Nodes->Post();

Branches->Active=false;
Branches->CommandText="Select * From �����_6 Where [����� ��������]="+IntToStr(NumParent);
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
if (View==true)
{
TreeView4->Items->AddChildObject(N1,TextBranches,MyNodePtr);
}

Branches->Edit();

Branches->Post();
Branches->Next();
}



Nodes->Active=false;
Nodes->CommandText="Select * From ����_6 Where ([��������])>0";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
int NumNode=Nodes->FieldByName("����� ����")->Value;
NumParent=Nodes->FieldByName("��������")->Value;
AnsiString TextNode=Nodes->FieldByName("��������")->Value;
TTreeNode *ParNod;
ParNod=N1;
for (unsigned int k=0; k<NodeVector_6.size();k++)
{
 if (NodeVector_6[k].Number==NumParent)
 {
  ParNod=NodeVector_6[k].Nod;
 }
}

MyNodePtr = new TMyNode;
MyNodePtr->Number=NumNode;
MyNodePtr->Parent=NumParent;
MyNodePtr->Node=true;
//N=TreeView1->Items->AddChildObject(ParNod,TextNode+" "+IntToStr(MyNodePtr->Number)+" "+IntToStr(MyNodePtr->Parent)+" ����",MyNodePtr);
N=TreeView4->Items->AddChildObject(ParNod,TextNode,MyNodePtr);

//N=TreeView1->Items->AddChild(ParNod,TextNode);
Nodes->Edit();

Nodes->Post();
Node NN;
NN.Number=NumNode;
NN.Parent=NumParent;
NN.Nod=N;
It_Node_6=NodeVector_6.end();
NodeVector_6.insert(It_Node_6,NN);


Branches->Active=false;
Branches->CommandText="Select * From �����_6 Where [����� ��������]="+IntToStr(NumNode);
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
TreeView4->Items->AddChildObject(N,TextBranches,MyNodePtr);


Branches->Edit();

Branches->Post();
Branches->Next();
}
Nodes->Next();
}
Nodes->Active=false;
Branches->Active=false;


TreeView4->Items->EndUpdate();
TreeView4->Items->Item[0]->Expand(false);
/*
for(int i=0;i<TreeView4->Items->Count;i++)
{
TreeView4->Items->Item[i]->Expand(true);
}
*/
}
//---------------------------------------------------------------------------------
void TDocuments::LoadTab5()
{
Nodes->Active=false;
Nodes->CommandText="Select * From ����_7";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
 int Parent=Nodes->FieldByName("��������")->Value;
 if (Parent!=0)
 {
 Branches->Active=false;
 Branches->CommandText="Select * From ����_7 Where [����� ����]="+IntToStr(Parent);
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
int NumParent;
TreeView5->Items->BeginUpdate();
TreeView5->Items->Clear();


Nodes->Active=true;
Nodes->First();
NumParent=Nodes->FieldByName("����� ����")->Value;
AnsiString Text=Nodes->FieldByName("��������")->Value;

MyNodePtr->Number=NumParent;
MyNodePtr->Parent=0;
MyNodePtr->Node=true;

Nodes->Edit();
N1=TreeView5->Items->AddChildObject(NULL,Text,MyNodePtr);

Nodes->Post();

Branches->Active=false;
Branches->CommandText="Select * From �����_7 Where [����� ��������]="+IntToStr(NumParent);
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
if (View==true)
{
TreeView5->Items->AddChildObject(N1,TextBranches,MyNodePtr);
}

Branches->Edit();

Branches->Post();
Branches->Next();
}



Nodes->Active=false;
Nodes->CommandText="Select * From ����_7 Where ([��������])>0";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
int NumNode=Nodes->FieldByName("����� ����")->Value;
NumParent=Nodes->FieldByName("��������")->Value;
AnsiString TextNode=Nodes->FieldByName("��������")->Value;
TTreeNode *ParNod;
ParNod=N1;
for (unsigned int k=0; k<NodeVector_7.size();k++)
{
 if (NodeVector_7[k].Number==NumParent)
 {
  ParNod=NodeVector_7[k].Nod;
 }
}

MyNodePtr = new TMyNode;
MyNodePtr->Number=NumNode;
MyNodePtr->Parent=NumParent;
MyNodePtr->Node=true;
N=TreeView5->Items->AddChildObject(ParNod,TextNode,MyNodePtr);


Nodes->Edit();

Nodes->Post();
Node NN;
NN.Number=NumNode;
NN.Parent=NumParent;
NN.Nod=N;
It_Node_7=NodeVector_7.end();
NodeVector_7.insert(It_Node_7,NN);


Branches->Active=false;
Branches->CommandText="Select * From �����_7 Where [����� ��������]="+IntToStr(NumNode);
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
TreeView5->Items->AddChildObject(N,TextBranches,MyNodePtr);


Branches->Edit();

Branches->Post();
Branches->Next();
}
Nodes->Next();
}
Nodes->Active=false;
Branches->Active=false;


TreeView5->Items->EndUpdate();
TreeView5->Items->Item[0]->Expand(false);
/*
for(int i=0;i<TreeView5->Items->Count;i++)
{
TreeView5->Items->Item[i]->Expand(true);
}
*/
}
//-----------------------------------------------------------------------------------------------------
void __fastcall TDocuments::TreeView1Edited(TObject *Sender,
      TTreeNode *Node, AnsiString &S)
{
bool IsNode=PMyNode(Node->Data)->Node;
int Number=PMyNode(Node->Data)->Number;
if (S=="")
{
S=" ";
}
if (IsNode==true)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From ����_3 Where [����� ����]="+IntToStr(Number);
 Nodes->Active=true;
 Nodes->Edit();
 Nodes->FieldByName("��������")->Value=S;
 Nodes->Post();

}
else
{
 Branches->Active=false;
 Branches->CommandText="Select * From �����_3 Where [����� �����]="+IntToStr(Number);
 Branches->Active=true;
 Branches->Edit();
 Branches->FieldByName("��������")->Value=S;
 Branches->Post();
}
}
//---------------------------------------------------------------------------


void __fastcall TDocuments::N3Click(TObject *Sender)
{
int Number;
bool IsNode;
TTreeNode *N;
//PMyNode(TreeView1->Selected->Data)->Number;

//����� ������
//����� ������ ����� ����� �����
int NumParentNode=PMyNode(SelNode->Data)->Number;
//�������� �� ������ �����?
IsNode=PMyNode(SelNode->Data)->Node;
if (IsNode==true)
{
//���� �������� �� ����� ������ �������� �� ���� ������
Branches->Active=false;
Branches->CommandText="Select * From �����_3";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("����� ��������")->Value=NumParentNode;
Branches->FieldByName("��������")->Value="����� �����������";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("����� �����")->Value;

MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;

}
else
{
//� ���� �������� ������ �� ����� �� ��������� � ���� � �������� �� ��� ������.
int Par;
AnsiString Text;
Branches->Active=false;
Branches->CommandText="Select * From �����_3 Where [����� �����]="+IntToStr(NumParentNode);
Branches->Active=true;
Par=Branches->FieldByName("����� ��������")->Value;
Text=Branches->FieldByName("��������")->Value;
Branches->Delete();
Nodes->Active=false;
Nodes->CommandText="Select * From ����_3";
Nodes->Active=true;
Nodes->Append();
Nodes->FieldByName("��������")->Value=Par;
Nodes->FieldByName("��������")->Value=Text;
Nodes->Post();
Nodes->Last();

//PMyNode(TreeView1->Selected->Data)->Node
PMyNode(SelNode->Data)->Number=Nodes->FieldByName("����� ����")->Value;
PMyNode(SelNode->Data)->Parent=Nodes->FieldByName("��������")->Value;
PMyNode(SelNode->Data)->Node=true;
NumParentNode=Nodes->FieldByName("����� ����")->Value;
Branches->Active=false;
Branches->CommandText="Select * From �����_3";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("����� ��������")->Value=NumParentNode;
Branches->FieldByName("��������")->Value="����� �����������";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("����� �����")->Value;


MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;
}
TreeView1->Items->AddChildObject(SelNode,"����� �����������",MyNodePtr);
}
//---------------------------------------------------------------------------
int TDocuments::DeleteNode(String NameNode, String NameBranch, int Number)
{
MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOConn;
Tab->CommandText="Select * From "+NameNode+" Where [��������]="+IntToStr(Number);
Tab->Active=true;

if(Tab->RecordCount!=0)
{
for(Tab->First();!Tab->Eof;Tab->Next())
{
 int Parent=Tab->FieldByName("����� ����")->AsInteger;
 DeleteNode(NameNode, NameBranch, Parent);
}

}
Comm->CommandText="Delete * From "+NameBranch+" Where [����� ��������]="+IntToStr(Number);
Comm->Execute();

 Comm->CommandText="Delete * From "+NameNode+" Where [����� ����]="+IntToStr(Number);
 Comm->Execute();



return -1;
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::N9Click(TObject *Sender)
{
//�������� �������� �� ���������� � �������� ��������� ����� (����) ��������
//� �������� ��� � ��������� ������
TTreeNode *N;


int Number=PMyNode(TreeView1->Selected->Data)->Number;
int Parent=PMyNode(TreeView1->Selected->Data)->Parent;
bool IsNode=PMyNode(TreeView1->Selected->Data)->Node;
if (IsNode==true)
{


DeleteNode("����_3","�����_5", Number);

}
else
{
Comm->CommandText="Delete * From �����_3 Where [����� �����]="+IntToStr(Number);
Comm->Execute();

}

Branches->Active=false;
Branches->CommandText="Select * From �����_3 Where [����� ��������]="+IntToStr(Parent);
Branches->Active=true;
Nodes->Active=false;
Nodes->CommandText="Select * From ����_3 Where ��������="+IntToStr(Parent);
Nodes->Active=true;

if (Branches->RecordCount==0 & Nodes->RecordCount==0)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From ����_3 Where [����� ����]="+IntToStr(Parent);
 Nodes->Active=true;
 if (Nodes->RecordCount==0)
 {
  ShowMessage("������! ��������� ������ � ������� ����� ����� "+IntToStr(Number)+" ������ ����� "+IntToStr(Parent));
  ShowMessage("�� ������� ������ � ������� ����� ����� "+IntToStr(Parent));

 }
 else
 {
 int NumParent=Nodes->FieldByName("��������")->Value;
 if (NumParent!=0)
 {
 AnsiString Text=Nodes->FieldByName("��������")->Value;;
 Nodes->Delete();
 Branches->Active=false;
 Branches->CommandText="Select * From �����_3";
 Branches->Active=true;
 Branches->Append();
 Branches->FieldByName("����� ��������")->Value=NumParent;
 Branches->FieldByName("��������")->Value=Text;
 Branches->Post();
 Branches->Last();
 int Number=Branches->FieldByName("����� �����")->Value;
 N=TreeView1->GetNodeAt(XX,YY);
 TTreeNode * ParNode=N->Parent;
 PMyNode(ParNode->Data)->Number=Number;
 PMyNode(ParNode->Data)->Parent=NumParent;
 PMyNode(ParNode->Data)->Node=false;
 //
 //ParNode->Text=Text+" "+IntToStr(Number)+" "+IntToStr(NumParent)+" �����";
}
}
}
TreeView1->Items->Delete(TreeView1->GetNodeAt(XX,YY));

}
//---------------------------------------------------------------------------

void __fastcall TDocuments::TreeView1MouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
TTreeNode* N=TreeView1->GetNodeAt(X,Y);
//if(N!=NULL)
//{
//N->Selected=true;
THitTests HT;
HT=TreeView1->GetHitTestInfoAt(X,Y);
if (HT.Contains(htOnItem))
{
N->Selected=true;
SelNode=TreeView1->Selected;
//ShowMessage(SelNode->Text);
 if (Button==mbRight)
 {
  TPoint TP;
  TP=TreeView1->ClientOrigin;
  if (SelNode->AbsoluteIndex==0)
  {
  PopupMenu1->Items->Items[0]->Enabled=true;
  PopupMenu1->Items->Items[1]->Enabled=false;
  }
  else
  {
  PopupMenu1->Items->Items[0]->Enabled=true;
  PopupMenu1->Items->Items[1]->Enabled=true;
  }
  PopupMenu1->Popup(X+TP.x,Y+TP.y);

 }

XX=X;
YY=Y;
// TreeView1->Items->Delete(TreeView1->GetNodeAt(X,Y));
}
//}
}
//---------------------------------------------------------------------------



void __fastcall TDocuments::TreeView2Edited(TObject *Sender,
      TTreeNode *Node, AnsiString &S)
{
bool IsNode=PMyNode(Node->Data)->Node;
int Number=PMyNode(Node->Data)->Number;
if (S=="")
{
S=" ";
}
if (IsNode==true)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From ����_4 Where [����� ����]="+IntToStr(Number);
 Nodes->Active=true;
 Nodes->Edit();
 Nodes->FieldByName("��������")->Value=S;
 Nodes->Post();

}
else
{
 Branches->Active=false;
 Branches->CommandText="Select * From �����_4 Where [����� �����]="+IntToStr(Number);
 Branches->Active=true;
 Branches->Edit();
 Branches->FieldByName("��������")->Value=S;
 Branches->Post();
}
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::TreeView2MouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
TTreeNode* N=TreeView2->GetNodeAt(X,Y);

THitTests HT;
HT=TreeView2->GetHitTestInfoAt(X,Y);
if (HT.Contains(htOnItem))
{
N->Selected=true;
SelNode=TreeView2->Selected;

//ShowMessage(SelNode->Text);
 if (Button==mbRight)
 {
  TPoint TP;
  TP=TreeView2->ClientOrigin;
  if (SelNode->AbsoluteIndex==0)
  {
  PopupMenu2->Items->Items[0]->Enabled=true;
  PopupMenu2->Items->Items[1]->Enabled=false;
  }
  else
  {
  PopupMenu2->Items->Items[0]->Enabled=true;
  PopupMenu2->Items->Items[1]->Enabled=true;
  }
  PopupMenu2->Popup(X+TP.x,Y+TP.y);

 }

XX=X;
YY=Y;
}
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N10Click(TObject *Sender)
{
int Number;
bool IsNode;
TTreeNode *N;
//PMyNode(TreeView1->Selected->Data)->Number;

//����� ������
//����� ������ ����� ����� �����
int NumParentNode=PMyNode(SelNode->Data)->Number;
//�������� �� ������ �����?
IsNode=PMyNode(SelNode->Data)->Node;
if (IsNode==true)
{
//���� �������� �� ����� ������ �������� �� ���� ������
Branches->Active=false;
Branches->CommandText="Select * From �����_4";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("����� ��������")->Value=NumParentNode;
Branches->FieldByName("��������")->Value="����� �����������";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("����� �����")->Value;

MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;

}
else
{
//� ���� �������� ������ �� ����� �� ��������� � ���� � �������� �� ��� ������.
int Par;
AnsiString Text;
Branches->Active=false;
Branches->CommandText="Select * From �����_4 Where [����� �����]="+IntToStr(NumParentNode);
Branches->Active=true;
Par=Branches->FieldByName("����� ��������")->Value;
Text=Branches->FieldByName("��������")->Value;
Branches->Delete();
Nodes->Active=false;
Nodes->CommandText="Select * From ����_4";
Nodes->Active=true;
Nodes->Append();
Nodes->FieldByName("��������")->Value=Par;
Nodes->FieldByName("��������")->Value=Text;
Nodes->Post();
Nodes->Last();

//PMyNode(TreeView1->Selected->Data)->Node
PMyNode(SelNode->Data)->Number=Nodes->FieldByName("����� ����")->Value;
PMyNode(SelNode->Data)->Parent=Nodes->FieldByName("��������")->Value;
PMyNode(SelNode->Data)->Node=true;
NumParentNode=Nodes->FieldByName("����� ����")->Value;
Branches->Active=false;
Branches->CommandText="Select * From �����_4";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("����� ��������")->Value=NumParentNode;
Branches->FieldByName("��������")->Value="����� �����������";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("����� �����")->Value;


MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;
}
TreeView2->Items->AddChildObject(SelNode,"����� �����������",MyNodePtr);
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N11Click(TObject *Sender)
{
//�������� �������� �� ���������� � �������� ��������� ����� (����) ��������
//� �������� ��� � ��������� ������
TTreeNode *N;


int Number=PMyNode(TreeView2->Selected->Data)->Number;
int Parent=PMyNode(TreeView2->Selected->Data)->Parent;
bool IsNode=PMyNode(TreeView2->Selected->Data)->Node;
if (IsNode==true)
{
DeleteNode("����_4","�����_4", Number);

}
else
{
Comm->CommandText="Delete * From �����_4 Where [����� �����]="+IntToStr(Number);
Comm->Execute();

}

Branches->Active=false;
Branches->CommandText="Select * From �����_4 Where [����� ��������]="+IntToStr(Parent);
Branches->Active=true;
Nodes->Active=false;
Nodes->CommandText="Select * From ����_4 Where ��������="+IntToStr(Parent);
Nodes->Active=true;
if (Branches->RecordCount==0 & Nodes->RecordCount==0)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From ����_4 Where [����� ����]="+IntToStr(Parent);
 Nodes->Active=true;
 if (Nodes->RecordCount==0)
 {
  ShowMessage("������! ��������� ������ � ������� ����� ����� "+IntToStr(Number)+" ������ ����� "+IntToStr(Parent));
  ShowMessage("�� ������� ������ � ������� ����� ����� "+IntToStr(Parent));

 }
 else
 {
 int NumParent=Nodes->FieldByName("��������")->Value;
 AnsiString Text=Nodes->FieldByName("��������")->Value;
  if (NumParent!=0)
 {
 Nodes->Delete();
 Branches->Active=false;
 Branches->CommandText="Select * From �����_4";
 Branches->Active=true;
 Branches->Append();
 Branches->FieldByName("����� ��������")->Value=NumParent;
 Branches->FieldByName("��������")->Value=Text;
 Branches->Post();
 Branches->Last();
 int Number=Branches->FieldByName("����� �����")->Value;
 N=TreeView2->GetNodeAt(XX,YY);
 TTreeNode * ParNode=N->Parent;
 PMyNode(ParNode->Data)->Number=Number;
 PMyNode(ParNode->Data)->Parent=NumParent;
 PMyNode(ParNode->Data)->Node=false;
 //
 //ParNode->Text=Text+" "+IntToStr(Number)+" "+IntToStr(NumParent)+" �����";
}
}
}
TreeView2->Items->Delete(TreeView2->GetNodeAt(XX,YY));

}
//---------------------------------------------------------------------------

void __fastcall TDocuments::TreeView3Edited(TObject *Sender,
      TTreeNode *Node, AnsiString &S)
{
bool IsNode=PMyNode(Node->Data)->Node;
int Number=PMyNode(Node->Data)->Number;
if (S=="")
{
S=" ";
}
if (IsNode==true)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From ����_5 Where [����� ����]="+IntToStr(Number);
 Nodes->Active=true;
 Nodes->Edit();
 Nodes->FieldByName("��������")->Value=S;
 Nodes->Post();

}
else
{
 Branches->Active=false;
 Branches->CommandText="Select * From �����_5 Where [����� �����]="+IntToStr(Number);
 Branches->Active=true;
 Branches->Edit();
 Branches->FieldByName("��������")->Value=S;
 Branches->Post();
}
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::PageControl1MouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
THitTests HT;
HT=TreeView3->GetHitTestInfoAt(X,Y);
if (HT.Contains(htOnItem))
{
SelNode=TreeView3->Selected;
//ShowMessage(SelNode->Text);
 if (Button==mbRight)
 {
  TPoint TP;
  TP=TreeView3->ClientOrigin;
  if (SelNode->AbsoluteIndex==0)
  {
  PopupMenu3->Items->Items[0]->Enabled=true;
  PopupMenu3->Items->Items[1]->Enabled=false;
  }
  else
  {
  PopupMenu3->Items->Items[0]->Enabled=true;
  PopupMenu3->Items->Items[1]->Enabled=true;
  }
  PopupMenu3->Popup(X+TP.x,Y+TP.y);

 }

XX=X;
YY=Y;
// TreeView1->Items->Delete(TreeView1->GetNodeAt(X,Y));
}
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N12Click(TObject *Sender)
{
int Number;
bool IsNode;
TTreeNode *N;
//PMyNode(TreeView1->Selected->Data)->Number;

//����� ������
//����� ������ ����� ����� �����
int NumParentNode=PMyNode(SelNode->Data)->Number;
//�������� �� ������ �����?
IsNode=PMyNode(SelNode->Data)->Node;
if (IsNode==true)
{
//���� �������� �� ����� ������ �������� �� ���� ������
Branches->Active=false;
Branches->CommandText="Select * From �����_5";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("����� ��������")->Value=NumParentNode;
Branches->FieldByName("��������")->Value="����� ����������";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("����� �����")->Value;

MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;

}
else
{
//� ���� �������� ������ �� ����� �� ��������� � ���� � �������� �� ��� ������.
int Par;
AnsiString Text;
Branches->Active=false;
Branches->CommandText="Select * From �����_5 Where [����� �����]="+IntToStr(NumParentNode);
Branches->Active=true;
Par=Branches->FieldByName("����� ��������")->Value;
Text=Branches->FieldByName("��������")->Value;
Branches->Delete();
Nodes->Active=false;
Nodes->CommandText="Select * From ����_5";
Nodes->Active=true;
Nodes->Append();
Nodes->FieldByName("��������")->Value=Par;
Nodes->FieldByName("��������")->Value=Text;
Nodes->Post();
Nodes->Last();

//PMyNode(TreeView1->Selected->Data)->Node
PMyNode(SelNode->Data)->Number=Nodes->FieldByName("����� ����")->Value;
PMyNode(SelNode->Data)->Parent=Nodes->FieldByName("��������")->Value;
PMyNode(SelNode->Data)->Node=true;
NumParentNode=Nodes->FieldByName("����� ����")->Value;
Branches->Active=false;
Branches->CommandText="Select * From �����_5";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("����� ��������")->Value=NumParentNode;
Branches->FieldByName("��������")->Value="����� ����������";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("����� �����")->Value;


MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;
}
TreeView3->Items->AddChildObject(SelNode,"����� ����������",MyNodePtr);

}
//---------------------------------------------------------------------------


void __fastcall TDocuments::TreeView3MouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
TTreeNode* N=TreeView3->GetNodeAt(X,Y);

THitTests HT;
HT=TreeView3->GetHitTestInfoAt(X,Y);
if (HT.Contains(htOnItem))
{
N->Selected=true;
SelNode=TreeView3->Selected;

 if (Button==mbRight)
 {
  TPoint TP;
  TP=TreeView3->ClientOrigin;
  if (SelNode->AbsoluteIndex==0)
  {
  PopupMenu3->Items->Items[0]->Enabled=true;
  PopupMenu3->Items->Items[1]->Enabled=false;
  }
  else
  {
  PopupMenu3->Items->Items[0]->Enabled=true;
  PopupMenu3->Items->Items[1]->Enabled=true;
  }
  PopupMenu3->Popup(X+TP.x,Y+TP.y);

 }

XX=X;
YY=Y;
}
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N13Click(TObject *Sender)
{
//�������� �������� �� ���������� � �������� ��������� ����� (����) ��������
//� �������� ��� � ��������� ������
TTreeNode *N;
//N=TreeView1->GetNodeAt(XX,YY);

int Number=PMyNode(TreeView3->Selected->Data)->Number;
int Parent=PMyNode(TreeView3->Selected->Data)->Parent;
bool IsNode=PMyNode(TreeView3->Selected->Data)->Node;
if (IsNode==true)
{
DeleteNode("����_5","�����_5", Number);

}
else
{
Comm->CommandText="Delete * From �����_5 Where [����� �����]="+IntToStr(Number);
Comm->Execute();

}

Branches->Active=false;
Branches->CommandText="Select * From �����_5 Where [����� ��������]="+IntToStr(Parent);
Branches->Active=true;
Nodes->Active=false;
Nodes->CommandText="Select * From ����_5 Where ��������="+IntToStr(Parent);
Nodes->Active=true;
if (Branches->RecordCount==0 & Nodes->RecordCount==0)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From ����_5 Where [����� ����]="+IntToStr(Parent);
 Nodes->Active=true;
 if (Nodes->RecordCount==0)
 {
  ShowMessage("������! ��������� ������ � ������� ����� ����� "+IntToStr(Number)+" ������ ����� "+IntToStr(Parent));
  ShowMessage("�� ������� ������ � ������� ����� ����� "+IntToStr(Parent));

 }
 else
 {
 int NumParent=Nodes->FieldByName("��������")->Value;
 AnsiString Text=Nodes->FieldByName("��������")->Value;
  if (NumParent!=0)
 {
 Nodes->Delete();
 Branches->Active=false;
 Branches->CommandText="Select * From �����_5";
 Branches->Active=true;
 Branches->Append();
 Branches->FieldByName("����� ��������")->Value=NumParent;
 Branches->FieldByName("��������")->Value=Text;
 Branches->Post();
 Branches->Last();
 int Number=Branches->FieldByName("����� �����")->Value;
 N=TreeView3->GetNodeAt(XX,YY);
 TTreeNode * ParNode=N->Parent;
 PMyNode(ParNode->Data)->Number=Number;
 PMyNode(ParNode->Data)->Parent=NumParent;
 PMyNode(ParNode->Data)->Node=false;
 //
 //ParNode->Text=Text+" "+IntToStr(Number)+" "+IntToStr(NumParent)+" �����";
}
}
}
TreeView3->Items->Delete(TreeView3->GetNodeAt(XX,YY));

}
//---------------------------------------------------------------------------

void __fastcall TDocuments::TreeView4Edited(TObject *Sender,
      TTreeNode *Node, AnsiString &S)
{
bool IsNode=PMyNode(Node->Data)->Node;
int Number=PMyNode(Node->Data)->Number;
if (S=="")
{
S=" ";
}
if (IsNode==true)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From ����_6 Where [����� ����]="+IntToStr(Number);
 Nodes->Active=true;
 Nodes->Edit();
 Nodes->FieldByName("��������")->Value=S;
 Nodes->Post();

}
else
{
 Branches->Active=false;
 Branches->CommandText="Select * From �����_6 Where [����� �����]="+IntToStr(Number);
 Branches->Active=true;
 Branches->Edit();
 Branches->FieldByName("��������")->Value=S;
 Branches->Post();
}
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::TreeView4MouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
TTreeNode* N=TreeView4->GetNodeAt(X,Y);

THitTests HT;
HT=TreeView4->GetHitTestInfoAt(X,Y);
if (HT.Contains(htOnItem))
{
N->Selected=true;
SelNode=TreeView4->Selected;
//ShowMessage(SelNode->Text);
 if (Button==mbRight)
 {
  TPoint TP;
  TP=TreeView4->ClientOrigin;
  if (SelNode->AbsoluteIndex==0)
  {
  PopupMenu4->Items->Items[0]->Enabled=true;
  PopupMenu4->Items->Items[1]->Enabled=false;
  }
  else
  {
  PopupMenu4->Items->Items[0]->Enabled=true;
  PopupMenu4->Items->Items[1]->Enabled=true;
  }
  PopupMenu4->Popup(X+TP.x,Y+TP.y);

 }

XX=X;
YY=Y;
}
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N14Click(TObject *Sender)
{
int Number;
bool IsNode;
TTreeNode *N;


//����� ������
//����� ������ ����� ����� �����
int NumParentNode=PMyNode(SelNode->Data)->Number;
//�������� �� ������ �����?
IsNode=PMyNode(SelNode->Data)->Node;
if (IsNode==true)
{
//���� �������� �� ����� ������ �������� �� ���� ������
Branches->Active=false;
Branches->CommandText="Select * From �����_6";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("����� ��������")->Value=NumParentNode;
Branches->FieldByName("��������")->Value="����� ��� ������������";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("����� �����")->Value;

MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;

}
else
{
//� ���� �������� ������ �� ����� �� ��������� � ���� � �������� �� ��� ������.
int Par;
AnsiString Text;
Branches->Active=false;
Branches->CommandText="Select * From �����_6 Where [����� �����]="+IntToStr(NumParentNode);
Branches->Active=true;
Par=Branches->FieldByName("����� ��������")->Value;
Text=Branches->FieldByName("��������")->Value;
Branches->Delete();
Nodes->Active=false;
Nodes->CommandText="Select * From ����_6";
Nodes->Active=true;
Nodes->Append();
Nodes->FieldByName("��������")->Value=Par;
Nodes->FieldByName("��������")->Value=Text;
Nodes->Post();
Nodes->Last();


PMyNode(SelNode->Data)->Number=Nodes->FieldByName("����� ����")->Value;
PMyNode(SelNode->Data)->Parent=Nodes->FieldByName("��������")->Value;
PMyNode(SelNode->Data)->Node=true;
NumParentNode=Nodes->FieldByName("����� ����")->Value;
Branches->Active=false;
Branches->CommandText="Select * From �����_6";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("����� ��������")->Value=NumParentNode;
Branches->FieldByName("��������")->Value="����� ��� ������������";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("����� �����")->Value;


MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;
}
TreeView4->Items->AddChildObject(SelNode,"����� ��� ������������",MyNodePtr);

}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N15Click(TObject *Sender)
{
//�������� �������� �� ���������� � �������� ��������� ����� (����) ��������
//� �������� ��� � ��������� ������
TTreeNode *N;
//N=TreeView1->GetNodeAt(XX,YY);

int Number=PMyNode(TreeView4->Selected->Data)->Number;
int Parent=PMyNode(TreeView4->Selected->Data)->Parent;
bool IsNode=PMyNode(TreeView4->Selected->Data)->Node;
if (IsNode==true)
{
DeleteNode("����_6","�����_6", Number);

}
else
{
Comm->CommandText="Delete * From �����_6 Where [����� �����]="+IntToStr(Number);
Comm->Execute();

}

Branches->Active=false;
Branches->CommandText="Select * From �����_6 Where [����� ��������]="+IntToStr(Parent);
Branches->Active=true;
Nodes->Active=false;
Nodes->CommandText="Select * From ����_6 Where ��������="+IntToStr(Parent);
Nodes->Active=true;
if (Branches->RecordCount==0 & Nodes->RecordCount==0)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From ����_6 Where [����� ����]="+IntToStr(Parent);
 Nodes->Active=true;
 if (Nodes->RecordCount==0)
 {
  ShowMessage("������! ��������� ������ � ������� ����� ����� "+IntToStr(Number)+" ������ ����� "+IntToStr(Parent));
  ShowMessage("�� ������� ������ � ������� ����� ����� "+IntToStr(Parent));

 }
 else
 {
 int NumParent=Nodes->FieldByName("��������")->Value;
 AnsiString Text=Nodes->FieldByName("��������")->Value;
  if (NumParent!=0)
 {
 Nodes->Delete();
 Branches->Active=false;
 Branches->CommandText="Select * From �����_6";
 Branches->Active=true;
 Branches->Append();
 Branches->FieldByName("����� ��������")->Value=NumParent;
 Branches->FieldByName("��������")->Value=Text;
 Branches->Post();
 Branches->Last();
 int Number=Branches->FieldByName("����� �����")->Value;
 N=TreeView4->GetNodeAt(XX,YY);
 TTreeNode * ParNode=N->Parent;
 PMyNode(ParNode->Data)->Number=Number;
 PMyNode(ParNode->Data)->Parent=NumParent;
 PMyNode(ParNode->Data)->Node=false;
 //
 //ParNode->Text=Text+" "+IntToStr(Number)+" "+IntToStr(NumParent)+" �����";
}
}
}
TreeView4->Items->Delete(TreeView4->GetNodeAt(XX,YY));

}
//---------------------------------------------------------------------------

void __fastcall TDocuments::TreeView5Edited(TObject *Sender,
      TTreeNode *Node, AnsiString &S)
{
bool IsNode=PMyNode(Node->Data)->Node;
int Number=PMyNode(Node->Data)->Number;
if (S=="")
{
S=" ";
}
if (IsNode==true)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From ����_7 Where [����� ����]="+IntToStr(Number);
 Nodes->Active=true;
 Nodes->Edit();
 Nodes->FieldByName("��������")->Value=S;
 Nodes->Post();

}
else
{
 Branches->Active=false;
 Branches->CommandText="Select * From �����_7 Where [����� �����]="+IntToStr(Number);
 Branches->Active=true;
 Branches->Edit();
 Branches->FieldByName("��������")->Value=S;
 Branches->Post();
}
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::TreeView5MouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
TTreeNode* N=TreeView5->GetNodeAt(X,Y);

THitTests HT;
HT=TreeView5->GetHitTestInfoAt(X,Y);
if (HT.Contains(htOnItem))
{
N->Selected=true;
SelNode=TreeView5->Selected;
//ShowMessage(SelNode->Text);
 if (Button==mbRight)
 {
  TPoint TP;
  TP=TreeView5->ClientOrigin;
  if (SelNode->AbsoluteIndex==0)
  {
  PopupMenu5->Items->Items[0]->Enabled=true;
  PopupMenu5->Items->Items[1]->Enabled=false;
  }
  else
  {
  PopupMenu5->Items->Items[0]->Enabled=true;
  PopupMenu5->Items->Items[1]->Enabled=true;
  }
  PopupMenu5->Popup(X+TP.x,Y+TP.y);

 }

XX=X;
YY=Y;
}
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N16Click(TObject *Sender)
{
int Number;
bool IsNode;
TTreeNode *N;


//����� ������
//����� ������ ����� ����� �����
int NumParentNode=PMyNode(SelNode->Data)->Number;
//�������� �� ������ �����?
IsNode=PMyNode(SelNode->Data)->Node;
if (IsNode==true)
{
//���� �������� �� ����� ������ �������� �� ���� ������
Branches->Active=false;
Branches->CommandText="Select * From �����_7";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("����� ��������")->Value=NumParentNode;
Branches->FieldByName("��������")->Value="����� ���������";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("����� �����")->Value;

MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;

}
else
{
//� ���� �������� ������ �� ����� �� ��������� � ���� � �������� �� ��� ������.
int Par;
AnsiString Text;
Branches->Active=false;
Branches->CommandText="Select * From �����_7 Where [����� �����]="+IntToStr(NumParentNode);
Branches->Active=true;
Par=Branches->FieldByName("����� ��������")->Value;
Text=Branches->FieldByName("��������")->Value;
Branches->Delete();
Nodes->Active=false;
Nodes->CommandText="Select * From ����_7";
Nodes->Active=true;
Nodes->Append();
Nodes->FieldByName("��������")->Value=Par;
Nodes->FieldByName("��������")->Value=Text;
Nodes->Post();
Nodes->Last();


PMyNode(SelNode->Data)->Number=Nodes->FieldByName("����� ����")->Value;
PMyNode(SelNode->Data)->Parent=Nodes->FieldByName("��������")->Value;
PMyNode(SelNode->Data)->Node=true;
NumParentNode=Nodes->FieldByName("����� ����")->Value;
Branches->Active=false;
Branches->CommandText="Select * From �����_7";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("����� ��������")->Value=NumParentNode;
Branches->FieldByName("��������")->Value="����� ���������";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("����� �����")->Value;


MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;
}
TreeView5->Items->AddChildObject(SelNode,"����� ���������",MyNodePtr);

}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N17Click(TObject *Sender)
{
//�������� �������� �� ���������� � �������� ��������� ����� (����) ��������
//� �������� ��� � ��������� ������
TTreeNode *N;


int Number=PMyNode(TreeView5->Selected->Data)->Number;
int Parent=PMyNode(TreeView5->Selected->Data)->Parent;
bool IsNode=PMyNode(TreeView5->Selected->Data)->Node;
if (IsNode==true)
{
DeleteNode("����_7","�����_7", Number);

}
else
{
Comm->CommandText="Delete * From �����_7 Where [����� �����]="+IntToStr(Number);
Comm->Execute();

}

Branches->Active=false;
Branches->CommandText="Select * From �����_7 Where [����� ��������]="+IntToStr(Parent);
Branches->Active=true;
Nodes->Active=false;
Nodes->CommandText="Select * From ����_7 Where ��������="+IntToStr(Parent);
Nodes->Active=true;
if (Branches->RecordCount==0 & Nodes->RecordCount==0)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From ����_7 Where [����� ����]="+IntToStr(Parent);
 Nodes->Active=true;
 if (Nodes->RecordCount==0)
 {
  ShowMessage("������! ��������� ������ � ������� ����� ����� "+IntToStr(Number)+" ������ ����� "+IntToStr(Parent));
  ShowMessage("�� ������� ������ � ������� ����� ����� "+IntToStr(Parent));

 }
 else
 {
 int NumParent=Nodes->FieldByName("��������")->Value;
 AnsiString Text=Nodes->FieldByName("��������")->Value;
  if (NumParent!=0)
 {
 Nodes->Delete();
 Branches->Active=false;
 Branches->CommandText="Select * From �����_7";
 Branches->Active=true;
 Branches->Append();
 Branches->FieldByName("����� ��������")->Value=NumParent;
 Branches->FieldByName("��������")->Value=Text;
 Branches->Post();
 Branches->Last();
 int Number=Branches->FieldByName("����� �����")->Value;
 N=TreeView5->GetNodeAt(XX,YY);
 TTreeNode * ParNode=N->Parent;
 PMyNode(ParNode->Data)->Number=Number;
 PMyNode(ParNode->Data)->Parent=NumParent;
 PMyNode(ParNode->Data)->Node=false;

}
}
}
TreeView5->Items->Delete(TreeView5->GetNodeAt(XX,YY));

}
//---------------------------------------------------------------------------


void __fastcall TDocuments::FormShow(TObject *Sender)
{
//ShowMessage("��������");
//Pass->Hide();
//Pass->Close();
//Zast->MClient->Start();
Zast->MClient->RegForm(this);

//Pass->ShowModal();
/*
NameForm(this, "��������� ����������");
*/


/////
ADODataSet1->Active=false;
ADODataSet1->CommandText="select * from ���������� Order By [��� �������]";
ADODataSet1->Active=true;
/////
//***
Initialize();
//***
//PageControl1->TabIndex=0;
PageControl1->ActivePageIndex=0;

Metod->Active=true;
ReadMetod->Checked=false;
DBMemo1->ReadOnly=!ReadMetod->Checked;

 Zast->MClient->WriteDiaryEvent("Hazards","������ ������ ��������� ��������","");

Zast->MClient->Stop();


}
//---------------------------------------------------------------------------
String TDocuments::LoadAspects()
{
 Zast->MClient->WriteDiaryEvent("Hazards","������ ������ (��������)","");
 String Ret="������";
try
{

Zast->MClient->PrepareLoadAspects(0, 497, 473, 243, 243);
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * From TempAspects";
Comm->Execute();

MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOAspect;
Local->CommandText="SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[������������ ����������], TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.�����,  TempAspects.�����������, TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects ORDER BY TempAspects.[����� �������];";
Local->Active=true;

Table* Server=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

Server->SetCommandText("SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[������������ ����������], TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.�����,  TempAspects.�����������, TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] FROM TempAspects ORDER BY TempAspects.[����� �������];");
Server->Active(true);
FProg->Label1->Caption="�������� ������";

Zast->MClient->LoadTable(Server, Local, FProg->Label1, FProg->Progress);

if(Zast->MClient->VerifyTable(Local, Server)==0)
{

MP<TADODataSet>Memo1(this);
Memo1->Connection=Zast->ADOAspect;
Memo1->CommandText="SELECT Memo1.Num, Memo1.NumStr, Memo1.Text FROM Memo1 ORDER BY Memo1.Num, Memo1.NumStr;";
Memo1->Active=true;

Table* SMemo1=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

SMemo1->SetCommandText("SELECT Memo1.Num, Memo1.NumStr, Memo1.Text FROM Memo1 ORDER BY Memo1.Num, Memo1.NumStr;");
SMemo1->Active(true);

Zast->MClient->LoadTable(SMemo1, Memo1);

if(Zast->MClient->VerifyTable(Memo1, SMemo1)==0)
{

MP<TADODataSet>Memo2(this);
Memo2->Connection=Zast->ADOAspect;
Memo2->CommandText="SELECT Memo2.Num, Memo2.NumStr, Memo2.Text FROM Memo2 ORDER BY Memo2.Num, Memo2.NumStr;";
Memo2->Active=true;

Table* SMemo2=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

SMemo2->SetCommandText("SELECT Memo2.Num, Memo2.NumStr, Memo2.Text FROM Memo2 ORDER BY Memo2.Num, Memo2.NumStr;");
SMemo2->Active(true);

Zast->MClient->LoadTable(SMemo2, Memo2);

if(Zast->MClient->VerifyTable(Memo2, SMemo2)==0)
{

MP<TADODataSet>Memo3(this);
Memo3->Connection=Zast->ADOAspect;
Memo3->CommandText="SELECT Memo3.Num, Memo3.NumStr, Memo3.Text FROM Memo3 ORDER BY Memo3.Num, Memo3.NumStr;";
Memo3->Active=true;

Table* SMemo3=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

SMemo3->SetCommandText("SELECT Memo3.Num, Memo3.NumStr, Memo3.Text FROM Memo3 ORDER BY Memo3.Num, Memo3.NumStr;");
SMemo3->Active(true);

Zast->MClient->LoadTable(SMemo3, Memo3);

if(Zast->MClient->VerifyTable(Memo3, SMemo3)==0)
{

MP<TADODataSet>Memo4(this);
Memo4->Connection=Zast->ADOAspect;
Memo4->CommandText="SELECT Memo4.Num, Memo4.NumStr, Memo4.Text FROM Memo4 ORDER BY Memo4.Num, Memo4.NumStr;";
Memo4->Active=true;

Table* SMemo4=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

SMemo4->SetCommandText("SELECT Memo4.Num, Memo4.NumStr, Memo4.Text FROM Memo4 ORDER BY Memo4.Num, Memo4.NumStr;");
SMemo4->Active(true);

Zast->MClient->LoadTable(SMemo4, Memo4);

if(Zast->MClient->VerifyTable(Memo4, SMemo4)==0)
{
Comm->CommandText="Delete * from Temp�������������";
Comm->Execute();



PrepareMergeAspects();
try
{
MergeAspects();
}
catch(...)
{
Zast->MClient->DeleteTable(this, SMemo4);
Zast->MClient->DeleteTable(this, SMemo3);
Zast->MClient->DeleteTable(this, SMemo2);
Zast->MClient->DeleteTable(this, SMemo1);
Zast->MClient->DeleteTable(this, Server);

 return Ret;
}
Ret="���������";

}
Zast->MClient->DeleteTable(this, SMemo4);
}
Zast->MClient->DeleteTable(this, SMemo3);
}
Zast->MClient->DeleteTable(this, SMemo2);

}
Zast->MClient->DeleteTable(this, SMemo1);
}
Zast->MClient->DeleteTable(this, Server);

}
catch(...)
{
 Zast->MClient->WriteDiaryEvent("������ NetAspects","������ �������� ������ (��������)"," ������ "+IntToStr(GetLastError()));


}
 return Ret;

}
//---------------------------------------------------------------------------
void TDocuments::PrepareMergeAspects()
{
 Zast->MClient->WriteDiaryEvent("Hazards","������ ���������� ����������� ������ (��������)","");

try
{
MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOAspect;
Temp->CommandText="Select * From TempAspects order by [����� �������]";
Temp->Active=true;

MP<TADODataSet>Memo(this);
Memo->Connection=Zast->ADOAspect;


for(Temp->First();!Temp->Eof;Temp->Next())
{
 int TempNum=Temp->FieldByName("����� �������")->Value;

MP<TMemo>M(this);
M->Visible=false;
M->Parent=this;
M->Width=1009;

TStrings* TT=M->Lines;
TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo1 Where Num="+IntToStr(TempNum);
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("������������� �����������")->Assign(TT);
Temp->Post();


TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo2 Where Num="+IntToStr(TempNum);
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("������������ �����������")->Assign(TT);
Temp->Post();


TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo3 Where Num="+IntToStr(TempNum);
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("���������� � ��������")->Assign(TT);
Temp->Post();

TT->Clear();

Memo->Active=false;
Memo->CommandText="Select * from Memo4 Where Num="+IntToStr(TempNum);
Memo->Active=true;

for(Memo->First();!Memo->Eof;Memo->Next())
{
String S=Memo->FieldByName("Text")->AsString;
TT->Append(S);
}
Temp->Edit();
Temp->FieldByName("������������ ���������� � ��������")->Assign(TT);
Temp->Post();
}
}
catch(...)
{
 Zast->MClient->WriteDiaryEvent("������ NetAspects","������ ���������� � ����������� ������ (��������)"," ������ "+IntToStr(GetLastError()));
}

}
//---------------------------------------------------------------------------
void TDocuments::MergeAspects()
{
 Zast->MClient->WriteDiaryEvent("Hazards","����������� ������ (��������)","");

try
{
MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOAspect;
Temp->CommandText="Select * From TempAspects";
Temp->Active=true;

MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOAspect;
Podr->CommandText="Select * From �������������";
Podr->Active=true;

for(Temp->First();!Temp->Eof;Temp->Next())
{
 int N=Temp->FieldByName("�������������")->Value;

 if(Podr->Locate("ServerNum", N, SO))
 {
  int Num=Podr->FieldByName("����� �������������")->Value;

  Temp->Edit();
  Temp->FieldByName("�������������")->Value=Num;
  Temp->Post();
 }
 else
 {
  ShowMessage("������ ����������� ������");
 }
}

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;



Comm->CommandText="Delete * From �������";
Comm->Execute();

/*
String CT="INSERT INTO ������� ( [����� �������], �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, G, O, R, S, T, L, N, Z, ����������, [������������ ����������] [���������� �����������], [������� �����������], ��������������, �����, [������������� �����������], [������������ �����������], [���������� � ��������], [������������ ���������� � ��������], �����������, [���� ��������], [������ ��������], [����� ��������] ) ";
CT=CT+" SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[������������ ����������],  TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.�����, TempAspects.[������������� �����������], TempAspects.[������������ �����������], TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������], TempAspects.�����������, TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] ";
CT=CT+" FROM TempAspects;";
*/
String CT="INSERT INTO ������� ( [����� �������], �������������, ��������, [��� ����������], ������������, �������������, ������, �����������, G, O, R, S, T, L, N, Z, ����������, [������������ ����������], [���������� �����������], [������� �����������], ��������������, �����, [������������� �����������], [������������ �����������], [���������� � ��������], [������������ ���������� � ��������], �����������, [���� ��������], [������ ��������], [����� ��������] ) ";
CT=CT+" SELECT TempAspects.[����� �������], TempAspects.�������������, TempAspects.��������, TempAspects.[��� ����������], TempAspects.������������, TempAspects.�������������, TempAspects.������, TempAspects.�����������, TempAspects.G, TempAspects.O, TempAspects.R, TempAspects.S, TempAspects.T, TempAspects.L, TempAspects.N, TempAspects.Z, TempAspects.����������, TempAspects.[������������ ����������], TempAspects.[���������� �����������], TempAspects.[������� �����������], TempAspects.��������������, TempAspects.�����, TempAspects.[������������� �����������], TempAspects.[������������ �����������], TempAspects.[���������� � ��������], TempAspects.[������������ ���������� � ��������], TempAspects.�����������, TempAspects.[���� ��������], TempAspects.[������ ��������], TempAspects.[����� ��������] ";
CT=CT+" FROM TempAspects";

Comm->CommandText=CT;
Comm->Execute();

Comm->CommandText="Delete * From TempAspects";
Comm->Execute();
}
catch(...)
{
 Zast->MClient->WriteDiaryEvent("������ NetAspects","������ ����������� ������ ����������"," ������ "+IntToStr(GetLastError()));
 throw 1;
}
}
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------
void __fastcall TDocuments::FormActivate(TObject *Sender)
{
this->Top=0;
this->Left=0;
this->Width=1024;
this->Height=742;

Podr->Active=false;
Podr->Active=true;
Sit->Active=false;
Sit->Active=true;

}
//---------------------------------------------------------------------------











void TDocuments::Crit(TDataSet *DataSet)
{
int NR=DataSet->RecNo;

while(!DataSet->Eof)
{
 DataSet->Edit();
 DataSet->FieldByName("��������1")->Value="��";
 DataSet->FieldByName("��������")->Value=true;
 DataSet->Post();
 DataSet->Next();
}
DataSet->First();
DataSet->MoveBy(NR-1);
}
//----------------------------------------------
void TDocuments::Crit1(TDataSet *DataSet)
{

int NR=DataSet->RecNo;

while(!DataSet->Bof)
{
 DataSet->Edit();
 DataSet->FieldByName("��������1")->Value="���";
 DataSet->FieldByName("��������")->Value=false;
 DataSet->Post();
 DataSet->Prior();
}
DataSet->First();
DataSet->MoveBy(NR-1);
}

//---------------------------
void TDocuments::Crit2(TDataSet *DataSet)
{
int NR=DataSet->RecNo;
ADODataSet1->Active=false;
ADODataSet1->CommandText="select * from ���������� Order By [��� �������]";
ADODataSet1->Active=true;
DataSet->Last();



double Min=DataSet->FieldByName("��� �������")->Value;
bool Cr=DataSet->FieldByName("��������")->Value;

DataSet->Edit();
DataSet->FieldByName("���� �������")->Value=Min+1;
DataSet->Post();

DataSet->Prior();
for(int i=1;i<DataSet->RecordCount;i++)
{
DataSet->Edit();
DataSet->FieldByName("���� �������")->Value=Min-1;
if(Cr==false)
{
DataSet->FieldByName("��������")->Value=false;
DataSet->FieldByName("��������1")->Value="���";

}
DataSet->Post();
Min=DataSet->FieldByName("��� �������")->Value;
Cr=DataSet->FieldByName("��������")->Value;
DataSet->Prior();
}
DataSet->First();
DataSet->Edit();
DataSet->FieldByName("��� �������")->Value=0;
DataSet->Post();
ADODataSet1->Active=false;
ADODataSet1->CommandText="select * from ���������� Order By [��� �������]";
ADODataSet1->Active=true;
DataSet->First();

//bool K=false;
for(int i=0;i<DataSet->RecordCount;i++)
{
if(DataSet->FieldByName("��� �������")->Value>DataSet->FieldByName("���� �������")->Value)
{
 ShowMessage("� ������ ��� ���� ���������� ������� ������ "+DataSet->FieldByName("��� �������")->Value+" ��������� ��");
//K=true;

break;
}
ADODataSet1->Next();
}

DataSet->First();
DataSet->MoveBy(NR-1);



}
//---------------------------------------------------------------------------
void __fastcall TDocuments::ADODataSet1BeforePost(TDataSet *DataSet)
{
if(C==false)
{
AnsiString A=DataSet->FieldByName("��������1")->Value;
if(A=="��")
{
 DataSet->FieldByName("��������")->Value=true;
 Kriteriy=true;
}
else
{
 DataSet->FieldByName("��������")->Value=false;
 Kriteriy=false;
}

}        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::ADODataSet1AfterPost(TDataSet *DataSet)
{
if(C==false)
{

 /*
ADODataSet1->Active=false;
ADODataSet1->CommandText="select * from ���������� Order By [��� �������]";
ADODataSet1->Active=true;
 */
bool B=Kriteriy;
if(B==true)
{
C=true;
Crit(DataSet);
C=false;
}
else
{
C=true;
Crit1(DataSet);
C=false;
}

C=true;
Crit2(DataSet);
C=false;
}
//Button1->Enabled=true;        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::MenuItem1Click(TObject *Sender)
{
// ��������
Ins=true;
ADODataSet1->Last();
int Max=ADODataSet1->FieldByName("���� �������")->Value;
bool Cr=ADODataSet1->FieldByName("��������")->Value;
ADODataSet1->Append();
ADODataSet1->FieldByName("��� �������")->Value=Max;
ADODataSet1->FieldByName("���� �������")->Value=Max+1;
ADODataSet1->FieldByName("��������")->Value=Cr;
ADODataSet1->FieldByName("������������ ����������")->Value="����� ����������";
ADODataSet1->FieldByName("����������� ����")->Value="����� ����������� ����";

if (Cr==true)
{
ADODataSet1->FieldByName("��������1")->Value="��";
}
else
{
ADODataSet1->FieldByName("��������1")->Value="���";
}

ADODataSet1->Post();
ADODataSet1->Last();
Ins=false;        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::MenuItem2Click(TObject *Sender)
{
// �������

ADODataSet1->Delete();
ADODataSet2->Active=false;
ADODataSet2->CommandText="Select * From ���������� Order By [��� �������]";
ADODataSet2->Active=true;
ADODataSet2->Last();

double Min=ADODataSet2->FieldByName("��� �������")->Value;

ADODataSet2->Prior();
for(int i=1;i<ADODataSet2->RecordCount;i++)
{
ADODataSet2->Edit();
// int Min=ADODataSet1->FieldByName("��� �������")->Value;


ADODataSet2->FieldByName("���� �������")->Value=Min-1;
Min=ADODataSet2->FieldByName("��� �������")->Value;
ADODataSet2->Post();
ADODataSet2->Prior();
}
ADODataSet1->Active=false;
ADODataSet1->Active=true;        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::PopupMenu6Popup(TObject *Sender)
{
if(ADODataSet1->RecordCount>2)
{
  PopupMenu1->Items->Items[0]->Enabled=true;
  PopupMenu1->Items->Items[1]->Enabled=true;
}
else
{
  PopupMenu1->Items->Items[0]->Enabled=true;
  PopupMenu1->Items->Items[1]->Enabled=false;
}        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::ADODataSet1AfterInsert(TDataSet *DataSet)
{
if(Ins==false)
{
DataSet->Prior();
DataSet->Next();
}
}
//---------------------------------------------------------------------------
void TDocuments::Initialize()
{
Podr->Active=true;
Sit->Active=true;
}

//---------------------------------------------------------------------------

void __fastcall TDocuments::MenuItem3Click(TObject *Sender)
{
Podr->Insert();
Podr->FieldByName("�������� �������������")->Value="����� �������������";
Podr->Post();


}
//---------------------------------------------------------------------------

void __fastcall TDocuments::MenuItem4Click(TObject *Sender)
{
int NumP=Podr->FieldByName("ServerNum")->AsInteger;

Zast->MClient->Start();
Table* STab=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

STab->SetCommandText("SELECT �������.[����� �������], �������.������������� FROM ������� WHERE (((�������.�������������)="+IntToStr(NumP)+"));");
STab->Active(true);

int N=0;
N=STab->RecordCount();

if(N==0)
{
Podr->Delete();
}
else
{
String S1;
int N1=N-((int)(N/10))*10;
switch(N1)
{
 case 1:
 {
 S1="�� ������� ������������ "+IntToStr(N)+" ����, ������������� ����� �������������.";
 break;
 }
 case 2: case 3: case 4:
 {
 S1="�� ������� ������������� "+IntToStr(N)+" �����, ������������� ����� �������������.";
 break;
 }
 default:
 {
 S1="�� ������� ������������� "+IntToStr(N)+" ������, ������������� ����� �������������.";
 }
}

String S=S1+"\r��������� \"�������� ������\" �������������� ���������� ��������� ������������� �� ������";
Application->MessageBoxA(S.c_str(),"�������� �������������",MB_ICONEXCLAMATION);
}
Zast->MClient->DeleteTable(this, STab);
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::MenuItem5Click(TObject *Sender)
{
Sit->Insert();
Sit->FieldByName("�������� ��������")->Value="����� ��������";
Sit->Post();        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::MenuItem6Click(TObject *Sender)
{

Sit->Delete();
}
//---------------------------------------------------------------------------


void __fastcall TDocuments::N6Click(TObject *Sender)
{
//������ ������� ����� �����������
//����_3 � �����_3
Zast->MClient->Start();
ShowMessage(ReadVidVozd(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String  TDocuments::ReadVidVozd(bool Pr)
{
if(Pr)
{
FProg->Label1->Caption="������ ����� �����������";
}
if(LoadNodeBranch("����_3","�����_3", Pr))
{
//LoadDirec("�����_3","�����������");
LoadTab1();
LoadDirVozd();

return "���������";

}
else
{
return "������";
}

}
//-----------------------------------
void TDocuments::LoadDirVozd()
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * From TempVozd";
Comm->Execute();

MP<TADODataSet>TempVozd(this);
TempVozd->Connection=Zast->ADOAspect;
TempVozd->CommandText="Select TempVozd.[����� �����������], TempVozd.[������������ �����������], TempVozd.[�����] From TempVozd order by [����� �����������]";
TempVozd->Active=true;

Table* SVozd=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

SVozd->SetCommandText("Select �����������.[����� �����������], �����������.[������������ �����������], �����������.[�����] From ����������� order by [����� �����������]");
SVozd->Active(true);

Zast->MClient->LoadTable(SVozd, TempVozd);

if(Zast->MClient->VerifyTable(SVozd, TempVozd)==0)
{

Comm->CommandText="UPDATE ����������� SET �����������.Del = False;";
Comm->Execute();

MP<TADODataSet>Vozd(this);
Vozd->Connection=Zast->ADOAspect;
Vozd->CommandText="Select * From �����������";
Vozd->Active=true;

 for(Vozd->First();!Vozd->Eof;Vozd->Next())
 {
  int N=Vozd->FieldByName("����� �����������")->Value;

  if(TempVozd->Locate("����� �����������", N, SO))
  {
   Vozd->Edit();
   Vozd->FieldByName("������������ �����������")->Value=TempVozd->FieldByName("������������ �����������")->Value;
   Vozd->FieldByName("�����")->Value=TempVozd->FieldByName("�����")->Value;
   Vozd->Post();

   TempVozd->Delete();
  }
  else
  {
   Vozd->Edit();
   Vozd->FieldByName("Del")->Value=true;
   Vozd->Post();
  }
 }
Comm->CommandText="DELETE �����������.Del FROM ����������� WHERE (((�����������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ����������� ( [����� �����������], [������������ �����������], ����� ) SELECT TempVozd.[����� �����������], TempVozd.[������������ �����������], TempVozd.����� FROM TempVozd;";
Comm->Execute();

Comm->CommandText="DELETE TempVozd.* FROM TempVozd;";
Comm->Execute();

}
Zast->MClient->DeleteTable(this, SVozd);
/*
 MP<TADODataSet>Branch(this);
 Branch->Connection=Zast->ADOConn;
 Branch->CommandText="Select * From �����_3 Order by [����� �����]";
 Branch->Active=true;

 MP<TADODataSet>Vozd(this);
 Vozd->Connection=Zast->ADOAspect;
 Vozd->CommandText="Select * From ����������� Order by [����� �����������]";
 Vozd->Active=true;

 MP<TADOCommand>Comm(this);
 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="UPDATE ����������� SET �����������.Del = False;";
 Comm->Execute();

 Comm->Connection=Zast->ADOConn;
 Comm->CommandText="UPDATE �����_3 SET �����_3.Del = False;";
 Comm->Execute();

 for(Branch->First();!Branch->Eof;Branch->Next())
 {
  int N=Branch->FieldByName("����� �����")->Value;

  if(Branch->FieldByName("�����")->AsBoolean)
  {
  if(Vozd->Locate("����� �����������",N,SO))
  {
   Vozd->Edit();
   Vozd->FieldByName("������������ �����������")->Value=Branch->FieldByName("��������")->Value;
   Vozd->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Vozd->FieldByName("Del")->Value=true;
   Vozd->Post();
  }
  else
  {
   Vozd->Insert();
   Vozd->FieldByName("����� �����������")->Value=Branch->FieldByName("����� �����")->Value;
   Vozd->FieldByName("������������ �����������")->Value=Branch->FieldByName("��������")->Value;
   Vozd->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Vozd->FieldByName("Del")->Value=true;
   Vozd->Post();
  }
  }
  else
  {
  if(Vozd->Locate("����� �����������",0,SO))
  {
   Vozd->Edit();
   Vozd->FieldByName("������������ �����������")->Value=Branch->FieldByName("��������")->Value;
   Vozd->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Vozd->FieldByName("Del")->Value=true;
   Vozd->Post();
  }
  else
  {
   Vozd->Insert();
   Vozd->FieldByName("����� �����������")->Value=0;
   Vozd->FieldByName("������������ �����������")->Value=Branch->FieldByName("��������")->Value;
   Vozd->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Vozd->FieldByName("Del")->Value=true;
   Vozd->Post();
  }
  }
 }
 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="UPDATE ����������� INNER JOIN ������� ON �����������.[����� �����������] = �������.����������� SET �������.����������� = 0 WHERE (((�����������.Del)=False));";
 Comm->Execute();

 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="Delete * from ����������� Where Del=false";
 Comm->Execute();


 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="UPDATE ����������� SET �����������.Del = False;";
 Comm->Execute();
 */
}
//------------------------------------
bool TDocuments::LoadNodeBranch(String NodeTable, String BranchTable, bool Pr)
{
//������ ������� ����� �����������
//����_3 � �����_3
bool Ret=false;
Zast->MClient->WriteDiaryEvent("Hazards","������ �������� ����� � ������ ������������ (��������)","����: "+NodeTable+" �����: "+BranchTable);
 try
 {
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
Comm->CommandText="Delete * From TempNode";
Comm->Execute();

Comm->CommandText="Delete * From TempBranch";
Comm->Execute();

MP<TADODataSet>TempNode(this);
TempNode->Connection=Zast->ADOConn;
TempNode->CommandText="Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by ��������, [����� ����]";
TempNode->Active=true;



Table* SNode=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("HReference")].ServerDB);

SNode->SetCommandText("Select "+NodeTable+".[����� ����], "+NodeTable+".[��������], "+NodeTable+".[��������] From "+NodeTable+" Order by ��������, [����� ����]");
SNode->Active(true);

Zast->MClient->LoadTable(SNode, TempNode);

if(Zast->MClient->VerifyTable(SNode, TempNode)==0)
{
MP<TADODataSet>TempBranch(this);
TempBranch->Connection=Zast->ADOConn;
TempBranch->CommandText="Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����]";
TempBranch->Active=true;


Table* SBranch=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("HReference")].ServerDB);

SBranch->SetCommandText("Select "+BranchTable+".[����� �����], "+BranchTable+".[����� ��������], "+BranchTable+".[��������], "+BranchTable+".[�����] From "+BranchTable+" Order by [����� ��������], [����� �����]");
SBranch->Active(true);

if(Pr)
{
Zast->MClient->LoadTable(SBranch, TempBranch, FProg->Label1, FProg->Progress);
}
else
{
Zast->MClient->LoadTable(SBranch, TempBranch);
}

if(Zast->MClient->VerifyTable(SBranch, TempBranch)==0)
{

MergeNode(NodeTable);
MergeBranch(BranchTable);
Zast->MClient->WriteDiaryEvent("Hazards","���������� �������� ����� � ������ ������������ (��������)","����: "+NodeTable+" �����: "+BranchTable);

Ret=true;
}

Zast->MClient->DeleteTable(this, SBranch);

}

Zast->MClient->DeleteTable(this, SNode);


Comm->CommandText="Delete * From TempNode";
Comm->Execute();

Comm->CommandText="Delete * From TempBranch";
Comm->Execute();

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","���� �������� ����� � ������ ������������ (��������)","����: "+NodeTable+" �����: "+BranchTable);
}

}
catch(...)
{
 Zast->MClient->WriteDiaryEvent("������ NetAspects ","������ �������� ����� � ������ ������������ (��������)","����: "+NodeTable+" �����: "+BranchTable+" ������ "+IntToStr(GetLastError()));
}
return Ret;
}
//-------------------------------------------------
void TDocuments::MergeNode(String NodeTable)
{
Zast->MClient->WriteDiaryEvent("Hazards","������ ����������� ����� ������������ (��������)","����: "+NodeTable);

try
{
TLocateOptions SO;
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
Comm->CommandText="UPDATE "+NodeTable+" SET "+NodeTable+".Del = False;";
Comm->Execute();

MP<TADODataSet>TempNode(this);
TempNode->Connection=Zast->ADOConn;
TempNode->CommandText="Select * From TempNode";
TempNode->Active=true;

MP<TADODataSet>Node(this);
Node->Connection=Zast->ADOConn;
Node->CommandText="Select * From "+NodeTable;
Node->Active=true;

for(Node->First();!Node->Eof;Node->Next())
{
int NumNode=Node->FieldByName("����� ����")->Value;
if(TempNode->Locate("����� ����", NumNode, SO))
{
 //������� ������������
 Node->Edit();
 Node->FieldByName("��������")->Value=TempNode->FieldByName("��������")->Value;
 Node->FieldByName("��������")->Value=TempNode->FieldByName("��������")->Value;
 Node->Post();

 TempNode->Delete();
}
else
{
 //��� ������������
 Node->Edit();
 Node->FieldByName("Del")->Value=true;
 Node->Post();
}
}

Comm->CommandText="Delete * From "+NodeTable+" Where Del=true";
Comm->Execute();

Comm->CommandText="INSERT INTO "+NodeTable+" ( [����� ����], ��������, ��������) SELECT TempNode.[����� ����], TempNode.��������, TempNode.�������� FROM TempNode;";
Comm->Execute();

Comm->CommandText="Delete * From TempNode";
Comm->Execute();

Zast->MClient->WriteDiaryEvent("Hazards","���������� ����������� ����� ������������ (��������)","����: "+NodeTable);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","������ ����������� ����� ������������ (��������)","����: "+NodeTable+" ������ "+IntToStr(GetLastError()));
}
}
//----------------------------------------------------
void TDocuments::MergeBranch(String BranchTable)
{
Zast->MClient->WriteDiaryEvent("Hazards","������ ����������� ������ ������������ (��������)"," �����: "+BranchTable);

try
{
TLocateOptions SO;
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
Comm->CommandText="UPDATE "+BranchTable+" SET "+BranchTable+".Del = False;";
Comm->Execute();

MP<TADODataSet>TempBranch(this);
TempBranch->Connection=Zast->ADOConn;
TempBranch->CommandText="Select * From TempBranch";
TempBranch->Active=true;

MP<TADODataSet>Branch(this);
Branch->Connection=Zast->ADOConn;
Branch->CommandText="Select * From "+BranchTable;
Branch->Active=true;

for(Branch->First();!Branch->Eof;Branch->Next())
{
int NumBranch=Branch->FieldByName("����� �����")->Value;
if(TempBranch->Locate("����� �����", NumBranch, SO))
{
 //������� ������������
 Branch->Edit();
 Branch->FieldByName("����� ��������")->Value=TempBranch->FieldByName("����� ��������")->Value;
 Branch->FieldByName("��������")->Value=TempBranch->FieldByName("��������")->Value;
 Branch->FieldByName("�����")->Value=TempBranch->FieldByName("�����")->Value;
 Branch->Post();

 TempBranch->Delete();
}
else
{
 //��� ������������
 Branch->Edit();
 Branch->FieldByName("Del")->Value=true;
 Branch->Post();
}
}

Comm->CommandText="Delete * From "+BranchTable+" Where Del=true";
Comm->Execute();

Comm->CommandText="INSERT INTO "+BranchTable+" ( [����� �����], [����� ��������], ��������, ����� ) SELECT TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.��������, TempBranch.����� FROM TempBranch;";
Comm->Execute();

Comm->CommandText="Delete * From TempBranch";
Comm->Execute();

Zast->MClient->WriteDiaryEvent("Hazards","����� ����������� ������ ������������ (��������)"," �����: "+BranchTable);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","������ ����������� ������ ������������ (��������)"," �����: "+BranchTable+" ������ "+IntToStr(GetLastError()));
}
}
//----------------------------------------------------
void __fastcall TDocuments::N7Click(TObject *Sender)
{
//������ ������� ��������������� �����������
//����_4 � �����_4
Zast->MClient->Start();
ShowMessage(ReadMeropr(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String TDocuments::ReadMeropr(bool Pr)
{
if(Pr)
{
FProg->Label1->Caption="������ ��������������� �����������";
}
if(LoadNodeBranch("����_4","�����_4", Pr))
{
LoadTab2();

return "���������";

}
else
{
return "������";
}
}
//--------------------------------------------------------

void __fastcall TDocuments::N20Click(TObject *Sender)
{
//������ ������� ����� ����������
//����_5 � �����_5
Zast->MClient->Start();
ShowMessage(ReadTerr(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String TDocuments::ReadTerr(bool Pr)
{
if(Pr)
{
FProg->Label1->Caption="������ ����� ����������";
}
if(LoadNodeBranch("����_5","�����_5", Pr))
{
LoadTab3();
LoadDirTerr();
return "���������";

}
else
{
return "������";
}
}
//------------------------------------------------------------
void TDocuments::LoadDirTerr()
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * From TempTerr";
Comm->Execute();

MP<TADODataSet>TempTerr(this);
TempTerr->Connection=Zast->ADOAspect;
TempTerr->CommandText="Select TempTerr.[����� ����������], TempTerr.[������������ ����������], TempTerr.[�����] From TempTerr order by [����� ����������]";
TempTerr->Active=true;

Table* STerr=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

STerr->SetCommandText("Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[�����] From ���������� order by [����� ����������]");
STerr->Active(true);

Zast->MClient->LoadTable(STerr, TempTerr);

if(Zast->MClient->VerifyTable(STerr, TempTerr)==0)
{

Comm->CommandText="UPDATE ���������� SET ����������.Del = False;";
Comm->Execute();

MP<TADODataSet>Terr(this);
Terr->Connection=Zast->ADOAspect;
Terr->CommandText="Select * From ����������";
Terr->Active=true;

 for(Terr->First();!Terr->Eof;Terr->Next())
 {
  int N=Terr->FieldByName("����� ����������")->Value;

  if(TempTerr->Locate("����� ����������", N, SO))
  {
   Terr->Edit();
   Terr->FieldByName("������������ ����������")->Value=TempTerr->FieldByName("������������ ����������")->Value;
   Terr->FieldByName("�����")->Value=TempTerr->FieldByName("�����")->Value;
   Terr->Post();

   TempTerr->Delete();
  }
  else
  {
   Terr->Edit();
   Terr->FieldByName("Del")->Value=true;
   Terr->Post();
  }
 }
Comm->CommandText="DELETE ����������.Del FROM ���������� WHERE (((����������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ���������� ( [����� ����������], [������������ ����������], ����� ) SELECT TempTerr.[����� ����������], TempTerr.[������������ ����������], TempTerr.����� FROM TempTerr;";
Comm->Execute();

Comm->CommandText="DELETE TempTerr.* FROM TempTerr;";
Comm->Execute();

}
Zast->MClient->DeleteTable(this, STerr);
/*
 MP<TADODataSet>Branch(this);
 Branch->Connection=Zast->ADOConn;
 Branch->CommandText="Select * From �����_5 Order by [����� �����]";
 Branch->Active=true;

 MP<TADODataSet>Terr(this);
 Terr->Connection=Zast->ADOAspect;
 Terr->CommandText="Select * From ���������� Order by [����� ����������]";
 Terr->Active=true;

 MP<TADOCommand>Comm(this);
 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="UPDATE ���������� SET ����������.Del = False;";
 Comm->Execute();

 Comm->Connection=Zast->ADOConn;
 Comm->CommandText="UPDATE �����_5 SET �����_5.Del = False;";
 Comm->Execute();

 for(Branch->First();!Branch->Eof;Branch->Next())
 {
  int N=Branch->FieldByName("����� �����")->Value;

  if(Branch->FieldByName("�����")->AsBoolean)
  {
  if(Terr->Locate("����� ����������",N,SO))
  {
   Terr->Edit();
   Terr->FieldByName("������������ ����������")->Value=Branch->FieldByName("��������")->Value;
   Terr->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Terr->FieldByName("Del")->Value=true;
   Terr->Post();
  }
  else
  {
   Terr->Insert();
   Terr->FieldByName("����� ����������")->Value=Branch->FieldByName("����� �����")->Value;
   Terr->FieldByName("������������ ����������")->Value=Branch->FieldByName("��������")->Value;
   Terr->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Terr->FieldByName("Del")->Value=true;
   Terr->Post();
  }
  }
  else
  {
  if(Terr->Locate("����� ����������",0,SO))
  {
   Terr->Edit();
   Terr->FieldByName("������������ ����������")->Value=Branch->FieldByName("��������")->Value;
   Terr->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Terr->FieldByName("Del")->Value=true;
   Terr->Post();
  }
  else
  {
   Terr->Insert();
   Terr->FieldByName("����� ����������")->Value=0;
   Terr->FieldByName("������������ ����������")->Value=Branch->FieldByName("��������")->Value;
   Terr->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Terr->FieldByName("Del")->Value=true;
   Terr->Post();
  }
  }
 }
 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="UPDATE ���������� INNER JOIN ������� ON ����������.[����� ����������] = �������.[��� ����������] SET �������.[��� ����������] = 0 WHERE (((����������.Del)=False));";
 Comm->Execute();

 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="Delete * from ���������� Where Del=false";
 Comm->Execute();


 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="UPDATE ���������� SET ����������.Del = False;";
 Comm->Execute();
 */
}
//------------------------------------
void __fastcall TDocuments::N24Click(TObject *Sender)
{
//������ ������� ����� ������������
//����_6 � �����_6
Zast->MClient->Start();
ShowMessage(ReadDeyat(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String TDocuments::ReadDeyat(bool Pr)
{
if(Pr)
{
FProg->Label1->Caption="������ ����� ������������";
}
if(LoadNodeBranch("����_6","�����_6", Pr))
{
LoadTab4();
LoadDirDeyat();
return "���������";

}
else
{
return "������";
}
}
//-----------------------------------------------------
void TDocuments::LoadDirDeyat()
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * From TempDeyat";
Comm->Execute();

MP<TADODataSet>TempDeyat(this);
TempDeyat->Connection=Zast->ADOAspect;
TempDeyat->CommandText="Select TempDeyat.[����� ������������], TempDeyat.[������������ ������������], TempDeyat.[�����] From TempDeyat order by [����� ������������]";
TempDeyat->Active=true;

Table* SDeyat=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

SDeyat->SetCommandText("Select ������������.[����� ������������], ������������.[������������ ������������], ������������.[�����] From ������������ order by [����� ������������]");
SDeyat->Active(true);

Zast->MClient->LoadTable(SDeyat, TempDeyat);

if(Zast->MClient->VerifyTable(SDeyat, TempDeyat)==0)
{

Comm->CommandText="UPDATE ������������ SET ������������.Del = False;";
Comm->Execute();

MP<TADODataSet>Deyat(this);
Deyat->Connection=Zast->ADOAspect;
Deyat->CommandText="Select * From ������������";
Deyat->Active=true;

 for(Deyat->First();!Deyat->Eof;Deyat->Next())
 {
  int N=Deyat->FieldByName("����� ������������")->Value;

  if(TempDeyat->Locate("����� ������������", N, SO))
  {
   Deyat->Edit();
   Deyat->FieldByName("������������ ������������")->Value=TempDeyat->FieldByName("������������ ������������")->Value;
   Deyat->FieldByName("�����")->Value=TempDeyat->FieldByName("�����")->Value;
   Deyat->Post();

   TempDeyat->Delete();
  }
  else
  {
   Deyat->Edit();
   Deyat->FieldByName("Del")->Value=true;
   Deyat->Post();
  }
 }
Comm->CommandText="DELETE ������������.Del FROM ������������ WHERE (((������������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ������������ ( [����� ������������], [������������ ������������], ����� ) SELECT TempDeyat.[����� ������������], TempDeyat.[������������ ������������], TempDeyat.����� FROM TempDeyat;";
Comm->Execute();

Comm->CommandText="DELETE TempDeyat.* FROM TempDeyat;";
Comm->Execute();

}
Zast->MClient->DeleteTable(this, SDeyat);

/*
 MP<TADODataSet>Branch(this);
 Branch->Connection=Zast->ADOConn;
 Branch->CommandText="Select * From �����_6 Order by [����� �����]";
 Branch->Active=true;

 MP<TADODataSet>Deyat(this);
 Deyat->Connection=Zast->ADOAspect;
 Deyat->CommandText="Select * From ������������ Order by [����� ������������]";
 Deyat->Active=true;

 MP<TADOCommand>Comm(this);
 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="UPDATE ������������ SET ������������.Del = False;";
 Comm->Execute();

 Comm->Connection=Zast->ADOConn;
 Comm->CommandText="UPDATE �����_6 SET �����_6.Del = False;";
 Comm->Execute();

 for(Branch->First();!Branch->Eof;Branch->Next())
 {
  int N=Branch->FieldByName("����� �����")->Value;

  if(Branch->FieldByName("�����")->AsBoolean)
  {
  if(Deyat->Locate("����� ������������",N,SO))
  {
   Deyat->Edit();
   Deyat->FieldByName("������������ ������������")->Value=Branch->FieldByName("��������")->Value;
   Deyat->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Deyat->FieldByName("Del")->Value=true;
   Deyat->Post();
  }
  else
  {
   Deyat->Insert();
   Deyat->FieldByName("����� ������������")->Value=Branch->FieldByName("����� �����")->Value;
   Deyat->FieldByName("������������ ������������")->Value=Branch->FieldByName("��������")->Value;
   Deyat->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Deyat->FieldByName("Del")->Value=true;
   Deyat->Post();
  }
  }
  else
  {
  if(Deyat->Locate("����� ������������",0,SO))
  {
   Deyat->Edit();
   Deyat->FieldByName("������������ ������������")->Value=Branch->FieldByName("��������")->Value;
   Deyat->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Deyat->FieldByName("Del")->Value=true;
   Deyat->Post();
  }
  else
  {
   Deyat->Insert();
   Deyat->FieldByName("����� ������������")->Value=0;
   Deyat->FieldByName("������������ ������������")->Value=Branch->FieldByName("��������")->Value;
   Deyat->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Deyat->FieldByName("Del")->Value=true;
   Deyat->Post();
  }
  }
 }
 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="UPDATE ������������ INNER JOIN ������� ON ������������.[����� ������������] = �������.[������������] SET �������.[������������] = 0 WHERE (((������������.Del)=False));";
 Comm->Execute();

 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="Delete * from ������������ Where Del=false";
 Comm->Execute();


 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="UPDATE ������������ SET ������������.Del = False;";
 Comm->Execute();
 */
}
//------------------------------------

void __fastcall TDocuments::N25Click(TObject *Sender)
{
//������ ������ ����������� ������������� ��������
//����_7 � �����_7
Zast->MClient->Start();
ShowMessage(ReadAspect(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String TDocuments::ReadAspect(bool Pr)
{
if(Pr)
{
FProg->Label1->Caption="������ ����������� ����������";
}
if(LoadNodeBranch("����_7","�����_7", Pr))
{
LoadTab5();
LoadDirAspect();
return "���������";

}
else
{
return "������";
}
}
//----------------------------------------------------
void TDocuments::LoadDirAspect()
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * From TempAspect";
Comm->Execute();

MP<TADODataSet>TempAspect(this);
TempAspect->Connection=Zast->ADOAspect;
TempAspect->CommandText="Select TempAspect.[����� �������], TempAspect.[������������ �������], TempAspect.[�����] From TempAspect order by [����� �������]";
TempAspect->Active=true;

Table* SDeyat=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

SDeyat->SetCommandText("Select ������.[����� �������], ������.[������������ �������], ������.[�����] From ������ order by [����� �������]");
SDeyat->Active(true);

Zast->MClient->LoadTable(SDeyat, TempAspect);

if(Zast->MClient->VerifyTable(SDeyat, TempAspect)==0)
{

Comm->CommandText="UPDATE ������ SET ������.Del = False;";
Comm->Execute();

MP<TADODataSet>Aspect(this);
Aspect->Connection=Zast->ADOAspect;
Aspect->CommandText="Select * From ������";
Aspect->Active=true;

 for(Aspect->First();!Aspect->Eof;Aspect->Next())
 {
  int N=Aspect->FieldByName("����� �������")->Value;

  if(TempAspect->Locate("����� �������", N, SO))
  {
   Aspect->Edit();
   Aspect->FieldByName("������������ �������")->Value=TempAspect->FieldByName("������������ �������")->Value;
   Aspect->FieldByName("�����")->Value=TempAspect->FieldByName("�����")->Value;
   Aspect->Post();

   TempAspect->Delete();
  }
  else
  {
   Aspect->Edit();
   Aspect->FieldByName("Del")->Value=true;
   Aspect->Post();
  }
 }
Comm->CommandText="DELETE ������.Del FROM ������ WHERE (((������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ������ ( [����� �������], [������������ �������], ����� ) SELECT TempAspect.[����� �������], TempAspect.[������������ �������], TempAspect.����� FROM TempAspect;";
Comm->Execute();

Comm->CommandText="DELETE TempDeyat.* FROM TempDeyat;";
Comm->Execute();

}
Zast->MClient->DeleteTable(this, SDeyat);

/*
 MP<TADODataSet>Branch(this);
 Branch->Connection=Zast->ADOConn;
 Branch->CommandText="Select * From �����_7 Order by [����� �����]";
 Branch->Active=true;

 MP<TADODataSet>Aspect(this);
 Aspect->Connection=Zast->ADOAspect;
 Aspect->CommandText="Select * From ������ Order by [����� �������]";
 Aspect->Active=true;

 MP<TADOCommand>Comm(this);
 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="UPDATE ������ SET ������.Del = False;";
 Comm->Execute();

 Comm->Connection=Zast->ADOConn;
 Comm->CommandText="UPDATE �����_7 SET �����_7.Del = False;";
 Comm->Execute();

 for(Branch->First();!Branch->Eof;Branch->Next())
 {
  int N=Branch->FieldByName("����� �����")->Value;

  if(Branch->FieldByName("�����")->AsBoolean)
  {
  if(Aspect->Locate("����� �������",N,SO))
  {
   Aspect->Edit();
   Aspect->FieldByName("������������ �������")->Value=Branch->FieldByName("��������")->Value;
   Aspect->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Aspect->FieldByName("Del")->Value=true;
   Aspect->Post();
  }
  else
  {
   Aspect->Insert();
   Aspect->FieldByName("����� �������")->Value=Branch->FieldByName("����� �����")->Value;
   Aspect->FieldByName("������������ �������")->Value=Branch->FieldByName("��������")->Value;
   Aspect->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Aspect->FieldByName("Del")->Value=true;
   Aspect->Post();
  }
  }
  else
  {
  if(Aspect->Locate("����� �������",0,SO))
  {
   Aspect->Edit();
   Aspect->FieldByName("������������ �������")->Value=Branch->FieldByName("��������")->Value;
   Aspect->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Aspect->FieldByName("Del")->Value=true;
   Aspect->Post();
  }
  else
  {
   Aspect->Insert();
   Aspect->FieldByName("����� �������")->Value=0;
   Aspect->FieldByName("������������ �������")->Value=Branch->FieldByName("��������")->Value;
   Aspect->FieldByName("�����")->Value=Branch->FieldByName("�����")->Value;
   Aspect->FieldByName("Del")->Value=true;
   Aspect->Post();
  }
  }
 }
 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="UPDATE ������ INNER JOIN ������� ON ������.[����� �������] = �������.[������] SET �������.[������] = 0 WHERE (((������.Del)=False));";
 Comm->Execute();

 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="Delete * from ������ Where Del=false";
 Comm->Execute();


 Comm->Connection=Zast->ADOAspect;
 Comm->CommandText="UPDATE ������ SET ������.Del = False;";
 Comm->Execute();
 */
}
//------------------------------------
void __fastcall TDocuments::N26Click(TObject *Sender)
{
//������ ���������
Zast->MClient->Start();
ShowMessage(ReadCrit(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String TDocuments::ReadCrit(bool Pr)
{
if(LoadCrit(Pr))
{
ADODataSet1->Active=false;
ADODataSet1->Active=true;
return "���������";

}
else
{
return "������";
}
}
//-----------------------------------------------
bool TDocuments::LoadCrit(bool Pr)
{
bool Ret=false;
Zast->MClient->WriteDiaryEvent("Hazards","������ �������� ��������� (��������)","");

try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
Comm->CommandText="Delete * From TempZn";
Comm->Execute();

MP<TADODataSet>TempZn(this);
TempZn->Connection=Zast->ADOConn;
TempZn->CommandText="Select TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.[��������1], TempZn.[��������], TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] From TempZn Order by [����� ����������]";
TempZn->Active=true;



Table* SZn=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("HReference")].ServerDB);

SZn->SetCommandText("Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[��������1], ����������.[��������], ����������.[��� �������], ����������.[���� �������], ����������.[����������� ����] From ���������� Order by [����� ����������]");
SZn->Active(true);
if(Pr)
{
FProg->Label1->Caption="������ ���������";
Zast->MClient->LoadTable(SZn, TempZn, FProg->Label1, FProg->Progress);
}
else
{
Zast->MClient->LoadTable(SZn, TempZn);
}
if(Zast->MClient->VerifyTable(SZn, TempZn)==0)
{
Comm->CommandText="Delete * From ����������";
Comm->Execute();

Comm->CommandText="INSERT INTO ���������� ( [����� ����������], [������������ ����������], ��������1, ��������, [��� �������], [���� �������], [����������� ����] ) SELECT TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.��������1, TempZn.��������, TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] FROM TempZn;";
Comm->Execute();
Zast->MClient->WriteDiaryEvent("Hazards","����� �������� ��������� (��������)","");

Ret=true;
}

Zast->MClient->DeleteTable(this, SZn);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","���� �������� ��������� (��������)","");
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","������ �������� ��������� (��������)"," ������ "+IntToStr(GetLastError()));
}
return Ret;
}
//---------------------------------------

void __fastcall TDocuments::N27Click(TObject *Sender)
{
//������ �������������
Zast->MClient->Start();
ShowMessage(ReadPodr(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String TDocuments::ReadPodr(bool Pr)
{
if(LoadPodr(Pr))
{
Podr->Active=false;
Podr->Active=true;
return "���������";

}
else
{
return "������";
}
}
//-------------------------------------------------------
bool TDocuments::LoadPodr(bool Pr)
{
bool Ret=false;
Zast->MClient->WriteDiaryEvent("Hazards","������ �������� ������������� (��������)","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * From Temp�������������";
Comm->Execute();

MP<TADODataSet>TempPodr(this);
TempPodr->Connection=Zast->ADOAspect;
TempPodr->CommandText="Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������] From Temp������������� order by [����� �������������]";
TempPodr->Active=true;



Table* SPodr=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

SPodr->SetCommandText("Select �������������.[����� �������������], �������������.[�������� �������������] From ������������� order by [����� �������������]");
SPodr->Active(true);
if(Pr)
{
FProg->Label1->Caption="������ �������������";
Zast->MClient->LoadTable(SPodr, TempPodr, FProg->Label1, FProg->Progress);
}
else
{
Zast->MClient->LoadTable(SPodr, TempPodr);
}

MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOAspect;
Podr->CommandText="Select * From �������������";
Podr->Active=true;

if(Zast->MClient->VerifyTable(SPodr, TempPodr)==0)
{
Comm->CommandText="UPDATE ������������� SET �������������.Del = False;";
Comm->Execute();

for(Podr->First();!Podr->Eof;Podr->Next())
{
 int N=Podr->FieldByName("ServerNum")->Value;
 if(TempPodr->Locate("����� �������������",N,SO))
 {
  Podr->Edit();
  Podr->FieldByName("�������� �������������")->Value=TempPodr->FieldByName("�������� �������������")->Value;
  Podr->Post();

  TempPodr->Delete();
 }
 else
 {
  Podr->Edit();
  Podr->FieldByName("Del")->Value=true;
  Podr->Post();
 }
}
Comm->CommandText="DELETE * FROM ������������� WHERE (((�������������.Del)=True));";
Comm->Execute();

Comm->CommandText="INSERT INTO ������������� ( ServerNum, [�������� �������������] ) SELECT Temp�������������.[����� �������������], Temp�������������.[�������� �������������] FROM Temp�������������;";
Comm->Execute();

Comm->CommandText="Delete * From Temp�������������";
Comm->Execute();

Zast->MClient->WriteDiaryEvent("Hazards","����� �������� ������������� (��������)","");
Ret=true;
}

Zast->MClient->DeleteTable(this, SPodr);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","������ �������� ������������� (��������)"," ������ "+IntToStr(GetLastError()));

}
return Ret;
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::N28Click(TObject *Sender)
{
//������ ��������
Zast->MClient->Start();
ShowMessage(ReadSit(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String TDocuments::ReadSit(bool Pr)
{
if(LoadSit(Pr))
{
Sit->Active=false;
Sit->Active=true;
return "���������";

}
else
{
return "������";
}
}
//-------------------------------------------
bool TDocuments::LoadSit(bool Pr)
{
bool Ret=false;
Zast->MClient->WriteDiaryEvent("Hazards","������ �������� �������� (��������)","");
try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
Comm->CommandText="Delete * From TempSit";
Comm->Execute();

MP<TADODataSet>TempSit(this);
TempSit->Connection=Zast->ADOConn;
TempSit->CommandText="Select TempSit.[����� ��������], TempSit.[�������� ��������] From TempSit order by [����� ��������]";
TempSit->Active=true;



Table* SSit=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("HReference")].ServerDB);

SSit->SetCommandText("Select ��������.[����� ��������], ��������.[�������� ��������] From �������� order by [����� ��������]");
SSit->Active(true);
if(Pr)
{
FProg->Label1->Caption="������ ��������";
Zast->MClient->LoadTable(SSit, TempSit, FProg->Label1, FProg->Progress);
}
else
{
Zast->MClient->LoadTable(SSit, TempSit);
}

if(Zast->MClient->VerifyTable(SSit, TempSit)==0)
{
Comm->CommandText="Delete * From ��������";
Comm->Execute();

Comm->CommandText="INSERT INTO �������� ( [����� ��������], [�������� ��������] ) SELECT TempSit.[����� ��������], TempSit.[�������� ��������] FROM TempSit;";
Comm->Execute();

//----
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * From TempSit";
Comm->Execute();

MP<TADODataSet>TempSit1(this);
TempSit1->Connection=Zast->ADOAspect;
TempSit1->CommandText="Select TempSit.[����� ��������], TempSit.[�������� ��������] From TempSit order by [����� ��������]";
TempSit1->Active=true;



Table* SSit1=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

SSit1->SetCommandText("Select ��������.[����� ��������], ��������.[�������� ��������] From �������� Where �����=True order by [����� ��������]");
SSit1->Active(true);

Zast->MClient->LoadTable(SSit1, TempSit1);

if(Zast->MClient->VerifyTable(SSit1, TempSit1)==0)
{

Comm->CommandText="Delete * From �������� Where �����=true";
Comm->Execute();

Comm->CommandText="INSERT INTO �������� ( [����� ��������], [�������� ��������], ����� ) SELECT TempSit.[����� ��������], TempSit.[�������� ��������], True FROM TempSit;";
Comm->Execute();

}
Zast->MClient->DeleteTable(this, SSit1);


Zast->MClient->WriteDiaryEvent("Hazards","����� �������� �������� (��������)","");
Ret=true;
}

Zast->MClient->DeleteTable(this, SSit);
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","������ �������� �������� (��������)"," ������ "+IntToStr(GetLastError()));

}
return Ret;
}
//----------------------------------------------------------------------------
void __fastcall TDocuments::DBMemo1Exit(TObject *Sender)
{
DataSetRefresh4->Execute();
DataSetPost1->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::DBGrid1Exit(TObject *Sender)
{
DataSetRefresh2->Execute();
DataSetPost2->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::ReadMetodClick(TObject *Sender)
{
DBMemo1->ReadOnly=!ReadMetod->Checked;        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N2Click(TObject *Sender)
{
//������ ��������

Zast->MClient->Start();
ShowMessage(ReadMet(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String TDocuments::ReadMet(bool Pr)
{
if(LoadMet(Pr))
{
Metod->Active=false;
Metod->Active=true;
return "���������";

}
else
{
return "������";
}
}
//-------------------------------------------------------
bool TDocuments::LoadMet(bool Pr)
{
bool Ret=false;
Zast->MClient->WriteDiaryEvent("Hazards","������ �������� �������� (��������)","");
try
{
if(Zast->MClient->PrepareLoadMetod("HReference"))
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
Comm->CommandText="Delete * From TempMet";
Comm->Execute();

MP<TADODataSet>TempMet(this);
TempMet->Connection=Zast->ADOConn;
TempMet->CommandText="Select TempMet.[�����], TempMet.[��������] From TempMet order by �����";
TempMet->Active=true;

Table* STMet=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("HReference")].ServerDB);

STMet->SetCommandText("Select TempMet.[�����], TempMet.[��������] From TempMet order by �����");
STMet->Active(true);
if(Pr)
{
FProg->Label1->Caption="������ ��������";
Zast->MClient->LoadTable(STMet, TempMet, FProg->Label1, FProg->Progress);
}
else
{
Zast->MClient->LoadTable(STMet, TempMet);
}

if(Zast->MClient->VerifyTable(STMet, TempMet)==0)
{
Comm->CommandText="Delete * from  ��������";
Comm->Execute();

MP<TADODataSet>Metod(this);
Metod->Connection=Zast->ADOConn;
Metod->CommandText="Select * from �������� ";
Metod->Active=true;

MP<TMemo>M(this);
M->Visible=false;
M->Parent=this;

TStrings* TT=M->Lines;
TT->Clear();
for(TempMet->First();!TempMet->Eof;TempMet->Next())
{
String S=TempMet->FieldByName("��������")->AsString;
TT->Append(S);
}
Metod->Insert();
Metod->FieldByName("�����")->Value=1;
Metod->FieldByName("��������")->Assign(TT);
Metod->Post();

Zast->MClient->WriteDiaryEvent("Hazards","����� �������� �������� (��������)","");
Ret=true;
}
Zast->MClient->DeleteTable(this, STMet);
}
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","������ �������� �������� (��������)"," ������ "+IntToStr(GetLastError()));

}
return Ret;
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::N32Click(TObject *Sender)
{
//������ ����� �����������
//������ � Reference
Zast->MClient->Start();
ShowMessage(WriteVozd(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String TDocuments::WriteVozd(bool Pr)
{
if(Pr)
{
FProg->Label1->Caption="������ �����������";
}
if(UploadNodeBranch("����_3","�����_3", Pr))
{

Zast->MClient->MergeNodeBranch("HReference", "����_3","�����_3","���������","�����������","�����������","����� �����������","������������ �����������");
//                       ���� ������������  | ����  | ����� |���� �������� |������� | ���� �������� | ���� ������� ������� | ������������ � ������� �������

//���������� � Aspects
return "���������";
}
else
{
return "������";
}
}
//---------------------------------------------------
bool TDocuments::UploadNodeBranch(String NodeTable, String BranchTable, bool Pr)
{
bool Ret=false;
Zast->MClient->WriteDiaryEvent("Hazards","������ ���������� ����� � ������ ������������ (��������)","����: "+NodeTable+" �����: "+BranchTable);

try
{
Zast->MClient->SetCommandText("HReference","Delete * From TempNode");
Zast->MClient->CommandExec("HReference");

MP<TADODataSet>Node(this);
Node->Connection=Zast->ADOConn;
Node->CommandText="Select "+NodeTable+".[����� ����], "+NodeTable+".[��������], "+NodeTable+".[��������] From "+NodeTable+" Order by [��������], [����� ����]";
Node->Active=true;

Table* TempNode=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("HReference")].ServerDB);

TempNode->SetCommandText("Select TempNode.[����� ����], TempNode.[��������], TempNode.[��������] From TempNode Order by [��������], [����� ����]");
TempNode->Active(true);

Zast->MClient->LoadTable(Node, TempNode);

if(Zast->MClient->VerifyTable(Node, TempNode)==0)
{

Zast->MClient->SetCommandText("HReference","Delete * From TempBranch");
Zast->MClient->CommandExec("HReference");

MP<TADODataSet>Branch(this);
Branch->Connection=Zast->ADOConn;
Branch->CommandText="Select "+BranchTable+".[����� �����], "+BranchTable+".[����� ��������], "+BranchTable+".[��������], "+BranchTable+".[�����] From "+BranchTable+" Order by [����� ��������], [����� �����]";
Branch->Active=true;

Table* TempBranch=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("HReference")].ServerDB);

TempBranch->SetCommandText("Select TempBranch.[����� �����], TempBranch.[����� ��������], TempBranch.[��������], TempBranch.[�����] From TempBranch Order by [����� ��������], [����� �����]");
TempBranch->Active(true);

if(Pr)
{
Zast->MClient->LoadTable(Branch, TempBranch, FProg->Label1, FProg->Progress);
}
else
{
Zast->MClient->LoadTable(Branch, TempBranch);
}
if(Zast->MClient->VerifyTable(Branch, TempBranch)==0)
{

Zast->MClient->WriteDiaryEvent("Hazards","����� ���������� ����� � ������ ������������ (��������)","����: "+NodeTable+" �����: "+BranchTable);

Ret=true;
}

Zast->MClient->DeleteTable(this, TempBranch);

if(!Ret)
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","���� ���������� ����� � ������ ������������ (��������)","����: "+NodeTable+" �����: "+BranchTable);

}
}
Zast->MClient->DeleteTable(this, TempNode);
}
catch(...)
{
 Zast->MClient->WriteDiaryEvent("������ NetAspects ","������ ���������� ����� � ������ ������������ (��������)","����: "+NodeTable+" �����: "+BranchTable+" ������ "+IntToStr(GetLastError()));

}
return Ret;
}
//---------------------------------------------------------
void __fastcall TDocuments::N33Click(TObject *Sender)
{
//������ ����� �����������
//������ � Reference
Zast->MClient->Start();
ShowMessage(WriteMeropr(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String TDocuments::WriteMeropr(bool Pr)
{
if(Pr)
{
FProg->Label1->Caption="������ �����������";
}
if(UploadNodeBranch("����_4","�����_4", Pr))
{
Zast->MClient->MergeNodeBranch("HReference", "����_4","�����_4");
//                       ���� ������������  | ����  | �����

//���������� � Aspects
return "���������";
}
else
{
return "������";
}
}
//--------------------------------------------
void __fastcall TDocuments::N34Click(TObject *Sender)
{
//������ ����� ����������
//������ � Reference
Zast->MClient->Start();
ShowMessage(WriteTerr(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String TDocuments::WriteTerr(bool Pr)
{
if(Pr)
{
FProg->Label1->Caption="������ ����������";
}
if(UploadNodeBranch("����_5","�����_5", Pr))
{
Zast->MClient->MergeNodeBranch("HReference", "����_5","�����_5","���������","����������","��� ����������","����� ����������","������������ ����������");
//                       ���� ������������  | ����  | ����� |���� �������� |������� | ���� �������� | ���� ������� ������� | ������������ � ������� �������

//���������� � Aspects
return "���������";
}
else
{
return "������";
}
}
//------------------------------------------------------------
void __fastcall TDocuments::N35Click(TObject *Sender)
{
//������ ����� ������������
//������ � Reference
Zast->MClient->Start();
ShowMessage(WriteDeyat(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String TDocuments::WriteDeyat(bool Pr)
{
if(Pr)
{
FProg->Label1->Caption="������ ����� ������������";
}
if(UploadNodeBranch("����_6","�����_6", Pr))
{
Zast->MClient->MergeNodeBranch("HReference", "����_6","�����_6","���������","������������","������������","����� ������������","������������ ������������");
//                       ���� ������������  | ����  | ����� |���� �������� |������� | ���� �������� | ���� ������� ������� | ������������ � ������� �������

//���������� � Aspects
return "���������";
}
else
{
return "������";
}
}
//----------------------------------------------------
void __fastcall TDocuments::N36Click(TObject *Sender)
{
//������ ������������� ��������
//������ � Reference
Zast->MClient->Start();
ShowMessage(WriteAspect(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------
String TDocuments::WriteAspect(bool Pr)
{
if(Pr)
{
FProg->Label1->Caption="������ ����� ������";
}
if(UploadNodeBranch("����_7","�����_7", Pr))
{
Zast->MClient->MergeNodeBranch("HReference", "����_7","�����_7","���������","������","������","����� �������","������������ �������");
//                       ���� ������������  | ����  | ����� |���� �������� |������� | ���� �������� | ���� ������� ������� | ������������ � ������� �������

//���������� � Aspects
return "���������";
}
else
{
return "������";
}
}
//-----------------------------------------------------
void __fastcall TDocuments::N37Click(TObject *Sender)
{
DataSetRefresh2->Execute();
DataSetPost2->Execute();
Zast->MClient->Start();
ShowMessage(WriteCryt(true));
Zast->MClient->Stop();

}
//---------------------------------------------------------------------------
String TDocuments::WriteCryt(bool Pr)
{
String Res="������";
Zast->MClient->WriteDiaryEvent("Hazards","������ ���������� ��������� (��������)","");

try
{
Zast->MClient->SetCommandText("HReference","Delete * From TempZn");
Zast->MClient->CommandExec("HReference");

MP<TADODataSet>Zn(this);
Zn->Connection=Zast->ADOConn;
Zn->CommandText="Select ����������.[����� ����������], ����������.[������������ ����������], ����������.[��������1], ����������.[��������], ����������.[��� �������], ����������.[���� �������], ����������.[����������� ����] From ���������� Order by [����� ����������]";
Zn->Active=true;

Table* TempZn=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("HReference")].ServerDB);

TempZn->SetCommandText("Select TempZn.[����� ����������], TempZn.[������������ ����������], TempZn.[��������1], TempZn.[��������], TempZn.[��� �������], TempZn.[���� �������], TempZn.[����������� ����] From TempZn Order by [����� ����������]");
TempZn->Active(true);
if(Pr)
{
FProg->Label1->Caption="������ ���������";
Zast->MClient->LoadTable(Zn, TempZn, FProg->Label1, FProg->Progress);
}
else
{
Zast->MClient->LoadTable(Zn, TempZn);
}

if(Zast->MClient->VerifyTable(Zn, TempZn)==0)
{
Zast->MClient->MergeZn("HReference","���������");

Zast->MClient->WriteDiaryEvent("Hazards","����� ���������� ��������� (��������)","");
Res="���������";
}

Zast->MClient->DeleteTable(this, TempZn);

if(Res!="���������")
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","���� ���������� ��������� (��������)","");
}
}
catch(...)
{
 Zast->MClient->WriteDiaryEvent("������ NetAspects ","������ ���������� ��������� (��������)","������ "+IntToStr(GetLastError()));

}
return Res;
}
//--------------------------------------------
void __fastcall TDocuments::N38Click(TObject *Sender)
{
//������ �������������
DataSetRefresh3->Execute();
DataSetPost3->Execute();
Zast->MClient->Start();
ShowMessage(WritePodr(true));
Zast->MClient->Stop();

}
//---------------------------------------------------------------------------
String TDocuments::WritePodr(bool Pr)
{
String Ret="������ ������";

Zast->MClient->WriteDiaryEvent("Hazards","������ ���������� ������������� (��������)","");

try
{
MP<TADODataSet>Podr(this);
Podr->Connection=Zast->ADOAspect;
Podr->CommandText="Select �������������.[����� �������������], �������������.[�������� �������������], �������������.[ServerNum] from ������������� Order by [����� �������������]";
Podr->Active=true;

Zast->MClient->SetCommandText("���������","Delete * From Temp�������������");
Zast->MClient->CommandExec("���������");

Table* TempPodr2=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

TempPodr2->SetCommandText("Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������], Temp�������������.[ServerNum] from Temp������������� Order by [����� �������������]");
TempPodr2->Active(true);

Zast->MClient->LoadTable(Podr, TempPodr2);

if(Zast->MClient->VerifyTable(Podr, TempPodr2)==0)
{
Zast->MClient->SavePodr("���������");

MP<TADODataSet>Temp(this);
Temp->Connection=Zast->ADOAspect;
Temp->CommandText="Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������] from Temp������������� Order by [����� �������������]";
Temp->Active=true;

Table* TempPodr=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

TempPodr->SetCommandText("Select Temp�������������.[����� �������������], Temp�������������.[�������� �������������] from Temp������������� Order by [����� �������������]");
TempPodr->Active(true);
if(Pr)
{
 FProg->Label1->Caption="������ �������������";
Zast->MClient->LoadTable(TempPodr, Temp, FProg->Label1, FProg->Progress);
}
else
{
Zast->MClient->LoadTable(TempPodr, Temp);
}

if(Zast->MClient->VerifyTable(Temp, TempPodr)==0)
{
Podr->Active=false;
Podr->CommandText="Select �������������.[����� �������������], �������������.[�������� �������������], �������������.[ServerNum] from �������������  Where ServerNum=0 Order by [����� �������������]";
Podr->Active=true;

for(Podr->First();!Podr->Eof;Podr->Next())
{
int N=Podr->FieldByName("����� �������������")->Value;
if(Temp->Locate("����� �������������", N, SO))
{
Podr->Edit();
Podr->FieldByName("ServerNum")->Value=Temp->FieldByName("����� �������������")->Value;
Podr->Post();
}
else
{
ShowMessage("������ ������ ������������� �����="+IntToStr(N));
}
}

Zast->MClient->SetCommandText("���������","Delete * From Temp�������������");
Zast->MClient->CommandExec("���������");

Zast->MClient->WriteDiaryEvent("Hazards","����� ���������� ������������� (��������)","");
Ret="���������";
}
Zast->MClient->DeleteTable(this, TempPodr);
}
Zast->MClient->DeleteTable(this, TempPodr2);

MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOAspect;
Comm->CommandText="Delete * from Temp�������������";
Comm->Execute();

if(Ret!="���������")
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","���� ���������� ������������� (��������)","");
}
}
catch(...)
{
 Zast->MClient->WriteDiaryEvent("������ NetAspects ","������ ���������� ��������� (��������)","������ "+IntToStr(GetLastError()));

}
return Ret;
}
//--------------------------------------------------
void __fastcall TDocuments::N39Click(TObject *Sender)
{
//������ ��������
DataSetRefresh1->Execute();
DataSetPost4->Execute();
Zast->MClient->Start();
ShowMessage(WriteSit(true));
Zast->MClient->Stop();

}
//---------------------------------------------------------------------------
String TDocuments::WriteSit(bool Pr)
{
String Ret="������ ������";
Zast->MClient->WriteDiaryEvent("Hazards","������ ���������� �������� (��������)","");

try
{
Zast->MClient->SetCommandText("HReference","Delete * From TempSit");
Zast->MClient->CommandExec("HReference");

MP<TADODataSet>Sit(this);
Sit->Connection=Zast->ADOConn;
Sit->CommandText="Select ��������.[����� ��������], ��������.[�������� ��������] from �������� Order by [����� ��������]";
Sit->Active=true;

Table* TempSit=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("HReference")].ServerDB);

TempSit->SetCommandText("Select TempSit.[����� ��������], TempSit.[�������� ��������] from TempSit Order by [����� ��������]");
TempSit->Active(true);

Zast->MClient->LoadTable(Sit, TempSit);

if(Zast->MClient->VerifyTable(Sit, TempSit)==0)
{
Zast->MClient->SetCommandText("���������","Delete * From TempSit");
Zast->MClient->CommandExec("���������");

Table* TempSit2=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("���������")].ServerDB);

TempSit2->SetCommandText("Select TempSit.[����� ��������], TempSit.[�������� ��������] from TempSit Order by [����� ��������]");
TempSit2->Active(true);
if(Pr)
{
FProg->Label1->Caption="������ ��������";
Zast->MClient->LoadTable(Sit, TempSit2, FProg->Label1, FProg->Progress);
}
else
{
Zast->MClient->LoadTable(Sit, TempSit2);
}

if(Zast->MClient->VerifyTable(Sit, TempSit2)==0)
{
Zast->MClient->SaveSit("HReference","���������");

Zast->MClient->WriteDiaryEvent("Hazards","����� ���������� �������� (��������)","");
Ret="���������";
}
Zast->MClient->DeleteTable(this, TempSit2);
}
Zast->MClient->DeleteTable(this, TempSit);

if(Ret!="���������")
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","���� ���������� �������� (��������)","");
}
}
catch(...)
{
 Zast->MClient->WriteDiaryEvent("������ NetAspects ","������ ���������� �������� (��������)","������ "+IntToStr(GetLastError()));

}
return Ret;
}
//------------------------------------------------
void __fastcall TDocuments::N18Click(TObject *Sender)
{
//������ ��������
DataSetRefresh4->Execute();
DataSetPost1->Execute();

Zast->MClient->Start();
ShowMessage(WriteMet(true));
Zast->MClient->Stop();


}
//---------------------------------------------------------------------------
String TDocuments::WriteMet(bool Pr)
{
String Ret="������ ������";

Zast->MClient->WriteDiaryEvent("Hazards","������ ���������� �������� (��������)","");

try
{
PrepareMet();

Zast->MClient->SetCommandText("HReference","Delete * From TempMet");
Zast->MClient->CommandExec("HReference");

MP<TADODataSet>Local(this);
Local->Connection=Zast->ADOConn;
Local->CommandText="Select TempMet.�����, TempMet.�������� from TempMet Order by �����";
Local->Active=true;

Table* Server=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("HReference")].ServerDB);

Server->SetCommandText("Select TempMet.�����, TempMet.�������� from TempMet Order by �����");
Server->Active(true);
if(Pr)
{
FProg->Label1->Caption="������ ��������";
Zast->MClient->LoadTable(Local, Server, FProg->Label1, FProg->Progress);
}
else
{
Zast->MClient->LoadTable(Local, Server);
}

if(Zast->MClient->VerifyTable(Local, Server)==0)
{
Zast->MClient->SaveMetod("HReference");

Zast->MClient->WriteDiaryEvent("Hazards","����� ���������� �������� (��������)","");
Ret="���������";
}

Zast->MClient->DeleteTable(this, Server);

if(Ret!="���������")
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","���� ���������� �������� (��������)","");
}
}
catch(...)
{
 Zast->MClient->WriteDiaryEvent("������ NetAspects ","������ ���������� �������� (��������)","������ "+IntToStr(GetLastError()));

}
return Ret;
}
//-------------------------------------------------
void TDocuments::PrepareMet()
{
Zast->MClient->WriteDiaryEvent("Hazards","������ ���������� � ���������� �������� (��������)","");

try
{
MP<TADOCommand>Comm(this);
Comm->Connection=Zast->ADOConn;
Comm->CommandText="Delete * from TempMet";
Comm->Execute();

MP<TADODataSet>Tab(this);
Tab->Connection=Zast->ADOConn;
Tab->CommandText="Select * from TempMet Order by �����";
Tab->Active=true;

MP<TADODataSet>Tab1(this);
Tab1->Connection=Zast->ADOConn;
Tab1->CommandText="Select * from �������� Order by �����";
Tab1->Active=true;
Tab1->First();

MP<TDataSource>DS(this);
DS->DataSet=Tab1;
DS->Enabled=true;

MP<TDBMemo>TDBM(this);
TDBM->Parent=this;
TDBM->Visible=false;
TDBM->Width=1009;
TDBM->DataSource=DS;
TDBM->DataField="��������";


TStrings* TS=TDBM->Lines;
int N=TS->Count;
for(int i=0;i<N;i++)
{
String S=TS->Strings[i];
Tab->Insert();
Tab->FieldByName("�����")->Value=i;
Tab->FieldByName("��������")->Value=S;
Tab->Post();
}

Zast->MClient->WriteDiaryEvent("Hazards","����� ���������� � ���������� �������� (��������)","");
}
catch(...)
{
Zast->MClient->WriteDiaryEvent("������ NetAspects","������ ���������� � ���������� �������� (��������)","");
}

}
//------------------------------------------------------------------




void __fastcall TDocuments::N31Click(TObject *Sender)
{
FProg->Visible=true;
FProg->Progress->Position=0;
FProg->Progress->Min=0;
FProg->Progress->Max=9;
FProg->Progress->Position++;

Zast->MClient->Start();

FProg->Progress->Position++;
FProg->Label1->Caption="������ �������� ������...";
FProg->Label1->Repaint();
ReadMet(false);
FProg->Progress->Position++;
FProg->Label1->Caption="������ ������ �������������...";
FProg->Label1->Repaint();
ReadPodr(false);
FProg->Progress->Position++;
FProg->Label1->Caption="������ ����������� ��������� ������...";
FProg->Label1->Repaint();
ReadCrit(false);

FProg->Progress->Position++;
FProg->Label1->Caption="������ ������ �����������...";
FProg->Label1->Repaint();
ReadVidVozd(false);
/*
FProg->Progress->Position++;
FProg->Label1->Caption="������ ������ �����������...";
FProg->Label1->Repaint();
ReadMeropr(false);
*/
FProg->Progress->Position++;
FProg->Label1->Caption="������ ������ ��������/���������...";
FProg->Label1->Repaint();
ReadTerr(false);
FProg->Progress->Position++;
FProg->Label1->Caption="������ ������ �������� ������...";
FProg->Label1->Repaint();
ReadDeyat(false);
FProg->Progress->Position++;
FProg->Label1->Caption="������ ����������� ����������... ";
FProg->Label1->Repaint();
ReadAspect(false);

FProg->Progress->Position++;
FProg->Label1->Caption="������ ������ ��������...";
FProg->Label1->Repaint();
ReadSit(false);

FProg->Visible=false;
Zast->MClient->Stop();
ShowMessage("���������");
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N41Click(TObject *Sender)
{
SaveAllTables(false);

ShowMessage("���������");
}
//---------------------------------------------------------------------------
void TDocuments::SaveAllTables(bool Full)
{
FProg->Visible=true;
FProg->Progress->Position=0;
Zast->MClient->Start();
if(Full)
{

FProg->Progress->Max=9;
}
else
{
FProg->Progress->Max=8;
}

FProg->Progress->Position++;
FProg->Label1->Caption="������ �������� ������...";
FProg->Label1->Repaint();
WriteMet(false);
FProg->Progress->Position++;
FProg->Label1->Caption="������ ������ �������������...";
FProg->Label1->Repaint();
WritePodr(false);
FProg->Progress->Position++;
FProg->Label1->Caption="������ ����������� ��������� ������...";
FProg->Label1->Repaint();
WriteCryt(false);
FProg->Progress->Position++;
FProg->Label1->Caption="������ ������ �����������...";
FProg->Label1->Repaint();
WriteVozd(false);
/*
FProg->Progress->Position++;
FProg->Label1->Caption="������ ������ �����������...";
FProg->Label1->Repaint();
WriteMeropr(false);
*/
FProg->Progress->Position++;
FProg->Label1->Caption="������ ������ ��������/���������...";
FProg->Label1->Repaint();
WriteTerr(false);
FProg->Progress->Position++;
FProg->Label1->Caption="������ ������ �������� ������...";
FProg->Label1->Repaint();
WriteDeyat(false);
FProg->Progress->Position++;
FProg->Label1->Caption="������ ����������� ����������... ";
FProg->Label1->Repaint();
WriteAspect(false);


FProg->Progress->Position++;
FProg->Label1->Caption="������ ������ ��������...";
FProg->Label1->Repaint();
WriteSit(false);

if(Full)
{
FProg->Progress->Position++;
FProg->Label1->Caption="������ ������...";
Zast->MClient->RegForm(MAsp);
MAsp->SaveAspects();

}


FProg->Visible=false;
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------


void __fastcall TDocuments::Button1Click(TObject *Sender)
{
//������ ��������
DataSetRefresh4->Execute();
DataSetPost1->Execute();

Zast->MClient->Start();
ShowMessage(WriteMet(true));
Zast->MClient->Stop();        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::Button2Click(TObject *Sender)
{
//������ �������������
DataSetRefresh3->Execute();
DataSetPost3->Execute();
Zast->MClient->Start();
ShowMessage(WritePodr(true));
Zast->MClient->Stop();        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::Button3Click(TObject *Sender)
{
DataSetRefresh2->Execute();
DataSetPost2->Execute();
Zast->MClient->Start();
ShowMessage(WriteCryt(true));
Zast->MClient->Stop();        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::Button4Click(TObject *Sender)
{
//������ ����� �����������
//������ � Reference
Zast->MClient->Start();
ShowMessage(WriteVozd(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::Button5Click(TObject *Sender)
{
//������ ����� �����������
//������ � Reference
Zast->MClient->Start();
ShowMessage(WriteMeropr(true));
Zast->MClient->Stop();        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::Button6Click(TObject *Sender)
{
//������ ����� ����������
//������ � Reference
Zast->MClient->Start();
ShowMessage(WriteTerr(true));
Zast->MClient->Stop();        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::Button7Click(TObject *Sender)
{
//������ ����� ������������
//������ � Reference
Zast->MClient->Start();
ShowMessage(WriteDeyat(true));
Zast->MClient->Stop();        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::Button8Click(TObject *Sender)
{
//������ ������������� ��������
//������ � Reference
Zast->MClient->Start();
ShowMessage(WriteAspect(true));
Zast->MClient->Stop();        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::Button9Click(TObject *Sender)
{
//������ ��������
DataSetRefresh1->Execute();
DataSetPost4->Execute();
Zast->MClient->Start();
ShowMessage(WriteSit(true));
Zast->MClient->Stop();
}
//---------------------------------------------------------------------------


//---------------------------------------------------------------------------

void __fastcall TDocuments::N19Click(TObject *Sender)
{
MAsp->ShowModal();
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::WMSysCommand(TMessage & Msg)
{
  switch (Msg.WParam)
  {
case SC_MINIMIZE:
{
Application->Minimize();
break;
}

default:
DefWindowProc(Handle,Msg.Msg,Msg.WParam,Msg.LParam);

 }
}







//-------------------------------------------------------------------------------
void __fastcall TDocuments::N00111Click(TObject *Sender)
{
//�001.1
Report1->Role=2;
Report1->Flt="";
Report1->FltName="��������";
Report1->PodrComText="select * From ������������� Order by [�������� �������������]";

Report1->NumRep=1;
Report1->RepBase=Zast->ADOAspect;
Report1->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N00121Click(TObject *Sender)
{
//�001.2
Report1->Role=2;
Report1->Flt="";
Report1->FltName="��������";
Report1->PodrComText="select * From ������������� Order by [�������� �������������]";

Report1->NumRep=2;
Report1->RepBase=Zast->ADOAspect;
Report1->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N22Click(TObject *Sender)
{
//������� �����
FSvod->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N23Click(TObject *Sender)
{
FAbout->ShowModal();
}
//---------------------------------------------------------------------------





void __fastcall TDocuments::FormKeyUp(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
if(Key==112)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"HAZARDS.HLP";
switch (PageControl1->TabIndex)
{
 case 0:
 {
  Application->HelpJump("IDH_��������������_��������");
  break;
 }
 case 1:
 {
  Application->HelpJump("IDH_��������������_�������������");
  break;
 }
 case 2:
 {
  Application->HelpJump("IDH_��������������_���������");
  break;
 }
 case 3:
 {
  Application->HelpJump("IDH_��������������_��������");
  break;
 }
 case 4: case 5: case 6: case 7: case 8:
 {
  Application->HelpJump("IDH_��������������_������������");
  break;
 }
}
}
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N1Click(TObject *Sender)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"HAZARDS.HLP";
  Application->HelpJump("IDH_�����_��������");
}
//---------------------------------------------------------------------------


