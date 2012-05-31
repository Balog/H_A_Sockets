//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Main.h"
#include "inifiles.hpp";
#include "CodeText.h"
#include "MasterPointer.h"
#include "PassForm.h"
#include "Zastavka.h"
//#include "EditLogin.h"
#include "Progress.h"
#include "About.h"
#include "Rep1.h"
#include "Svod.h"
using namespace std;
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
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::FormClose(TObject *Sender, TCloseAction &Action)
{

switch (Application->MessageBoxA("�������� ��� ��������� �� ������?","����� �� ���������",MB_YESNOCANCEL	+MB_ICONQUESTION+MB_DEFBUTTON1))
{
 case IDYES:
 {
Zast->Stop=true;
Zast->MClient->Act.WaitCommand=0;

Zast->Saved=false;
//Zast->PrepareSaveLogins->Execute();


 Action=caNone;
 break;
 }
 case IDNO:
 {
Zast->MClient->WriteDiaryEvent("NetAspects","���������� ������ ����� �� ������","");
Zast->Stop=true;
Sleep(2000);
Zast->Close();
 Action=caNone;
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
void __fastcall TDocuments::FormShow(TObject *Sender)
{
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
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::N9Click(TObject *Sender)
{
//this->Close();
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
//-------------------------------------------------------------------
void TDocuments::Initialize()
{
Podr->Active=true;
Sit->Active=true;
}
//-------------------------------------------------------------------
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
Branches->FieldByName("��������")->Value="����� ������";
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
Branches->FieldByName("��������")->Value="����� ������";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("����� �����")->Value;


MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;
}
TreeView5->Items->AddChildObject(SelNode,"����� ������",MyNodePtr);        
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
void __fastcall TDocuments::MenuItem3Click(TObject *Sender)
{
Podr->Insert();
Podr->FieldByName("�������� �������������")->Value="����� �������������";
Podr->Post();        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::MenuItem4Click(TObject *Sender)
{
// Socket->Socket->SendText("Command:5;2|"+IntToStr(NameDB.Length())+"#"+NameDB+"|"+ServerSQL.Length()+"#"+ServerSQL+"|");
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("DeletePodr");

String DBName="�������";
String ServerSQL="SELECT �������.[����� �������], �������.������������� FROM �������";
String ClientSQL="SELECT [����� �������], ������������� FROM TempAspects";
Zast->MClient->ReadTable(DBName, ServerSQL, ClientSQL);

}
//---------------------------------------------------------------------------

void __fastcall TDocuments::MenuItem1Click(TObject *Sender)
{
// ��������
Ins=true;
/*
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
*/
ADODataSet1->Last();
int Max=ADODataSet1->FieldByName("���� �������")->Value;
int Min=ADODataSet1->FieldByName("��� �������")->Value;
bool Cr=ADODataSet1->FieldByName("��������")->Value;
ADODataSet1->Edit();
ADODataSet1->FieldByName("���� �������")->Value=Max-1;
ADODataSet1->Post();

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
ADODataSet2->CommandText="Select * From ���������� Order By [��� �������] DESC";
ADODataSet2->Active=true;
/*
int Min=ADODataSet2->FieldByName("��� �������")->Value;

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
*/
ADODataSet2->First();
int Min=ADODataSet2->FieldByName("��� �������")->Value;
ADODataSet2->Edit();
ADODataSet2->FieldByName("���� �������")->Value=Min+1;
ADODataSet2->Post();
for(ADODataSet2->First();!ADODataSet2->Eof;ADODataSet2->Next())
{
int NewMin=ADODataSet2->FieldByName("��� �������")->Value;
if(Min!=NewMin)
{
ADODataSet2->Edit();
ADODataSet2->FieldByName("���� �������")->Value=Min-1;
ADODataSet2->Post();
Min=NewMin;
}
else
{
 if(ADODataSet2->FieldByName("���� �������")->Value==ADODataSet2->FieldByName("��� �������")->Value)
 {
  ADODataSet2->Edit();
  ADODataSet2->FieldByName("���� �������")->Value=ADODataSet2->FieldByName("��� �������")->Value+1;
  ADODataSet2->Post();
 }
}

}

ADODataSet1->Active=false;
ADODataSet1->Active=true;
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

void __fastcall TDocuments::ReadMetodClick(TObject *Sender)
{
DBMemo1->ReadOnly=!ReadMetod->Checked;        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::DBGrid1Exit(TObject *Sender)
{
DataSetRefresh2->Execute();
DataSetPost2->Execute();        
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

void __fastcall TDocuments::ADODataSet1AfterPost(TDataSet *DataSet)
{
if(C==false)
{

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



int Min=DataSet->FieldByName("��� �������")->Value;
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

void __fastcall TDocuments::N8Click(TObject *Sender)
{
this->Close();        
}
//---------------------------------------------------------------------------


void __fastcall TDocuments::N1Click(TObject *Sender)
{
Application->HelpFile=ExtractFilePath(Application->ExeName)+"NetAspects.HLP";
  Application->HelpJump("IDH_�����_��������");        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N2Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="ReadMetodika";
S.Text="������ ��������...";
S.Num=1;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N27Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="ReadPodrazd";
S.Text="������ �������������...";
S.Num=2;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N26Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="ReadCrit";
S.Text="������ ���������...";
S.Num=3;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N28Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="ReadSit";
S.Text="������ ��������...";
S.Num=4;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N6Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="ReadVozd1";
S.Text="������ ������ �����������...";
S.Num=5;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------
void TDocuments::MergeNode(String NodeTable)
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
}
//--------------------------------------------------------------------------
void TDocuments::MergeBranch(String BranchTable)
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
}
//--------------------------------------------------------------------------

void __fastcall TDocuments::N7Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="ReadMeropr1";
S.Text="������ ������ �����������...";
S.Num=6;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N20Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="ReadTerr1";
S.Text="������ ������ ����������...";
S.Num=7;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N24Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="ReadDeyat1";
S.Text="������ ������ ����� ������������...";
S.Num=8;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N25Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="ReadAspect1";
S.Text="������ ������ ������������� ��������...";
S.Num=8;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N31Click(TObject *Sender)
{
Prog->Show();
Prog->PB->Min=0;
Prog->PB->Position=0;
Prog->PB->Max=9;

ReadWrite.clear();
Str_RW S;
S.NameAction="ReadMetodika";
S.Text="������ ��������...";
S.Num=1;
ReadWrite.push_back(S);

S.NameAction="ReadPodrazd";
S.Text="������ �������������...";
S.Num=2;
ReadWrite.push_back(S);

S.NameAction="ReadCrit";
S.Text="������ ���������...";
S.Num=3;
ReadWrite.push_back(S);

S.NameAction="ReadSit";
S.Text="������ ��������...";
S.Num=4;
ReadWrite.push_back(S);

S.NameAction="ReadVozd1";
S.Text="������ ������ �����������...";
S.Num=5;
ReadWrite.push_back(S);

S.NameAction="ReadMeropr1";
S.Text="������ ������ �����������...";
S.Num=6;
ReadWrite.push_back(S);

S.NameAction="ReadTerr1";
S.Text="������ ������ ����������...";
S.Num=7;
ReadWrite.push_back(S);

S.NameAction="ReadDeyat1";
S.Text="������ ������ ����� ������������...";
S.Num=8;
ReadWrite.push_back(S);

S.NameAction="ReadAspect1";
S.Text="������ ������ ������������� ��������...";
S.Num=9;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::DBMemo1Exit(TObject *Sender)
{
DataSetRefresh4->Execute();
DataSetPost1->Execute();
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::DBGrid2Exit(TObject *Sender)
{
DataSetRefresh4->Execute();
DataSetPost1->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::DBGrid3Exit(TObject *Sender)
{
DataSetRefresh1->Execute();
DataSetPost4->Execute();
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::N18Click(TObject *Sender)
{
DataSetRefresh4->Execute();
DataSetPost1->Execute();

ReadWrite.clear();
Str_RW S;
S.NameAction="WriteMetodika";
S.Text="������ ��������...";
S.Num=1;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::N38Click(TObject *Sender)
{
DataSetRefresh3->Execute();
DataSetPost3->Execute();

ReadWrite.clear();
Str_RW S;
S.NameAction="WritePodr";
S.Text="������ �������������...";
S.Num=2;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::N37Click(TObject *Sender)
{
DataSetRefresh2->Execute();
DataSetPost2->Execute();

ReadWrite.clear();
Str_RW S;
S.NameAction="WriteCrit";
S.Text="������ ���������...";
S.Num=3;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N39Click(TObject *Sender)
{
DataSetRefresh1->Execute();
DataSetPost4->Execute();

ReadWrite.clear();
Str_RW S;
S.NameAction="WriteSit";
S.Text="������ ��������...";
S.Num=4;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N32Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="WriteVozd1";
S.Text="������ �����������...";
S.Num=5;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N33Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="WriteMeropr1";
S.Text="������ �����������...";
S.Num=6;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N34Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="WriteTerr1";
S.Text="������ ����������...";
S.Num=7;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N35Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="WriteDeyat1";
S.Text="������ ����� ������������...";
S.Num=8;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N36Click(TObject *Sender)
{
ReadWrite.clear();
Str_RW S;
S.NameAction="WriteAspect1";
S.Text="������ ������ ������������� ��������...";
S.Num=9;
ReadWrite.push_back(S);
Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N41Click(TObject *Sender)
{
Prog->Show();
Prog->PB->Min=0;
Prog->PB->Position=0;
Prog->PB->Max=9;

ReadWrite.clear();
Str_RW S;
S.NameAction="WriteMetodika";
S.Text="������ ��������...";
S.Num=1;
ReadWrite.push_back(S);

S.NameAction="WritePodr";
S.Text="������ �������������...";
S.Num=2;
ReadWrite.push_back(S);


S.NameAction="WriteCrit";
S.Text="������ ���������...";
S.Num=3;
ReadWrite.push_back(S);


S.NameAction="WriteSit";
S.Text="������ ��������...";
S.Num=4;
ReadWrite.push_back(S);


S.NameAction="WriteVozd1";
S.Text="������ �����������...";
S.Num=5;
ReadWrite.push_back(S);


S.NameAction="WriteMeropr1";
S.Text="������ �����������...";
S.Num=6;
ReadWrite.push_back(S);


S.NameAction="WriteTerr1";
S.Text="������ ����������...";
S.Num=7;
ReadWrite.push_back(S);


S.NameAction="WriteDeyat1";
S.Text="������ ����� ������������...";
S.Num=8;
ReadWrite.push_back(S);


S.NameAction="WriteAspect1";
S.Text="������ ������ ������������� ��������...";
S.Num=9;
ReadWrite.push_back(S);

Zast->ReadWriteDoc->Execute();
}
//---------------------------------------------------------------------------

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

