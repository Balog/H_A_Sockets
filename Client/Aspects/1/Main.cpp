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
int NumP=Podr->FieldByName("ServerNum")->AsInteger;

Zast->MClient->Start();
Table* STab=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("�������")].ServerDB);

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
 S1="�� ������� ������������ "+IntToStr(N)+" ������, ������������� ����� �������������.";
 break;
 }
 case 2: case 3: case 4:
 {
 S1="�� ������� ������������� "+IntToStr(N)+" �������, ������������� ����� �������������.";
 break;
 }
 default:
 {
 S1="�� ������� ������������� "+IntToStr(N)+" ��������, ������������� ����� �������������.";
 }
}

String S=S1+"\r��������� \"�������� ��������\" �������������� ���������� ��������� ������������� �� ��������";
Application->MessageBoxA(S.c_str(),"�������� �������������",MB_ICONEXCLAMATION);
}
Zast->MClient->DeleteTable(this, STab);
Zast->MClient->Stop();        
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

