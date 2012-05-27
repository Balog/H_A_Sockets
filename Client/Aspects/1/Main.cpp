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


Metod->CommandText="Select * From Методика";

Podr->Active=true;
Sit->Active=true;
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::FormClose(TObject *Sender, TCloseAction &Action)
{

switch (Application->MessageBoxA("Записать все изменения на сервер?","Выход из программы",MB_YESNOCANCEL	+MB_ICONQUESTION+MB_DEFBUTTON1))
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
Zast->MClient->WriteDiaryEvent("NetAspects","Завершение работы отказ от записи","");
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
ADODataSet1->CommandText="select * from Значимость Order By [Мин граница]";
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
//Добавить проверку на отсутствие у родителя удаляемой ветви (узла) потомков
//и перевода его в категорию ветвей
TTreeNode *N;


int Number=PMyNode(TreeView1->Selected->Data)->Number;
int Parent=PMyNode(TreeView1->Selected->Data)->Parent;
bool IsNode=PMyNode(TreeView1->Selected->Data)->Node;
if (IsNode==true)
{


DeleteNode("Узлы_3","Ветви_5", Number);

}
else
{
Comm->CommandText="Delete * From Ветви_3 Where [Номер ветви]="+IntToStr(Number);
Comm->Execute();

}

Branches->Active=false;
Branches->CommandText="Select * From Ветви_3 Where [Номер родителя]="+IntToStr(Parent);
Branches->Active=true;
Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_3 Where Родитель="+IntToStr(Parent);
Nodes->Active=true;

if (Branches->RecordCount==0 & Nodes->RecordCount==0)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From Узлы_3 Where [Номер узла]="+IntToStr(Parent);
 Nodes->Active=true;
 if (Nodes->RecordCount==0)
 {
  ShowMessage("Ошибка! Удалялась запись в таблице узлов номер "+IntToStr(Number)+" предок номер "+IntToStr(Parent));
  ShowMessage("Не найдена запись в таблице узлов номер "+IntToStr(Parent));

 }
 else
 {
 int NumParent=Nodes->FieldByName("Родитель")->Value;
 if (NumParent!=0)
 {
 AnsiString Text=Nodes->FieldByName("Название")->Value;;
 Nodes->Delete();
 Branches->Active=false;
 Branches->CommandText="Select * From Ветви_3";
 Branches->Active=true;
 Branches->Append();
 Branches->FieldByName("Номер родителя")->Value=NumParent;
 Branches->FieldByName("Название")->Value=Text;
 Branches->Post();
 Branches->Last();
 int Number=Branches->FieldByName("Номер ветви")->Value;
 N=TreeView1->GetNodeAt(XX,YY);
 TTreeNode * ParNode=N->Parent;
 PMyNode(ParNode->Data)->Number=Number;
 PMyNode(ParNode->Data)->Parent=NumParent;
 PMyNode(ParNode->Data)->Node=false;
 //
 //ParNode->Text=Text+" "+IntToStr(Number)+" "+IntToStr(NumParent)+" Ветвь";
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
Nodes->CommandText="Select * From Узлы_3";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
 int Parent=Nodes->FieldByName("Родитель")->Value;
 if (Parent!=0)
 {
 Branches->Active=false;
 Branches->CommandText="Select * From Узлы_3 Where [Номер узла]="+IntToStr(Parent);
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
NumParent=Nodes->FieldByName("Номер узла")->Value;
AnsiString Text=Nodes->FieldByName("Название")->Value;

MyNodePtr->Number=NumParent;
MyNodePtr->Parent=0;
MyNodePtr->Node=true;

Nodes->Edit();
N1=TreeView1->Items->AddChildObject(NULL,Text,MyNodePtr);

Nodes->Post();

Branches->Active=false;
Branches->CommandText="Select * From Ветви_3 Where [Номер родителя]="+IntToStr(NumParent);
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
Nodes->CommandText="Select * From Узлы_3 Where ([Родитель])>0";
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


Nodes->Edit();

Nodes->Post();
Node NN;
NN.Number=NumNode;
NN.Parent=NumParent;
NN.Nod=N;
It_Node_3=NodeVector_3.end();
NodeVector_3.insert(It_Node_3,NN);


Branches->Active=false;
Branches->CommandText="Select * From Ветви_3 Where [Номер родителя]="+IntToStr(NumNode);
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
Nodes->CommandText="Select * From Узлы_4";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
 int Parent=Nodes->FieldByName("Родитель")->Value;
 if (Parent!=0)
 {
 Branches->Active=false;
 Branches->CommandText="Select * From Узлы_4 Where [Номер узла]="+IntToStr(Parent);
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
NumParent=Nodes->FieldByName("Номер узла")->Value;
AnsiString Text=Nodes->FieldByName("Название")->Value;

MyNodePtr->Number=NumParent;
MyNodePtr->Parent=0;
MyNodePtr->Node=true;

Nodes->Edit();
N1=TreeView2->Items->AddChildObject(NULL,Text,MyNodePtr);

Nodes->Post();

Branches->Active=false;
Branches->CommandText="Select * From Ветви_4 Where [Номер родителя]="+IntToStr(NumParent);
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
TreeView2->Items->AddChildObject(N1,TextBranches,MyNodePtr);
}

Branches->Edit();

Branches->Post();
Branches->Next();
}



Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_4 Where ([Родитель])>0";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
int NumNode=Nodes->FieldByName("Номер узла")->Value;
NumParent=Nodes->FieldByName("Родитель")->Value;
AnsiString TextNode=Nodes->FieldByName("Название")->Value;
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
Branches->CommandText="Select * From Ветви_4 Where [Номер родителя]="+IntToStr(NumNode);
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
Nodes->CommandText="Select * From Узлы_5";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
 int Parent=Nodes->FieldByName("Родитель")->Value;
 if (Parent!=0)
 {
 Branches->Active=false;
 Branches->CommandText="Select * From Узлы_5 Where [Номер узла]="+IntToStr(Parent);
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
NumParent=Nodes->FieldByName("Номер узла")->Value;
AnsiString Text=Nodes->FieldByName("Название")->Value;

MyNodePtr->Number=NumParent;
MyNodePtr->Parent=0;
MyNodePtr->Node=true;

Nodes->Edit();
N1=TreeView3->Items->AddChildObject(NULL,Text,MyNodePtr);

Nodes->Post();

Branches->Active=false;
Branches->CommandText="Select * From Ветви_5 Where [Номер родителя]="+IntToStr(NumParent);
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
TreeView3->Items->AddChildObject(N1,TextBranches,MyNodePtr);
}

Branches->Edit();

Branches->Post();
Branches->Next();
}



Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_5 Where ([Родитель])>0";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
int NumNode=Nodes->FieldByName("Номер узла")->Value;
NumParent=Nodes->FieldByName("Родитель")->Value;
AnsiString TextNode=Nodes->FieldByName("Название")->Value;
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
//N=TreeView1->Items->AddChildObject(ParNod,TextNode+" "+IntToStr(MyNodePtr->Number)+" "+IntToStr(MyNodePtr->Parent)+" Узел",MyNodePtr);
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
Branches->CommandText="Select * From Ветви_5 Where [Номер родителя]="+IntToStr(NumNode);
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
Nodes->CommandText="Select * From Узлы_6";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
 int Parent=Nodes->FieldByName("Родитель")->Value;
 if (Parent!=0)
 {
 Branches->Active=false;
 Branches->CommandText="Select * From Узлы_6 Where [Номер узла]="+IntToStr(Parent);
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
NumParent=Nodes->FieldByName("Номер узла")->Value;
AnsiString Text=Nodes->FieldByName("Название")->Value;

MyNodePtr->Number=NumParent;
MyNodePtr->Parent=0;
MyNodePtr->Node=true;

Nodes->Edit();
N1=TreeView4->Items->AddChildObject(NULL,Text,MyNodePtr);

Nodes->Post();

Branches->Active=false;
Branches->CommandText="Select * From Ветви_6 Where [Номер родителя]="+IntToStr(NumParent);
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
TreeView4->Items->AddChildObject(N1,TextBranches,MyNodePtr);
}

Branches->Edit();

Branches->Post();
Branches->Next();
}



Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_6 Where ([Родитель])>0";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
int NumNode=Nodes->FieldByName("Номер узла")->Value;
NumParent=Nodes->FieldByName("Родитель")->Value;
AnsiString TextNode=Nodes->FieldByName("Название")->Value;
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
//N=TreeView1->Items->AddChildObject(ParNod,TextNode+" "+IntToStr(MyNodePtr->Number)+" "+IntToStr(MyNodePtr->Parent)+" Узел",MyNodePtr);
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
Branches->CommandText="Select * From Ветви_6 Where [Номер родителя]="+IntToStr(NumNode);
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
Nodes->CommandText="Select * From Узлы_7";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
 int Parent=Nodes->FieldByName("Родитель")->Value;
 if (Parent!=0)
 {
 Branches->Active=false;
 Branches->CommandText="Select * From Узлы_7 Where [Номер узла]="+IntToStr(Parent);
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
NumParent=Nodes->FieldByName("Номер узла")->Value;
AnsiString Text=Nodes->FieldByName("Название")->Value;

MyNodePtr->Number=NumParent;
MyNodePtr->Parent=0;
MyNodePtr->Node=true;

Nodes->Edit();
N1=TreeView5->Items->AddChildObject(NULL,Text,MyNodePtr);

Nodes->Post();

Branches->Active=false;
Branches->CommandText="Select * From Ветви_7 Where [Номер родителя]="+IntToStr(NumParent);
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
TreeView5->Items->AddChildObject(N1,TextBranches,MyNodePtr);
}

Branches->Edit();

Branches->Post();
Branches->Next();
}



Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_7 Where ([Родитель])>0";
Nodes->Active=true;
Nodes->First();
for(int i=0;i<Nodes->RecordCount;i++)
{
int NumNode=Nodes->FieldByName("Номер узла")->Value;
NumParent=Nodes->FieldByName("Родитель")->Value;
AnsiString TextNode=Nodes->FieldByName("Название")->Value;
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
Branches->CommandText="Select * From Ветви_7 Where [Номер родителя]="+IntToStr(NumNode);
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

//Номер предка
//Нужно узнать номер новой ветви
int NumParentNode=PMyNode(SelNode->Data)->Number;
//Является ли предок узлом?
IsNode=PMyNode(SelNode->Data)->Node;
if (IsNode==true)
{
//Если является то нужно только оформить на него ссылку
Branches->Active=false;
Branches->CommandText="Select * From Ветви_3";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("Номер родителя")->Value=NumParentNode;
Branches->FieldByName("Название")->Value="Новое воздействие";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("Номер ветви")->Value;

MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;

}
else
{
//А если является ветвью то нужно ее перенести в узлы и оформить на нее ссылку.
int Par;
AnsiString Text;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_3 Where [Номер ветви]="+IntToStr(NumParentNode);
Branches->Active=true;
Par=Branches->FieldByName("Номер родителя")->Value;
Text=Branches->FieldByName("Название")->Value;
Branches->Delete();
Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_3";
Nodes->Active=true;
Nodes->Append();
Nodes->FieldByName("Родитель")->Value=Par;
Nodes->FieldByName("Название")->Value=Text;
Nodes->Post();
Nodes->Last();

//PMyNode(TreeView1->Selected->Data)->Node
PMyNode(SelNode->Data)->Number=Nodes->FieldByName("Номер узла")->Value;
PMyNode(SelNode->Data)->Parent=Nodes->FieldByName("Родитель")->Value;
PMyNode(SelNode->Data)->Node=true;
NumParentNode=Nodes->FieldByName("Номер узла")->Value;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_3";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("Номер родителя")->Value=NumParentNode;
Branches->FieldByName("Название")->Value="Новое воздействие";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("Номер ветви")->Value;


MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;
}
TreeView1->Items->AddChildObject(SelNode,"Новое воздействие",MyNodePtr);        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N10Click(TObject *Sender)
{
int Number;
bool IsNode;
TTreeNode *N;
//PMyNode(TreeView1->Selected->Data)->Number;

//Номер предка
//Нужно узнать номер новой ветви
int NumParentNode=PMyNode(SelNode->Data)->Number;
//Является ли предок узлом?
IsNode=PMyNode(SelNode->Data)->Node;
if (IsNode==true)
{
//Если является то нужно только оформить на него ссылку
Branches->Active=false;
Branches->CommandText="Select * From Ветви_4";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("Номер родителя")->Value=NumParentNode;
Branches->FieldByName("Название")->Value="Новое мероприятие";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("Номер ветви")->Value;

MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;

}
else
{
//А если является ветвью то нужно ее перенести в узлы и оформить на нее ссылку.
int Par;
AnsiString Text;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_4 Where [Номер ветви]="+IntToStr(NumParentNode);
Branches->Active=true;
Par=Branches->FieldByName("Номер родителя")->Value;
Text=Branches->FieldByName("Название")->Value;
Branches->Delete();
Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_4";
Nodes->Active=true;
Nodes->Append();
Nodes->FieldByName("Родитель")->Value=Par;
Nodes->FieldByName("Название")->Value=Text;
Nodes->Post();
Nodes->Last();

//PMyNode(TreeView1->Selected->Data)->Node
PMyNode(SelNode->Data)->Number=Nodes->FieldByName("Номер узла")->Value;
PMyNode(SelNode->Data)->Parent=Nodes->FieldByName("Родитель")->Value;
PMyNode(SelNode->Data)->Node=true;
NumParentNode=Nodes->FieldByName("Номер узла")->Value;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_4";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("Номер родителя")->Value=NumParentNode;
Branches->FieldByName("Название")->Value="Новое мероприятие";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("Номер ветви")->Value;


MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;
}
TreeView2->Items->AddChildObject(SelNode,"Новое мероприятие",MyNodePtr);        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N11Click(TObject *Sender)
{
//Добавить проверку на отсутствие у родителя удаляемой ветви (узла) потомков
//и перевода его в категорию ветвей
TTreeNode *N;


int Number=PMyNode(TreeView2->Selected->Data)->Number;
int Parent=PMyNode(TreeView2->Selected->Data)->Parent;
bool IsNode=PMyNode(TreeView2->Selected->Data)->Node;
if (IsNode==true)
{
DeleteNode("Узлы_4","Ветви_4", Number);

}
else
{
Comm->CommandText="Delete * From Ветви_4 Where [Номер ветви]="+IntToStr(Number);
Comm->Execute();

}

Branches->Active=false;
Branches->CommandText="Select * From Ветви_4 Where [Номер родителя]="+IntToStr(Parent);
Branches->Active=true;
Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_4 Where Родитель="+IntToStr(Parent);
Nodes->Active=true;
if (Branches->RecordCount==0 & Nodes->RecordCount==0)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From Узлы_4 Where [Номер узла]="+IntToStr(Parent);
 Nodes->Active=true;
 if (Nodes->RecordCount==0)
 {
  ShowMessage("Ошибка! Удалялась запись в таблице узлов номер "+IntToStr(Number)+" предок номер "+IntToStr(Parent));
  ShowMessage("Не найдена запись в таблице узлов номер "+IntToStr(Parent));

 }
 else
 {
 int NumParent=Nodes->FieldByName("Родитель")->Value;
 AnsiString Text=Nodes->FieldByName("Название")->Value;
  if (NumParent!=0)
 {
 Nodes->Delete();
 Branches->Active=false;
 Branches->CommandText="Select * From Ветви_4";
 Branches->Active=true;
 Branches->Append();
 Branches->FieldByName("Номер родителя")->Value=NumParent;
 Branches->FieldByName("Название")->Value=Text;
 Branches->Post();
 Branches->Last();
 int Number=Branches->FieldByName("Номер ветви")->Value;
 N=TreeView2->GetNodeAt(XX,YY);
 TTreeNode * ParNode=N->Parent;
 PMyNode(ParNode->Data)->Number=Number;
 PMyNode(ParNode->Data)->Parent=NumParent;
 PMyNode(ParNode->Data)->Node=false;
 //
 //ParNode->Text=Text+" "+IntToStr(Number)+" "+IntToStr(NumParent)+" Ветвь";
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

//Номер предка
//Нужно узнать номер новой ветви
int NumParentNode=PMyNode(SelNode->Data)->Number;
//Является ли предок узлом?
IsNode=PMyNode(SelNode->Data)->Node;
if (IsNode==true)
{
//Если является то нужно только оформить на него ссылку
Branches->Active=false;
Branches->CommandText="Select * From Ветви_5";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("Номер родителя")->Value=NumParentNode;
Branches->FieldByName("Название")->Value="Новая территория";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("Номер ветви")->Value;

MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;

}
else
{
//А если является ветвью то нужно ее перенести в узлы и оформить на нее ссылку.
int Par;
AnsiString Text;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_5 Where [Номер ветви]="+IntToStr(NumParentNode);
Branches->Active=true;
Par=Branches->FieldByName("Номер родителя")->Value;
Text=Branches->FieldByName("Название")->Value;
Branches->Delete();
Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_5";
Nodes->Active=true;
Nodes->Append();
Nodes->FieldByName("Родитель")->Value=Par;
Nodes->FieldByName("Название")->Value=Text;
Nodes->Post();
Nodes->Last();

//PMyNode(TreeView1->Selected->Data)->Node
PMyNode(SelNode->Data)->Number=Nodes->FieldByName("Номер узла")->Value;
PMyNode(SelNode->Data)->Parent=Nodes->FieldByName("Родитель")->Value;
PMyNode(SelNode->Data)->Node=true;
NumParentNode=Nodes->FieldByName("Номер узла")->Value;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_5";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("Номер родителя")->Value=NumParentNode;
Branches->FieldByName("Название")->Value="Новая территория";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("Номер ветви")->Value;


MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;
}
TreeView3->Items->AddChildObject(SelNode,"Новая территория",MyNodePtr);        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N13Click(TObject *Sender)
{
//Добавить проверку на отсутствие у родителя удаляемой ветви (узла) потомков
//и перевода его в категорию ветвей
TTreeNode *N;
//N=TreeView1->GetNodeAt(XX,YY);

int Number=PMyNode(TreeView3->Selected->Data)->Number;
int Parent=PMyNode(TreeView3->Selected->Data)->Parent;
bool IsNode=PMyNode(TreeView3->Selected->Data)->Node;
if (IsNode==true)
{
DeleteNode("Узлы_5","Ветви_5", Number);

}
else
{
Comm->CommandText="Delete * From Ветви_5 Where [Номер ветви]="+IntToStr(Number);
Comm->Execute();

}

Branches->Active=false;
Branches->CommandText="Select * From Ветви_5 Where [Номер родителя]="+IntToStr(Parent);
Branches->Active=true;
Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_5 Where Родитель="+IntToStr(Parent);
Nodes->Active=true;
if (Branches->RecordCount==0 & Nodes->RecordCount==0)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From Узлы_5 Where [Номер узла]="+IntToStr(Parent);
 Nodes->Active=true;
 if (Nodes->RecordCount==0)
 {
  ShowMessage("Ошибка! Удалялась запись в таблице узлов номер "+IntToStr(Number)+" предок номер "+IntToStr(Parent));
  ShowMessage("Не найдена запись в таблице узлов номер "+IntToStr(Parent));

 }
 else
 {
 int NumParent=Nodes->FieldByName("Родитель")->Value;
 AnsiString Text=Nodes->FieldByName("Название")->Value;
  if (NumParent!=0)
 {
 Nodes->Delete();
 Branches->Active=false;
 Branches->CommandText="Select * From Ветви_5";
 Branches->Active=true;
 Branches->Append();
 Branches->FieldByName("Номер родителя")->Value=NumParent;
 Branches->FieldByName("Название")->Value=Text;
 Branches->Post();
 Branches->Last();
 int Number=Branches->FieldByName("Номер ветви")->Value;
 N=TreeView3->GetNodeAt(XX,YY);
 TTreeNode * ParNode=N->Parent;
 PMyNode(ParNode->Data)->Number=Number;
 PMyNode(ParNode->Data)->Parent=NumParent;
 PMyNode(ParNode->Data)->Node=false;
 //
 //ParNode->Text=Text+" "+IntToStr(Number)+" "+IntToStr(NumParent)+" Ветвь";
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


//Номер предка
//Нужно узнать номер новой ветви
int NumParentNode=PMyNode(SelNode->Data)->Number;
//Является ли предок узлом?
IsNode=PMyNode(SelNode->Data)->Node;
if (IsNode==true)
{
//Если является то нужно только оформить на него ссылку
Branches->Active=false;
Branches->CommandText="Select * From Ветви_6";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("Номер родителя")->Value=NumParentNode;
Branches->FieldByName("Название")->Value="Новый вид деятельности";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("Номер ветви")->Value;

MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;

}
else
{
//А если является ветвью то нужно ее перенести в узлы и оформить на нее ссылку.
int Par;
AnsiString Text;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_6 Where [Номер ветви]="+IntToStr(NumParentNode);
Branches->Active=true;
Par=Branches->FieldByName("Номер родителя")->Value;
Text=Branches->FieldByName("Название")->Value;
Branches->Delete();
Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_6";
Nodes->Active=true;
Nodes->Append();
Nodes->FieldByName("Родитель")->Value=Par;
Nodes->FieldByName("Название")->Value=Text;
Nodes->Post();
Nodes->Last();


PMyNode(SelNode->Data)->Number=Nodes->FieldByName("Номер узла")->Value;
PMyNode(SelNode->Data)->Parent=Nodes->FieldByName("Родитель")->Value;
PMyNode(SelNode->Data)->Node=true;
NumParentNode=Nodes->FieldByName("Номер узла")->Value;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_6";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("Номер родителя")->Value=NumParentNode;
Branches->FieldByName("Название")->Value="Новый вид деятельности";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("Номер ветви")->Value;


MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;
}
TreeView4->Items->AddChildObject(SelNode,"Новый вид деятельности",MyNodePtr);
        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N15Click(TObject *Sender)
{
//Добавить проверку на отсутствие у родителя удаляемой ветви (узла) потомков
//и перевода его в категорию ветвей
TTreeNode *N;
//N=TreeView1->GetNodeAt(XX,YY);

int Number=PMyNode(TreeView4->Selected->Data)->Number;
int Parent=PMyNode(TreeView4->Selected->Data)->Parent;
bool IsNode=PMyNode(TreeView4->Selected->Data)->Node;
if (IsNode==true)
{
DeleteNode("Узлы_6","Ветви_6", Number);

}
else
{
Comm->CommandText="Delete * From Ветви_6 Where [Номер ветви]="+IntToStr(Number);
Comm->Execute();

}

Branches->Active=false;
Branches->CommandText="Select * From Ветви_6 Where [Номер родителя]="+IntToStr(Parent);
Branches->Active=true;
Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_6 Where Родитель="+IntToStr(Parent);
Nodes->Active=true;
if (Branches->RecordCount==0 & Nodes->RecordCount==0)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From Узлы_6 Where [Номер узла]="+IntToStr(Parent);
 Nodes->Active=true;
 if (Nodes->RecordCount==0)
 {
  ShowMessage("Ошибка! Удалялась запись в таблице узлов номер "+IntToStr(Number)+" предок номер "+IntToStr(Parent));
  ShowMessage("Не найдена запись в таблице узлов номер "+IntToStr(Parent));

 }
 else
 {
 int NumParent=Nodes->FieldByName("Родитель")->Value;
 AnsiString Text=Nodes->FieldByName("Название")->Value;
  if (NumParent!=0)
 {
 Nodes->Delete();
 Branches->Active=false;
 Branches->CommandText="Select * From Ветви_6";
 Branches->Active=true;
 Branches->Append();
 Branches->FieldByName("Номер родителя")->Value=NumParent;
 Branches->FieldByName("Название")->Value=Text;
 Branches->Post();
 Branches->Last();
 int Number=Branches->FieldByName("Номер ветви")->Value;
 N=TreeView4->GetNodeAt(XX,YY);
 TTreeNode * ParNode=N->Parent;
 PMyNode(ParNode->Data)->Number=Number;
 PMyNode(ParNode->Data)->Parent=NumParent;
 PMyNode(ParNode->Data)->Node=false;
 //
 //ParNode->Text=Text+" "+IntToStr(Number)+" "+IntToStr(NumParent)+" Ветвь";
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


//Номер предка
//Нужно узнать номер новой ветви
int NumParentNode=PMyNode(SelNode->Data)->Number;
//Является ли предок узлом?
IsNode=PMyNode(SelNode->Data)->Node;
if (IsNode==true)
{
//Если является то нужно только оформить на него ссылку
Branches->Active=false;
Branches->CommandText="Select * From Ветви_7";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("Номер родителя")->Value=NumParentNode;
Branches->FieldByName("Название")->Value="Новый аспект";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("Номер ветви")->Value;

MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;

}
else
{
//А если является ветвью то нужно ее перенести в узлы и оформить на нее ссылку.
int Par;
AnsiString Text;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_7 Where [Номер ветви]="+IntToStr(NumParentNode);
Branches->Active=true;
Par=Branches->FieldByName("Номер родителя")->Value;
Text=Branches->FieldByName("Название")->Value;
Branches->Delete();
Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_7";
Nodes->Active=true;
Nodes->Append();
Nodes->FieldByName("Родитель")->Value=Par;
Nodes->FieldByName("Название")->Value=Text;
Nodes->Post();
Nodes->Last();


PMyNode(SelNode->Data)->Number=Nodes->FieldByName("Номер узла")->Value;
PMyNode(SelNode->Data)->Parent=Nodes->FieldByName("Родитель")->Value;
PMyNode(SelNode->Data)->Node=true;
NumParentNode=Nodes->FieldByName("Номер узла")->Value;
Branches->Active=false;
Branches->CommandText="Select * From Ветви_7";
Branches->Active=true;
Branches->Append();
Branches->FieldByName("Номер родителя")->Value=NumParentNode;
Branches->FieldByName("Название")->Value="Новый аспект";
Branches->Post();
Branches->Last();
Number=Branches->FieldByName("Номер ветви")->Value;


MyNodePtr = new TMyNode;
MyNodePtr->Number=Number;
MyNodePtr->Parent=NumParentNode;
MyNodePtr->Node=false;
}
TreeView5->Items->AddChildObject(SelNode,"Новый аспект",MyNodePtr);        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::N17Click(TObject *Sender)
{
//Добавить проверку на отсутствие у родителя удаляемой ветви (узла) потомков
//и перевода его в категорию ветвей
TTreeNode *N;


int Number=PMyNode(TreeView5->Selected->Data)->Number;
int Parent=PMyNode(TreeView5->Selected->Data)->Parent;
bool IsNode=PMyNode(TreeView5->Selected->Data)->Node;
if (IsNode==true)
{
DeleteNode("Узлы_7","Ветви_7", Number);

}
else
{
Comm->CommandText="Delete * From Ветви_7 Where [Номер ветви]="+IntToStr(Number);
Comm->Execute();

}

Branches->Active=false;
Branches->CommandText="Select * From Ветви_7 Where [Номер родителя]="+IntToStr(Parent);
Branches->Active=true;
Nodes->Active=false;
Nodes->CommandText="Select * From Узлы_7 Where Родитель="+IntToStr(Parent);
Nodes->Active=true;
if (Branches->RecordCount==0 & Nodes->RecordCount==0)
{
 Nodes->Active=false;
 Nodes->CommandText="Select * From Узлы_7 Where [Номер узла]="+IntToStr(Parent);
 Nodes->Active=true;
 if (Nodes->RecordCount==0)
 {
  ShowMessage("Ошибка! Удалялась запись в таблице узлов номер "+IntToStr(Number)+" предок номер "+IntToStr(Parent));
  ShowMessage("Не найдена запись в таблице узлов номер "+IntToStr(Parent));

 }
 else
 {
 int NumParent=Nodes->FieldByName("Родитель")->Value;
 AnsiString Text=Nodes->FieldByName("Название")->Value;
  if (NumParent!=0)
 {
 Nodes->Delete();
 Branches->Active=false;
 Branches->CommandText="Select * From Ветви_7";
 Branches->Active=true;
 Branches->Append();
 Branches->FieldByName("Номер родителя")->Value=NumParent;
 Branches->FieldByName("Название")->Value=Text;
 Branches->Post();
 Branches->Last();
 int Number=Branches->FieldByName("Номер ветви")->Value;
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
Tab->CommandText="Select * From "+NameNode+" Where [Родитель]="+IntToStr(Number);
Tab->Active=true;

if(Tab->RecordCount!=0)
{
for(Tab->First();!Tab->Eof;Tab->Next())
{
 int Parent=Tab->FieldByName("Номер узла")->AsInteger;
 DeleteNode(NameNode, NameBranch, Parent);
}

}
Comm->CommandText="Delete * From "+NameBranch+" Where [Номер родителя]="+IntToStr(Number);
Comm->Execute();

 Comm->CommandText="Delete * From "+NameNode+" Where [Номер узла]="+IntToStr(Number);
 Comm->Execute();



return -1;
}
//---------------------------------------------------------------------------
void __fastcall TDocuments::MenuItem3Click(TObject *Sender)
{
Podr->Insert();
Podr->FieldByName("Название подразделения")->Value="Новое подразделение";
Podr->Post();        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::MenuItem4Click(TObject *Sender)
{
// Socket->Socket->SendText("Command:5;2|"+IntToStr(NameDB.Length())+"#"+NameDB+"|"+ServerSQL.Length()+"#"+ServerSQL+"|");
 Zast->MClient->Act.ParamComm.clear();
 Zast->MClient->Act.ParamComm.push_back("DeletePodr");

String DBName="Аспекты";
String ServerSQL="SELECT Аспекты.[Номер аспекта], Аспекты.Подразделение FROM Аспекты";
String ClientSQL="SELECT [Номер аспекта], Подразделение FROM TempAspects";
Zast->MClient->ReadTable(DBName, ServerSQL, ClientSQL);

/*
int NumP=Podr->FieldByName("ServerNum")->AsInteger;

Zast->MClient->Start();
Table* STab=Zast->MClient->CreateTable(this, Zast->ServerName, Zast->VDB[Zast->GetIDDBName("Аспекты")].ServerDB);

STab->SetCommandText("SELECT Аспекты.[Номер аспекта], Аспекты.Подразделение FROM Аспекты WHERE (((Аспекты.Подразделение)="+IntToStr(NumP)+"));");
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
 S1="На сервере зафиксирован "+IntToStr(N)+" аспект, принадлежащий этому подразделению.";
 break;
 }
 case 2: case 3: case 4:
 {
 S1="На сервере зафиксировано "+IntToStr(N)+" аспекта, принадлежащих этому подразделению.";
 break;
 }
 default:
 {
 S1="На сервере зафиксировано "+IntToStr(N)+" аспектов, принадлежащих этому подразделению.";
 }
}

String S=S1+"\rИспользуя \"Движение аспектов\" предварительно освободите удаляемое подразделение от аспектов";
Application->MessageBoxA(S.c_str(),"Удаление подразделения",MB_ICONEXCLAMATION);
}
Zast->MClient->DeleteTable(this, STab);
Zast->MClient->Stop();
*/
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::MenuItem1Click(TObject *Sender)
{
// Добавить
Ins=true;
ADODataSet1->Last();
int Max=ADODataSet1->FieldByName("Макс граница")->Value;
bool Cr=ADODataSet1->FieldByName("Критерий")->Value;
ADODataSet1->Append();
ADODataSet1->FieldByName("Мин граница")->Value=Max;
ADODataSet1->FieldByName("Макс граница")->Value=Max+1;
ADODataSet1->FieldByName("Критерий")->Value=Cr;
ADODataSet1->FieldByName("Наименование значимости")->Value="Новая значимость";
ADODataSet1->FieldByName("Необходимая мера")->Value="Новая необходимая мера";

if (Cr==true)
{
ADODataSet1->FieldByName("Критерий1")->Value="Да";
}
else
{
ADODataSet1->FieldByName("Критерий1")->Value="Нет";
}

ADODataSet1->Post();
ADODataSet1->Last();
Ins=false;        
}
//---------------------------------------------------------------------------

void __fastcall TDocuments::MenuItem2Click(TObject *Sender)
{
// Удалить

ADODataSet1->Delete();
ADODataSet2->Active=false;
ADODataSet2->CommandText="Select * From Значимость Order By [Мин граница]";
ADODataSet2->Active=true;
ADODataSet2->Last();

int Min=ADODataSet2->FieldByName("Мин граница")->Value;

ADODataSet2->Prior();
for(int i=1;i<ADODataSet2->RecordCount;i++)
{
ADODataSet2->Edit();
// int Min=ADODataSet1->FieldByName("Мин граница")->Value;


ADODataSet2->FieldByName("Макс граница")->Value=Min-1;
Min=ADODataSet2->FieldByName("Мин граница")->Value;
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
Sit->FieldByName("Название ситуации")->Value="Новая ситуация";
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

 /*
ADODataSet1->Active=false;
ADODataSet1->CommandText="select * from Значимость Order By [Мин граница]";
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

void __fastcall TDocuments::ADODataSet1BeforePost(TDataSet *DataSet)
{
if(C==false)
{
AnsiString A=DataSet->FieldByName("Критерий1")->Value;
if(A=="Да")
{
 DataSet->FieldByName("Критерий")->Value=true;
 Kriteriy=true;
}
else
{
 DataSet->FieldByName("Критерий")->Value=false;
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
 DataSet->FieldByName("Критерий1")->Value="Да";
 DataSet->FieldByName("Критерий")->Value=true;
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
 DataSet->FieldByName("Критерий1")->Value="Нет";
 DataSet->FieldByName("Критерий")->Value=false;
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
ADODataSet1->CommandText="select * from Значимость Order By [Мин граница]";
ADODataSet1->Active=true;
DataSet->Last();



int Min=DataSet->FieldByName("Мин граница")->Value;
bool Cr=DataSet->FieldByName("Критерий")->Value;

DataSet->Edit();
DataSet->FieldByName("Макс граница")->Value=Min+1;
DataSet->Post();

DataSet->Prior();
for(int i=1;i<DataSet->RecordCount;i++)
{
DataSet->Edit();
DataSet->FieldByName("Макс граница")->Value=Min-1;
if(Cr==false)
{
DataSet->FieldByName("Критерий")->Value=false;
DataSet->FieldByName("Критерий1")->Value="Нет";

}
DataSet->Post();
Min=DataSet->FieldByName("Мин граница")->Value;
Cr=DataSet->FieldByName("Критерий")->Value;
DataSet->Prior();
}
DataSet->First();
DataSet->Edit();
DataSet->FieldByName("Мин граница")->Value=0;
DataSet->Post();
ADODataSet1->Active=false;
ADODataSet1->CommandText="select * from Значимость Order By [Мин граница]";
ADODataSet1->Active=true;
DataSet->First();

//bool K=false;
for(int i=0;i<DataSet->RecordCount;i++)
{
if(DataSet->FieldByName("Мин граница")->Value>DataSet->FieldByName("Макс граница")->Value)
{
 ShowMessage("В списке уже есть минмальная граница равная "+DataSet->FieldByName("Мин граница")->Value+" исправьте ее");
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
 Nodes->CommandText="Select * From Узлы_3 Where [Номер узла]="+IntToStr(Number);
 Nodes->Active=true;
 Nodes->Edit();
 Nodes->FieldByName("Название")->Value=S;
 Nodes->Post();

}
else
{
 Branches->Active=false;
 Branches->CommandText="Select * From Ветви_3 Where [Номер ветви]="+IntToStr(Number);
 Branches->Active=true;
 Branches->Edit();
 Branches->FieldByName("Название")->Value=S;
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
 Nodes->CommandText="Select * From Узлы_4 Where [Номер узла]="+IntToStr(Number);
 Nodes->Active=true;
 Nodes->Edit();
 Nodes->FieldByName("Название")->Value=S;
 Nodes->Post();

}
else
{
 Branches->Active=false;
 Branches->CommandText="Select * From Ветви_4 Where [Номер ветви]="+IntToStr(Number);
 Branches->Active=true;
 Branches->Edit();
 Branches->FieldByName("Название")->Value=S;
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
 Nodes->CommandText="Select * From Узлы_6 Where [Номер узла]="+IntToStr(Number);
 Nodes->Active=true;
 Nodes->Edit();
 Nodes->FieldByName("Название")->Value=S;
 Nodes->Post();

}
else
{
 Branches->Active=false;
 Branches->CommandText="Select * From Ветви_6 Where [Номер ветви]="+IntToStr(Number);
 Branches->Active=true;
 Branches->Edit();
 Branches->FieldByName("Название")->Value=S;
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
 Nodes->CommandText="Select * From Узлы_7 Where [Номер узла]="+IntToStr(Number);
 Nodes->Active=true;
 Nodes->Edit();
 Nodes->FieldByName("Название")->Value=S;
 Nodes->Post();

}
else
{
 Branches->Active=false;
 Branches->CommandText="Select * From Ветви_7 Where [Номер ветви]="+IntToStr(Number);
 Branches->Active=true;
 Branches->Edit();
 Branches->FieldByName("Название")->Value=S;
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
 Nodes->CommandText="Select * From Узлы_5 Where [Номер узла]="+IntToStr(Number);
 Nodes->Active=true;
 Nodes->Edit();
 Nodes->FieldByName("Название")->Value=S;
 Nodes->Post();

}
else
{
 Branches->Active=false;
 Branches->CommandText="Select * From Ветви_5 Where [Номер ветви]="+IntToStr(Number);
 Branches->Active=true;
 Branches->Edit();
 Branches->FieldByName("Название")->Value=S;
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
  Application->HelpJump("IDH_СТАРТ_ГЛАВСПЕЦ");        
}
//---------------------------------------------------------------------------

