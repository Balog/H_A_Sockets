//---------------------------------------------------------------------------


#pragma hdrstop

#include "ServerClass.h"
#include "MasterPointer.h"
//---------------------------------------------------------------------------

#pragma package(smart_init)

Clients::Clients(TComponent* Owner, String DiaryBase)
{
pOwner=Owner;
this->DiaryBase=DiaryBase;
//VBases=VB;
}
//-------------------------------------------------------
Clients::~Clients()
{
for(unsigned int i=0; i<VBases.size();i++)
{
  delete VBases[i].Database;
  delete VBases[i].Command;
}
VBases.clear();

for(unsigned int i=0;i<VClients.size();i++)
{
delete VClients[i];
}
VClients.clear();
}
//-------------------------------------------------------
//********************************************************
Client::Client(Clients* Cls)
{
 Parent=Cls;
}
//**************************************************
Client::~Client()
{
for(unsigned int i=0;i<VForm.size();i++)
{
delete VForm[i];
}
VForm.clear();
}
//***************************************************
void Client::CommandExec(int Comm, vector<String>Parameters)
{
if(LastCommand==Comm | LastCommand==0)
{
 switch(Comm)
 {
  case 1:
  {
   //����� �� ������� ������� IP � �������
   this->IP=Parameters[0];
   this->AppPatch=Parameters[1];
   //�������� ������� ���������� � ���������� ������
   //����� �� ���� �� ������������
   LastCommand=0;
   this->Socket->SendText("Command:2;0|");
  break;
  }
  case 3:
  {
  //������������ ����������� ����� �� �������
//   ShowMessage(Parameters[0]);
   LastCommand=0;
   mForm *F=new mForm();
   F->IDF=VForm.size();
   F->NameForm=Parameters[0];
   VForm.push_back(F);
   //���� �������� ���� � ��� �� �������� � ������ ������������ ����� ����� �� �������
   this->Socket->SendText("Command:3;1|"+IntToStr(F->IDF)+"|");
   break;
  }
  case 4:
  {
  LastCommand=0;
  for(unsigned int i=0;i<Parent->VBases.size();i++)
  {
   if(Parent->VBases[i].Name==Parameters[0])
   {
    if(Parameters[1]=="true")
    {
    Parent->VBases[i].Database->Connected=true;
    }
    else
    {
     Parent->VBases[i].Database->Connected=false;
    }
   //������� �������� ����
   this->Socket->SendText("Command:4;1|1|");
   break;
      }
  }
  break;
  }
  case 5:
   {
   //����� ������� ������ �������

   MP<TADODataSet>Tab((TForm*)Parent->pOwner);
   Tab->Connection=GetDatabase(Parameters[0]);
   Tab->CommandText=Parameters[1];
   Tab->Active=true;
   break;
   }
 }

}
}
//---------------------------------------------------------------------------
TADOConnection* Client::GetDatabase(String NameDB)
{

 for(unsigned int i=0;i<Parent->VBases.size();i++)
 {
  if(Parent->VBases[i].Name==NameDB)
  {
   return Parent->VBases[i].Database;
  }
 }
}
//---------------------------------------------------------------------------
//***************************************************************************
 mForm::mForm()
 {

 }
 //**************************************************************************
 mForm::~mForm()
 {

 }
 //**************************************************************************