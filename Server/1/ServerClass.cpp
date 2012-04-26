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
   this->Socket->SendText("Command:3;1|"+IntToStr(IntToStr(F->IDF).Length())+"#"+IntToStr(F->IDF)+"|");
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
   this->Socket->SendText("Command:4;1|1#1|");
   break;
      }
  }
  break;
  }
  case 5:
   {
   //����� ������� ������ �������
   String Text=TableToStr(Parameters[0], Parameters[1]);
   //ShowMessage(Text);
   this->Socket->SendText("Command:5;1|"+IntToStr(Text.Length())+"#"+Text+"|");
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
 return 0;//������, �� ������� ���� ������
}
//---------------------------------------------------------------------------
String Client::TableToStr(String NameDB, String SQLText)
{
   MP<TADODataSet>Tab((TForm*)Parent->pOwner);
   Tab->Connection=GetDatabase(NameDB);
   Tab->CommandText=SQLText;
   Tab->Active=true;

   //������ �������� �������
   //������ ������� 27 1
   //int Esc=27;
   int S1=1;
   int S2=2;
   int S3=3;
   int S4=4;
   int S5=5;
   int S6=6;
   int S7=7;
   
   char C=VK_ESCAPE;
   char C1=((char)S1);
   char C2=((char)S2);
   char C3=((char)S3);
   char C4=((char)S4);
   char C5=((char)S5);
   char C6=((char)S6);
   char C7=((char)S7);

   String BeginTable=(String)C+(String)C1;
   //����� ������� 27 2
   String EndTable=(String)C+(String)C2;
   //������ ������ 27 3
   String BeginRecord=(String)C+(String)C3;
   //����� ������ 27 4
   String EndRecord=(String)C+(String)C4;
   //������ ���� 27 5
   String BeginField=(String)C+(String)C5;
   //����� ���� 27 6
   String EndField=(String)C+(String)C6;
   //���� ������� �� ���� ���� � �������� ������������ ��������� 27 6
   String FieldSeparator=(String)C+(String)C7;
   //���� �����
   //ftString - 1
   //ftInteger - 2
   //ftBoolean - 3
   //ftFloat - 4
   //ftDateTime - 5
   //ftMemo - 6
   //���� Memo ���� �������������� ��� ������ � ��� ���������
   //��� ���� ����� ��������� �������������, Boolean true - 1, false - 0
   String Ret=BeginTable;
   for(Tab->First();!Tab->Eof;Tab->Next())
   {
   Ret=Ret+BeginRecord;
    for(int i=0; i<Tab->FieldCount;i++)
    {
    Ret=Ret+BeginField;
     switch (Tab->FieldList->Fields[i]->DataType)
     {
      case ftWideString:
      {
      Ret=Ret+(String)C1+Tab->FieldList->Fields[i]->AsString;
      break;
      }
      case ftInteger:
      {
      Ret=Ret+(String)C2+FieldSeparator+IntToStr(Tab->FieldList->Fields[i]->AsInteger);
      break;
      }
      case ftBoolean:
      {
      if(Tab->FieldList->Fields[i]->AsBoolean)
      {
      Ret=Ret+(String)C3+FieldSeparator+((char)1);
      }
      else
      {
      Ret=Ret+(String)C3+FieldSeparator+((char)0);
      }
      break;
      }
      case ftFloat:
      {
      Ret=Ret+(String)C4+FieldSeparator+FloatToStr(Tab->FieldList->Fields[i]->AsFloat);
      break;
      }
      case ftDateTime:
      {
      Ret=Ret+(String)C5+FieldSeparator+Tab->FieldList->Fields[i]->AsDateTime.DateTimeString();
      break;
      }
      case ftMemo:
      {
      Ret=Ret+(String)C6+FieldSeparator+Tab->FieldList->Fields[i]->AsString;
      break;
      }
     }
     Ret=Ret+EndField;
    }
    Ret=Ret+EndRecord;
   }
   Ret=Ret+EndTable;
   return Ret;
}
//--------------------------------------------------------------------------
//***************************************************************************
 mForm::mForm()
 {

 }
 //**************************************************************************
 mForm::~mForm()
 {

 }
 //**************************************************************************