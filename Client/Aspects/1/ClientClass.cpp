//---------------------------------------------------------------------------


#pragma hdrstop

#include "ClientClass.h"
#include "MasterPointer.h"
#include "Main.h"
#include "Zastavka.h"
//---------------------------------------------------------------------------

#pragma package(smart_init)
Client::Client(TClientSocket *ClientSocket, TActionManager *ActMan, TForm *Form)
{
Socket=ClientSocket;
ActionManager=ActMan;
Owner=Form;
IP="�� ����������";
Login="�� �����";
}
//********************************************************************
Client::~Client()
{

}
//********************************************************************
void Client::Connect(String ServerName, int Port)
{
//��������� ����� �� ������� 1

Act.WaitCommand=1;
Socket->Address=ServerName;
Socket->Host=ServerName;
Socket->Port=Port;
Socket->Active=true;
}
//********************************************************************
void Client::CommandExec(int Comm, vector<String>Parameters)
{
if(Act.WaitCommand==Comm | Act.WaitCommand==0 | Comm==7)
{
switch(Comm)
{
 case 0:
 {
 Act.WaitCommand=0;
 String Mess="Command:0;1|"+IntToStr(IP.Length())+"#"+IP+"|";
 Socket->Socket->SendText(Mess);
 break;
 }
 case 1:
 {

  //�������� �� ������� 1
  String S1=GetIP();
  IP=S1;
  String S2=Application->ExeName;
  String Mess="Command:1;2|"+IntToStr(S1.Length())+"#"+GetIP()+"|"+IntToStr(S2.Length())+"#"+Application->ExeName;
  //���������� ��������� ��� ��������� ������� ����� 2
  Act.WaitCommand=2;
  Socket->Socket->SendText(Mess);
  break;
 }
 case 2:
 {
 //����� ������� ��� ������ ����� � ���������� ����� ������������� ������� ������ ������������ �����
 //��������� �������� - ����������� ������� �����.
 //����� ��������� ��� ����������� �����
 Act.WaitCommand=Act.NextCommand;
 StartAction("PrepareConnectBase");

 break;
 }
 case 3:
 {


 //������ ������ � ���������� ���������
 String Action=Act.ParamComm[0];

 Act.ParamComm.clear();

 Act.WaitCommand=0;
 String NumPrevForm=Parameters[0];
 Act.ParamComm.push_back(NumPrevForm);
 StartAction(Action);
 break;
 }
 case 4:
 {
 /*
 String Action=Act.ParamComm[0];
 StartAction(Action);
 */

 /*
 ShowMessage(Act.ParamComm[0]);
 ShowMessage(Act.ParamComm[1]);
 ShowMessage(Act.ParamComm[2]);
 ShowMessage(Act.ParamComm[3]);
 ShowMessage(Act.ParamComm[4]);
 ShowMessage(Act.ParamComm[5]);
 ShowMessage(Act.ParamComm[6]);
 */
 RegisterDatabase(Parameters[2], StrToInt(Parameters[1]));
 this->VTrigger[0].Var++;
 this->ActTrigger(0);

 break;
 }
 case 5:
 {
 //ShowMessage(Parameters[0]);
 DecodeTable(Act.ParamComm[1],Act.ParamComm[2],Parameters[0]);
 StartAction(Act.ParamComm[0]);
 break;
 }
 case 6:
 {

 //����� - ���������� ������ �������������
 //����� ���� ����� ���������� Form1 ����� ����������.
Zast->BeginWork->Execute();
 break;
 }
 case 7:
 {
  //����� �� ������ � ������
  //���������� ����� �� ����������

 StartAction(Act.ParamComm[0]);
 break;
 }
 case 8:
 {
 //����� �� ������� ������ �������
 StartAction(Act.ParamComm[0]);
 break;
 }
 case 9:
 {
 //ShowMessage("������ ������ ���������");
 StartAction("LoadNewLogins");
 break;
 }
 case 10:
 {
 if(Parameters[0]=="0")
 {
 //������� �� ���������� �������. � ���������� ��������.

 ShowMessage("� ���������� ��������. ������ � ����� ������� ��� ��������� � �������");
 Zast->Close();
 }
 else
 {
  //������� �� ����������� ������
//  Zast->PrepareUpdateOtd->Execute();
 }
 break;
 }
}
}
else
{
 ShowMessage("��������� ������� "+IntToStr(Act.WaitCommand)+" � ������ ����� �� ������� "+IntToStr(Comm));
}
}
//*********************************************************************
String Client::GetIP()
{
    String addr="";

    const int WSVer = 0x101;
    WSAData wsaData;
    hostent *h;
    char Buf[128];
    if (WSAStartup(WSVer, &wsaData) == 0)
    {
    if (gethostname(&Buf[0], 128) == 0)
    {
    h = gethostbyname(&Buf[0]);
    if (h != NULL)
    {
    //ShowMessage(inet_ntoa (*(reinterpret_cast<in_addr *>(*(h->h_addr_list)))));
    addr=inet_ntoa (*(reinterpret_cast<in_addr *>(*(h->h_addr_list))));

    }
    else MessageBox(0,"�� �� � ����. � IP ������ � ��� ���.",0,0);
    }
    WSACleanup;
    }


    return addr;
}
//**************************************************************************
Form* Client::RegForm(String FormName)
{
Form* F=new Form();
F->RegForm(FormName);
VForms.push_back(F);
//��������� ����������� ����� �� �������, ������� 3
  String Mess="Command:3;1|"+IntToStr(FormName.Length())+"#"+FormName;
  //���������� ��������� ��� ��������� ������� ����� 3
  Act.WaitCommand=3;
  Socket->Socket->SendText(Mess);

return F;
}
//****************************************************************************
void Client::StartAction(String NameAction)
{
for(int i=0;i<ActionManager->ActionCount;i++)
{
 if(ActionManager->Actions[i]->Name==NameAction)
 {
  ActionManager->Actions[i]->Execute();
 }
}
}
//***************************************************************************
void Client::ConnectDatabase(String Name,int Number, bool Connect)
{
String C;
if(Connect)
{
C="true";
}
else
{
C="false";
}
Act.WaitCommand=4;
String Mess="Command:4;3|"+IntToStr(Name.Length())+"#"+Name+"|"+IntToStr(Number).Length()+"#"+IntToStr(Number)+"|"+C.Length()+"#"+C+"|";
Socket->Socket->SendText(Mess);

}
//*************************************************************************
void Client::ReadTable(String NameDB, String ServerSQL, String ClientSQL)
{
 //Act.ParamComm.clear();
 Act.WaitCommand=5;
 Act.ParamComm.push_back(NameDB);
 Act.ParamComm.push_back(ClientSQL);

 Socket->Socket->SendText("Command:5;2|"+IntToStr(NameDB.Length())+"#"+NameDB+"|"+ServerSQL.Length()+"#"+ServerSQL+"|");
}
//*************************************************************************
void Client::DecodeTable(String NameDB, String ClientSQL, String Text)
{
String Record;
String Field;
try
{
MDBConnector* DB;
if(NameDB=="�������")
{
DB=ADOAspect;
}
if(NameDB=="Reference")
{
DB=ADOConn;
}
int FromPos=ClientSQL.LowerCase().Pos("from");
String DelText="Delete * "+ClientSQL.SubString(FromPos, ClientSQL.Length());
FromPos=DelText.LowerCase().Pos("order");
if(FromPos>0)
{
DelText=DelText.SubString(0, FromPos-2);
}

 MP<TADOCommand>Comm(Owner);
 Comm->Connection=DB;
 Comm->CommandText=DelText;
 Comm->Execute();

 MP<TADODataSet>Tab(Owner);
 Tab->Connection=DB;
 Tab->CommandText=ClientSQL;
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

   if(Text.SubString(1,2)==BeginTable | Text.SubString(Text.Length()-2,2)==EndTable)
   {
    Text=Text.SubString(3,Text.Length()-4);
    do
    {
    if(Text.SubString(1,2)==BeginRecord)
    {
     int EndRecordPos=Text.Pos(EndRecord);
     Record=Text.SubString(3,EndRecordPos-3);

     Tab->Insert();
     int i=0;
     do
     {
      if(Record.SubString(1,2)==BeginField)
      {
       int EndFieldPos=Record.Pos(EndField);
       Field=Record.SubString(3, EndFieldPos-3);

       char T=Field[1];
       int Type=T;

       switch(Type)
       {
        case 1:
        {
        Tab->FieldList->Fields[i]->Value=Field.SubString(2, Field.Length());
        break;
        }
        case 2:
        {
        String F=Field.SubString(4, Field.Length());
        Tab->FieldList->Fields[i]->Value=StrToInt(F);
        break;
        }
        case 3:
        {

        break;
        }
        case 4:
        {

        }
        case 5:
        {
        String F=Field.SubString(4, Field.Length());
        Tab->FieldList->Fields[i]->Value=StrToDateTime(F);
        break;
        }
        case 6:
        {
        String F=Field.SubString(4, Field.Length());
        Tab->FieldList->Fields[i]->Value=F;
        break;
        }

       }
       i++;

       Record=Record.SubString(EndFieldPos+2, Record.Length());
      }
      else
      {
       ShowMessage("��� ����� ������ ����");
      }
     }
     while(Record.Length()!=0);
       Tab->Post();
     Text=Text.SubString(EndRecordPos+2, Text.Length());
    }
    else
    {
//     ShowMessage("��� ����� ������ ������");
    }
    }
    while(Text.Length()!=0);
   }
   else
   {
    ShowMessage("��� ������ ������ � ����� ������� "+ClientSQL+" ���� "+NameDB);
   }
}
catch (...)
{
 ShowMessage(Text);
 ShowMessage(Record);
 ShowMessage(Field);
}
}
//**************************************************************************
void Client::LoginResult(String Login, String Pass, int NumRole, bool Ok)
{
//�������� �� ������ ��������� ������, ���������� ������ (���� �������� �� ������) � ���������� ��������.
//�������� ������ ��������� �� ������� ��� ������������ ���� � ��������������
//�� ������� ������� ������ ��������� � ����������� �������. ���� ������ ������� �� ����� ������ �� ���������� ������.
 Act.ParamComm.clear();
 Act.ParamComm.push_back(Login);
 Act.ParamComm.push_back(IntToStr(NumRole));
 Act.WaitCommand=0;

if(Ok)
{
Socket->Socket->SendText("Command:6;3|"+IntToStr(Login.Length())+"#"+Login+"|"+"4#Pass"+"|"+"1#1"+"|");
}
else
{
Socket->Socket->SendText("Command:6;3|"+IntToStr(Login.Length())+"#"+Login+"|"+IntToStr(Pass.Length())+"#"+Pass+"|"+"1#0"+"|");
}
}
//***************************************************************************
int Client::GetIDDBName(String Name)
{

int Ret=-1;
for(unsigned int i=0;i<VDB.size();i++)
{
 if(VDB[i].Name==Name)
 {
  Ret=i;
  break;
 }
}
return Ret;

}
//***************************************************************************
int Client::GetLicenseCount(String DBName)
{
int Res=-1;
for(unsigned int i=0;i<VDB.size();i++)
{
 if(VDB[i].Name==DBName)
 {
  Res=VDB[i].NumLicense;
  break;
 }
}
return Res;
}
//*************************************************************************
void Client::RegisterDatabase(String Name, int NumLic)
{
for(unsigned int i=0;i<VDB.size();i++)
{
 if(VDB[i].Name==Name)
 {
  VDB[i].NumLicense=NumLic;
  if(NumLic!=0)
  {
   this->Reg=true;
  }
  else
  {
   this->Reg=false;
  }
  break;
 }
}
}
//*************************************************************************
void Client::VerifyLicense(String NameDB)
{
/*
int NumDB;
int LicCount;
for(unsigned int i=0;i<VDB.size();i++)
{
 if(VDB[i].Name==NameDB)
 {
  NumDB=VDB[i].NumDatabase;
  LicCount=VDB[i].NumLicense;
  break;
 }
}

MP<TADODataSet>Log3(Owner);
Log3->Connection=Database;
Log3->CommandText="Select * From Logins Where Role=3 AND NumDatabase="+IntToStr(NumDB);
Log3->Active=true;

if(LicCount==-1)
{
//Main->MClient->WriteDiaryEvent("AdminARM ��������","������������� ��������","���� ������: "+CBDatabase->Text);

}
else
{
if(Log3->RecordCount>LicCount)
{
//Main->MClient->WriteDiaryEvent("AdminARM ��������","������� ����� ������������� �� ������� ��������","���� ������: "+CBDatabase->Text+" ����: "+IntToStr(Log3->RecordCount)+" �����: "+IntToStr(LicCount));
//Main->MClient->WriteDiaryEvent("AdminARM ��������","������ ������ �������","���� ������: "+CBDatabase->Text+" ����: "+IntToStr(Log3->RecordCount)+" �����: "+IntToStr(LicCount));

StartAction("LoadLogins");
//LoadLogins();
 //�������� �������
// Main->MClient->WriteDiaryEvent("AdminARM ��������","�������� ������� ���������","���� ������: "+CBDatabase->Text);

}
}
*/
}
//************************************************************************
void Client::WriteDiaryEvent(String Type, String Name, String Prim)
{
Act.WaitCommand=0;
Act.ParamComm.clear();
Socket->Socket->SendText("Command:7;5|"+IntToStr(IP.Length())+"#"+IP+"|"+IntToStr(Login.Length())+"#"+Login+"|"+IntToStr(Type.Length())+"#"+Type+"|"+IntToStr(Name.Length())+"#"+Name+"|"+IntToStr(Prim.Length())+"#"+Prim+"|");

}
//************************************************************************
void Client::WriteDiaryEvent(String Type, String Name)
{
WriteDiaryEvent(Type, Name, "");
}
//************************************************************************
String Client::TableToStr(String NameDB, String SQLText)
{
   MP<TADODataSet>Tab(Owner);
   //Tab->Connection=Database;
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
      case ftAutoInc:
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
//****************************************************************************
void Client::WriteTable(String Database, String ClientSQLText, String ServerSQLText)
{
 String Text=TableToStr(Database, ClientSQLText);
 Socket->Socket->SendText("Command:8;3|"+IntToStr(Database.Length())+"#"+Database+"|"+IntToStr(ServerSQLText.Length())+"#"+ServerSQLText+"|"+IntToStr(Text.Length())+"#"+Text+"|");
}
//***************************************************************************
void Client::ActTrigger(int NumTrigger)
{
Sleep(200);
if(VTrigger[NumTrigger].Var<VTrigger[NumTrigger].Max)
{

 StartAction(VTrigger[NumTrigger].TrueAction);
 //VTrigger[NumTrigger].Var++;
}
else
{
 StartAction(VTrigger[NumTrigger].FalseAction);
}
}
//****************************************************************************
/////////////////////////////////////////////////////////////////////////////
Form::Form()
{
IDF=-1;
FormName="-";
}
//////////////////////////////////////////////////////////////////////////////
Form::~Form()
{

}
/////////////////////////////////////////////////////////////////////////////
void Form::RegForm(String FormName)
{
this->FormName=FormName;

}
//////////////////////////////////////////////////////////////////////////////
