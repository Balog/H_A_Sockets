//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Main.h"
#include "MasterPointer.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm1::N1Click(TObject *Sender)
{
if(OpenDialog1->Execute())
{
 String Patch=OpenDialog1->FileName;
 String Name=ExtractFileName(Patch);

 try
 {
  Table->Insert();
  Table->FieldByName("Patch")->Value=Patch;
  Table->FieldByName("Result")->Value="�� ���������";
  Table->Post();


 }
 catch(...)
 {
  ShowMessage("������ ����������� � �����\n�������� ��������� ����");
 }
}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::FormShow(TObject *Sender)
{
String Path=ExtractFilePath(Application->ExeName);
String DBPatch=Path+"Database\\DB.dtb";

Database->Connected=false;
Database->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+DBPatch+";Persist Security Info=False";
Database->LoginPrompt=false;
Database->Connected=true;

Table->Active=false;
Table->CommandText="Select * from Users Order by Num";
Table->Active=true;

for(Table->First();!Table->Eof;Table->Next())
{
 Table->Edit();
 Table->FieldByName("Result")->Value="�� ���������";
 Table->Post();
}
}
//---------------------------------------------------------------------------
void __fastcall TForm1::N2Click(TObject *Sender)
{
Table->Delete();
}
//---------------------------------------------------------------------------


void __fastcall TForm1::Button1Click(TObject *Sender)
{
if(OpenDialog2->Execute())
{
 AspRefs->Text=OpenDialog2->FileName;
 Edit1->Text="�� ���������";
}        
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button2Click(TObject *Sender)
{
MP<TADOCommand>Comm(this);
if(AspRefs->Text!="" & Table->RecordCount!=0)
{
//����������� ��������
String Path=ExtractFilePath(Application->ExeName);
String PDiary=Path+"Templates\\Diary.dtb";
String PAspect=Path+"Templates\\Eco_Aspects.dtb";
String PRefs=Path+"Templates\\Reference.dtb";

String RDiary=Path+"Result\\Diary.dtb";
String RAspect=Path+"Result\\Eco_Aspects.dtb";
String RRefs=Path+"Result\\Reference.dtb";

DeleteFile(RDiary.c_str());
DeleteFile(RAspect.c_str());
DeleteFile(RRefs.c_str());

CopyFile(PDiary.c_str(), RDiary.c_str(), true);
CopyFile(PAspect.c_str(), RAspect.c_str(), true);
CopyFile(PRefs.c_str(), RRefs.c_str(), true);

try
{
//����������� ����� � ������
AspectTo->Connected=false;
AspectTo->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+RAspect+";Persist Security Info=False";
AspectTo->LoginPrompt=false;
AspectTo->Connected=true;


TempFrom->Connected=false;
TempFrom->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+AspRefs->Text+";Persist Security Info=False";
TempFrom->LoginPrompt=false;
TempFrom->Connected=true;


TempTo->Connected=false;
TempTo->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+RRefs+";Persist Security Info=False";
TempTo->LoginPrompt=false;
TempTo->Connected=true;

CopyTable("����_3","�����_3");
CopyTable("����_4","�����_4");
CopyTable("����_5","�����_5");
CopyTable("����_6","�����_6");
CopyTable("����_7","�����_7");

//����������� ��������
MP<TADODataSet>MetFrom(this);
MetFrom->Connection=TempFrom;
MetFrom->CommandText="select * from ��������";
MetFrom->Active=true;

MP<TADODataSet>MetTo(this);
MetTo->Connection=TempTo;
MetTo->CommandText="select * from ��������";
MetTo->Active=true;


Comm->Connection=TempTo;
Comm->CommandText="Delete * from ��������";
Comm->Execute();

MetTo->Insert();
MetTo->FieldByName("�����")->Value=MetFrom->FieldByName("�����")->Value;
MetTo->FieldByName("��������")->Value=MetFrom->FieldByName("��������")->Value;
MetTo->Post();

//����������� ����������
MP<TADODataSet>ZnFrom(this);
ZnFrom->Connection=TempFrom;
ZnFrom->CommandText="select * from ����������";
ZnFrom->Active=true;

MP<TADODataSet>ZnTo(this);
ZnTo->Connection=TempTo;
ZnTo->CommandText="select * from ����������";
ZnTo->Active=true;

MP<TADODataSet>ZnAsp(this);
ZnAsp->Connection=AspectTo;
ZnAsp->CommandText="select * from ����������";
ZnAsp->Active=true;

Comm->Connection=TempTo;
Comm->CommandText="Delete * From ����������";
Comm->Execute();

for(ZnFrom->First();!ZnFrom->Eof;ZnFrom->Next())
{
 ZnTo->Insert();
// ZnTo->FieldByName("����� ����������")->Value=ZnFrom->FieldByName("����� ����������")->Value;
 ZnTo->FieldByName("������������ ����������")->Value=ZnFrom->FieldByName("������������ ����������")->Value;
 ZnTo->FieldByName("��������1")->Value=ZnFrom->FieldByName("��������1")->Value;
 ZnTo->FieldByName("��������")->Value=ZnFrom->FieldByName("��������")->Value;
 ZnTo->FieldByName("��� �������")->Value=ZnFrom->FieldByName("��� �������")->Value;
 ZnTo->FieldByName("���� �������")->Value=ZnFrom->FieldByName("���� �������")->Value;
 ZnTo->FieldByName("����������� ����")->Value=ZnFrom->FieldByName("����������� ����")->Value;
 ZnTo->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From ����������";
Comm->Execute();

for(ZnFrom->First();!ZnFrom->Eof;ZnFrom->Next())
{
 ZnAsp->Insert();
// ZnAsp->FieldByName("����� ����������")->Value=ZnFrom->FieldByName("����� ����������")->Value;
 ZnAsp->FieldByName("����� ����������")->Value=ZnFrom->FieldByName("����� ����������")->Value;
 ZnAsp->FieldByName("������������ ����������")->Value=ZnFrom->FieldByName("������������ ����������")->Value;
 ZnAsp->FieldByName("��������1")->Value=ZnFrom->FieldByName("��������1")->Value;
 ZnAsp->FieldByName("��������")->Value=ZnFrom->FieldByName("��������")->Value;
 ZnAsp->FieldByName("��� �������")->Value=ZnFrom->FieldByName("��� �������")->Value;
 ZnAsp->FieldByName("���� �������")->Value=ZnFrom->FieldByName("���� �������")->Value;
 ZnAsp->FieldByName("����������� ����")->Value=ZnFrom->FieldByName("����������� ����")->Value;
 ZnAsp->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From ������";
Comm->Execute();

MP<TADODataSet>AspectFrom(this);
AspectFrom->Connection=TempTo;
AspectFrom->CommandText="select * from �����_7";
AspectFrom->Active=true;

MP<TADODataSet>AspTo(this);
AspTo->Connection=AspectTo;
AspTo->CommandText="select * from ������";
AspTo->Active=true;


for(AspectFrom->First();!AspectFrom->Eof;AspectFrom->Next())
{
 AspTo->Insert();
 if(!AspectFrom->FieldByName("�����")->AsBoolean)
 {
 AspTo->FieldByName("����� �������")->Value=0; 
 }
 else
 {
 AspTo->FieldByName("����� �������")->Value=AspectFrom->FieldByName("����� �����")->Value;
 }
 AspTo->FieldByName("������������ �������")->Value=AspectFrom->FieldByName("��������")->Value;
 AspTo->FieldByName("�����")->Value=AspectFrom->FieldByName("�����")->Value;
 AspTo->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From �����������";
Comm->Execute();

MP<TADODataSet>VozdFrom(this);
VozdFrom->Connection=TempTo;
VozdFrom->CommandText="select * from �����_3";
VozdFrom->Active=true;

MP<TADODataSet>VozdTo(this);
VozdTo->Connection=AspectTo;
VozdTo->CommandText="select * from �����������";
VozdTo->Active=true;

for(VozdFrom->First();!VozdFrom->Eof;VozdFrom->Next())
{
 VozdTo->Insert();
 if(!VozdFrom->FieldByName("�����")->AsBoolean)
 {
 VozdTo->FieldByName("����� �����������")->Value=0;
 }
 else
 {
 VozdTo->FieldByName("����� �����������")->Value=VozdFrom->FieldByName("����� �����")->Value;
 }
 VozdTo->FieldByName("������������ �����������")->Value=VozdFrom->FieldByName("��������")->Value;
 VozdTo->FieldByName("�����")->Value=VozdFrom->FieldByName("�����")->Value;
 VozdTo->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From ������������";
Comm->Execute();

MP<TADODataSet>DeyatFrom(this);
DeyatFrom->Connection=TempTo;
DeyatFrom->CommandText="select * from �����_6";
DeyatFrom->Active=true;

MP<TADODataSet>DeyatTo(this);
DeyatTo->Connection=AspectTo;
DeyatTo->CommandText="select * from ������������";
DeyatTo->Active=true;

for(DeyatFrom->First();!DeyatFrom->Eof;DeyatFrom->Next())
{
 DeyatTo->Insert();
 if(!DeyatFrom->FieldByName("�����")->AsBoolean)
 {
  DeyatTo->FieldByName("����� ������������")->Value=0;
 }
 else
 {
 DeyatTo->FieldByName("����� ������������")->Value=DeyatFrom->FieldByName("����� �����")->Value;
 }
 DeyatTo->FieldByName("������������ ������������")->Value=DeyatFrom->FieldByName("��������")->Value;
 DeyatTo->FieldByName("�����")->Value=DeyatFrom->FieldByName("�����")->Value;
 DeyatTo->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From ����������";
Comm->Execute();

MP<TADODataSet>ZFrom(this);
ZFrom->Connection=TempFrom;
ZFrom->CommandText="select * from ����������";
ZFrom->Active=true;

MP<TADODataSet>ZTo(this);
ZTo->Connection=AspectTo;
ZTo->CommandText="select * from ����������";
ZTo->Active=true;

for(ZFrom->First();!ZFrom->Eof;ZFrom->Next())
{
 ZTo->Insert();
 ZTo->FieldByName("����� ����������")->Value=ZFrom->FieldByName("����� ����������")->Value;
 ZTo->FieldByName("������������ ����������")->Value=ZFrom->FieldByName("������������ ����������")->Value;
 ZTo->FieldByName("��������")->Value=ZFrom->FieldByName("��������")->Value;
 ZTo->FieldByName("��������1")->Value=ZFrom->FieldByName("��������1")->Value;
 ZTo->FieldByName("��� �������")->Value=ZFrom->FieldByName("��� �������")->Value;
 ZTo->FieldByName("���� �������")->Value=ZFrom->FieldByName("���� �������")->Value;
 ZTo->FieldByName("����������� ����")->Value=ZFrom->FieldByName("����������� ����")->Value;
 ZTo->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From ����������";
Comm->Execute();

MP<TADODataSet>TerrFrom(this);
TerrFrom->Connection=TempTo;
TerrFrom->CommandText="select * from �����_5";
TerrFrom->Active=true;

MP<TADODataSet>TerrTo(this);
TerrTo->Connection=AspectTo;
TerrTo->CommandText="select * from ����������";
TerrTo->Active=true;

for(TerrFrom->First();!TerrFrom->Eof;TerrFrom->Next())
{
 TerrTo->Insert();
 if(!TerrFrom->FieldByName("�����")->AsBoolean)
 {
 TerrTo->FieldByName("����� ����������")->Value=0;
 }
 else
 {
 TerrTo->FieldByName("����� ����������")->Value=TerrFrom->FieldByName("����� �����")->Value;
 }
 TerrTo->FieldByName("������������ ����������")->Value=TerrFrom->FieldByName("��������")->Value;
 TerrTo->FieldByName("�����")->Value=TerrFrom->FieldByName("�����")->Value;
 TerrTo->Post();
}


Edit1->Text="���������";

Table->First();
 String P1=Table->FieldByName("Patch")->Value;
MP<TADOConnection>AspConnect(this);
AspConnect->Connected=false;
AspConnect->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+P1+";Persist Security Info=False";
AspConnect->LoginPrompt=false;
AspConnect->Connected=true;

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From �������� Where �����=True";
Comm->Execute();

MP<TADODataSet>SitFrom(this);
SitFrom->Connection=AspConnect;
SitFrom->CommandText="select * from ��������";
SitFrom->Active=true;

MP<TADODataSet>SitTo(this);
SitTo->Connection=AspectTo;
SitTo->CommandText="select * from ��������";
SitTo->Active=true;

Comm->Connection=TempTo;
Comm->CommandText="Delete * From ��������";
Comm->Execute();

MP<TADODataSet>SitTo2(this);
SitTo2->Connection=TempTo;
SitTo2->CommandText="select * from ��������";
SitTo2->Active=true;

for(SitFrom->First();!SitFrom->Eof;SitFrom->Next())
{
 SitTo2->Insert();
 SitTo2->FieldByName("����� ��������")->Value=SitFrom->FieldByName("����� ��������")->AsInteger;
 SitTo2->FieldByName("�������� ��������")->Value=SitFrom->FieldByName("�������� ��������")->AsString;
 SitTo2->Post();

 SitTo->Insert();
 SitTo->FieldByName("����� ��������")->Value=SitFrom->FieldByName("����� ��������")->AsInteger;
 SitTo->FieldByName("�������� ��������")->Value=SitFrom->FieldByName("�������� ��������")->Value;
 SitTo->FieldByName("�����")->Value=true;
 SitTo->Post();
}

Comm->Connection=AspectTo;
Comm->CommandText="Delete * From �������������";
Comm->Execute();

for(Table->First();!Table->Eof;Table->Next())
{
try
{
 String P=Table->FieldByName("Patch")->Value;

AspConnect->Connected=false;
AspConnect->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+P+";Persist Security Info=False";
AspConnect->LoginPrompt=false;
AspConnect->Connected=true;

MP<TADODataSet>FromPodr(this);
FromPodr->Connection=AspConnect;
FromPodr->CommandText="Select * From ������������� Order by [����� �������������]";
FromPodr->Active=true;



MP<TADODataSet>AspectsTo(this);
AspectsTo->Connection=AspectTo;
AspectsTo->CommandText="Select * from �������";
AspectsTo->Active=true;




 SitTo2->Active=false;
 SitTo2->Active=true;
for(FromPodr->First();!FromPodr->Eof;FromPodr->Next())
{
 int NumPodrFrom=FromPodr->FieldByName("����� �������������")->Value;

 MP<TADODataSet>ToPodr(this);
 ToPodr->Connection=AspectTo;
 ToPodr->CommandText="Select * From ������������� Order by [����� �������������]";
 ToPodr->Active=true;

 ToPodr->Insert();
 ToPodr->FieldByName("�������� �������������")->Value=FromPodr->FieldByName("�������� �������������")->Value;
 ToPodr->Post();

 ToPodr->Active=false;
 ToPodr->Active=true;
 ToPodr->Last();

 int NumPodrTo=ToPodr->FieldByName("����� �������������")->Value;

 MP<TADODataSet>AspectsFrom(this);
 AspectsFrom->Connection=AspConnect;
 AspectsFrom->CommandText="Select * from ������� Where �������������="+IntToStr(NumPodrFrom)+" Order by [����� �������]";
 AspectsFrom->Active=true;
 int N;
 for(AspectsFrom->First();!AspectsFrom->Eof;AspectsFrom->Next())
 {

  N=0;
  AspectsTo->Insert();
  //AspectsTo->FieldByName("����� �������")->Value=AspectsFrom->FieldByName("����� �������")->AsInteger;
  AspectsTo->FieldByName("�������������")->Value=NumPodrTo;
  N=AspectsFrom->FieldByName("��������")->AsInteger;
  AspectsTo->FieldByName("��������")->Value=N;

  N=EncodeBranch("�����_5", AspectsFrom->FieldByName("��� ����������")->AsInteger);
  AspectsTo->FieldByName("��� ����������")->Value=N;//AspectsFrom->FieldByName("��� ����������")->Value;
  N=EncodeBranch("�����_6", AspectsFrom->FieldByName("������������")->AsInteger);
  AspectsTo->FieldByName("������������")->Value=N;

  AspectsTo->FieldByName("�������������")->Value=AspectsFrom->FieldByName("�������������")->AsString+" ";

  //ShowMessage(AspectsFrom->FieldByName("������")->Value);
  N=EncodeBranch("�����_7", AspectsFrom->FieldByName("������")->AsInteger);
  AspectsTo->FieldByName("������")->Value=N;
  N=EncodeBranch("�����_3", AspectsFrom->FieldByName("�����������")->AsInteger);
  AspectsTo->FieldByName("�����������")->Value=N;

  AspectsTo->FieldByName("G")->Value=AspectsFrom->FieldByName("G")->AsInteger;
  AspectsTo->FieldByName("O")->Value=AspectsFrom->FieldByName("O")->AsInteger;
  AspectsTo->FieldByName("R")->Value=AspectsFrom->FieldByName("R")->AsInteger;
  AspectsTo->FieldByName("S")->Value=AspectsFrom->FieldByName("S")->AsInteger;
  AspectsTo->FieldByName("T")->Value=AspectsFrom->FieldByName("T")->AsInteger;
  AspectsTo->FieldByName("L")->Value=AspectsFrom->FieldByName("L")->AsInteger;
  AspectsTo->FieldByName("N")->Value=AspectsFrom->FieldByName("N")->AsInteger;
  AspectsTo->FieldByName("Z")->Value=AspectsFrom->FieldByName("Z")->AsInteger;
  AspectsTo->FieldByName("����������")->Value=AspectsFrom->FieldByName("����������")->AsBoolean;

  N=AspectsFrom->FieldByName("���������� �����������")->AsInteger;
  AspectsTo->FieldByName("���������� �����������")->Value=N;
  N=AspectsFrom->FieldByName("������� �����������")->AsInteger;
  AspectsTo->FieldByName("������� �����������")->Value=N;
  N=AspectsFrom->FieldByName("��������������")->AsInteger;
  AspectsTo->FieldByName("��������������")->Value=N;

  AspectsTo->FieldByName("������������� �����������")->Value=AspectsFrom->FieldByName("������������� �����������")->Value;
  AspectsTo->FieldByName("������������ �����������")->Value=AspectsFrom->FieldByName("������������ �����������")->Value;
  AspectsTo->FieldByName("���������� � ��������")->Value=AspectsFrom->FieldByName("���������� � ��������")->Value;
  AspectsTo->FieldByName("������������ ���������� � ��������")->Value=AspectsFrom->FieldByName("������������ ���������� � ��������")->Value;
  AspectsTo->FieldByName("�����������")->Value=AspectsFrom->FieldByName("�����������")->Value;
  AspectsTo->FieldByName("���� ��������")->Value=AspectsFrom->FieldByName("���� ��������")->Value;
  AspectsTo->FieldByName("������ ��������")->Value=AspectsFrom->FieldByName("������ ��������")->Value;
  AspectsTo->FieldByName("����� ��������")->Value=AspectsFrom->FieldByName("����� ��������")->Value;

  AspectsTo->Post();
 }


}
Table->Edit();
Table->FieldByName("Result")->Value="���������";
Table->Post();
}
catch(...)
{
Table->Edit();
Table->FieldByName("Result")->Value="������";
Table->Post();
}

}
}
catch(...)
{
Edit1->Text="������";

}

//��������� Aspects

Comm->Connection=TempTo;
Comm->CommandText="UPDATE �����_3 SET �����_3.NumCopy = 0;";
Comm->Execute();

Comm->CommandText="UPDATE �����_4 SET �����_4.NumCopy = 0;";
Comm->Execute();

Comm->CommandText="UPDATE �����_5 SET �����_5.NumCopy = 0;";
Comm->Execute();

Comm->CommandText="UPDATE �����_6 SET �����_6.NumCopy = 0;";
Comm->Execute();

Comm->CommandText="UPDATE �����_7 SET �����_7.NumCopy = 0;";
Comm->Execute();

AspectTo->Connected=false;
TempFrom->Connected=false;
TempTo->Connected=false;


ShowMessage("���������");
}
else
{
 ShowMessage("�� ������� ��� ����� ��� ��������������");
}
}
//---------------------------------------------------------------------------
void TForm1::CopyTable(String NameNode, String NameBranch)
{

MP<TADODataSet>FromNode(this);
FromNode->Connection=TempFrom;
FromNode->CommandText="Select * From "+NameNode+" order by [��������], [����� ����];";
FromNode->Active=true;

MP<TADODataSet>FromBranch(this);
FromBranch->Connection=TempFrom;
FromBranch->CommandText="Select * From "+NameBranch;
FromBranch->Active=true;

MP<TADODataSet>ToBranch(this);
ToBranch->Connection=TempTo;
ToBranch->CommandText="Select * From "+NameBranch;
ToBranch->Active=true;

MP<TADODataSet>ToNode(this);
ToNode->Connection=TempTo;
ToNode->CommandText="Select * From "+NameNode;
ToNode->Active=true;

MP<TADOCommand>Comm(this);
Comm->Connection=TempTo;
Comm->CommandText="Delete * from "+NameNode;
Comm->Execute();

Comm->CommandText="Delete * from TempNode";
Comm->Execute();

Comm->CommandText="Delete * from TempBranch";
Comm->Execute();

MP<TADODataSet>TempNode(this);
TempNode->Connection=TempTo;
TempNode->CommandText="Select * From TempNode";
TempNode->Active=true;

MP<TADODataSet>TempBranch(this);
TempBranch->Connection=TempTo;
TempBranch->CommandText="Select * From TempBranch";
TempBranch->Active=true;

for(FromNode->First();!FromNode->Eof;FromNode->Next())
{
 TempNode->Insert();
 TempNode->FieldByName("����� ����")->Value=FromNode->FieldByName("����� ����")->Value;
 TempNode->FieldByName("��������")->Value=FromNode->FieldByName("��������")->Value;
 TempNode->FieldByName("��������")->Value=FromNode->FieldByName("��������")->Value;
 TempNode->Post();
}

for(TempBranch->First();!TempBranch->Eof;TempBranch->Next())
{
 TempBranch->Insert();
 TempBranch->FieldByName("����� �����")->Value=FromNode->FieldByName("����� ����")->Value;
 TempBranch->FieldByName("����� ��������")->Value=FromNode->FieldByName("��������")->Value;
 TempBranch->FieldByName("��������")->Value=FromNode->FieldByName("��������")->Value;
 TempBranch->FieldByName("�����")->Value=FromNode->FieldByName("�����")->Value;
 TempBranch->Post();
}

TempNode->First();
do
{
 int OldNumberNode=TempNode->FieldByName("����� ����")->Value;

 ToNode->Insert();
 ToNode->FieldByName("��������")->Value=TempNode->FieldByName("��������")->Value;
 ToNode->FieldByName("��������")->Value=TempNode->FieldByName("��������")->Value;
 ToNode->Post();

 ToNode->Active=false;
 ToNode->Active=true;
 ToNode->Last();
 int NewNumberNode=ToNode->FieldByName("����� ����")->Value;

 MP<TADODataSet>FromBranch(this);
 FromBranch->Connection=TempFrom;
 FromBranch->CommandText="Select * From "+NameBranch+" Where [����� ��������]="+IntToStr(OldNumberNode);
 FromBranch->Active=true;

 for(FromBranch->First();!FromBranch->Eof;FromBranch->Next())
 {
  ToBranch->Insert();
  ToBranch->FieldByName("����� ��������")->Value=NewNumberNode;
  ToBranch->FieldByName("��������")->Value=FromBranch->FieldByName("��������")->Value;
  ToBranch->FieldByName("�����")->Value=FromBranch->FieldByName("�����")->Value;
  ToBranch->FieldByName("NumCopy")->Value=FromBranch->FieldByName("����� �����")->Value;
  ToBranch->Post();
 }

 MP<TADODataSet>CorrectNode(this);
 CorrectNode->Connection=TempTo;
 CorrectNode->CommandText="Select * from TempNode Where ��������="+IntToStr(OldNumberNode);
 CorrectNode->Active=true;

 for(CorrectNode->First();!CorrectNode->Eof;CorrectNode->Next())
 {
  CorrectNode->Edit();
  CorrectNode->FieldByName("��������")->Value=NewNumberNode;
  CorrectNode->Post();
 }

 Comm->CommandText="Delete * from TempNode Where [����� ����]="+IntToStr(OldNumberNode);
Comm->Execute();

 TempNode->Active=false;
 TempNode->Active=true;
}
while(TempNode->RecordCount!=0);

}
//--------------------------------------------------------------------------
int TForm1::EncodeBranch(String NameBranch, int NumCode)
{
MP<TADODataSet>Branch(this);
Branch->Connection=TempTo;
Branch->CommandText="Select * From "+NameBranch;
Branch->Active=true;

if(NumCode==0)
{
return 0;
}
else
{
Branch->Locate("NumCopy", NumCode, SO);
}
return Branch->FieldByName("����� �����")->AsInteger;
}
//-------------------------------------------------------------------------
