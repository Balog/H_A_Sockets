//---------------------------------------------------------------------------


#pragma hdrstop
#include <FileCtrl.hpp>
#include "MDBConnector.h"

//---------------------------------------------------------------------------



MDBConnector::MDBConnector(String PatchDB, String NameDB, String PathBackup, bool AutoConnect, bool AutoPackDB, TComponent* Owner): TADOConnection(Owner)
{
if(PatchDB.SubString(PatchDB.Length(),2)=="\\")
{
pPatchDB=PatchDB;
}
else
{
pPatchDB=PatchDB+"\\";
}
pNameDB=NameDB;
pPathBackup=PathBackup;
pAutoPackDB=AutoPackDB;
pOwner=Owner;

//String Path=ExtractFilePath(Application->ExeName);

//this->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source="+PatchDB+"\\"+NameDB+".mdb"+";Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False";
this->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source="+PatchDB+NameDB+".dtb"+";Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False";

this->LoginPrompt=false;

if(!PathBackup.IsEmpty())
{
Backup();

}

if(AutoConnect)
{
this->Connected=true;
}

}
//----------------------------------------------
MDBConnector::MDBConnector(String PatchDB, String NameDB, TComponent* Owner): TADOConnection(Owner)
{
if(PatchDB.SubString(PatchDB.Length(),2)=="\\")
{
pPatchDB=PatchDB;
}
else
{
pPatchDB=PatchDB+"\\";
}
pNameDB=NameDB;
pPathBackup="";
pAutoPackDB="";
pOwner=Owner;

String Path=ExtractFilePath(Application->ExeName);

//this->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source="+PatchDB+"\\"+NameDB+".mdb"+";Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False";
this->ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source="+PatchDB+NameDB+".dtb"+";Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False";
this->LoginPrompt=false;
}
//----------------------------------------------
void MDBConnector::SetPatchBackUp(String Patch)
{
pPathBackup=Patch;
}
//-----------------------------------------------
__fastcall MDBConnector::~MDBConnector()
{
this->Connected=false;
if(pAutoPackDB)
{
PackDB();
}
}
//----------------------------------------------
bool MDBConnector::Backup(String PatchBackup)
{
pPathBackup=PatchBackup;
bool B=Backup();

return B;
}
//----------------------------------------------
bool MDBConnector::Backup(String PatchBackup, String NameBackUp)
{
pPathBackup=PatchBackup;
pNamebackUp=NameBackUp;
PackDB();
Backup();
return FileExists(pPathBackup+pNamebackUp+".rar");
}
//----------------------------------------------
bool MDBConnector::Backup()
{
SetCurrentDir(ExtractFilePath(Application->ExeName));
String CurDir=GetCurrentDir();
if(!DirectoryExists(pPathBackup))
{
 CreateDir(pPathBackup);
}

bool B;
bool Ret=false;
String Name;
if(pPathBackup!=NULL)
{

bool Connected=this->Connected;
DisconnectDB();


if (pNamebackUp.IsEmpty())
{
Name=CreateName2();
}
else
{
Name=pPathBackup+pNamebackUp+".rar";
}

if(FileExists(Name)==false)
{

if(!FileExists(pPatchDB+pNameDB+".ldb"))
{
 _STARTUPINFOA StartInfo = { sizeof(TStartupInfo) };
 _PROCESS_INFORMATION ProcInfo;
 StartInfo.cb = sizeof(StartInfo);
 StartInfo.dwFlags = STARTF_USESHOWWINDOW;
 StartInfo.wShowWindow = SW_HIDE;
 String N="\""+ExtractFilePath(Application->ExeName)+"Resurs.dat\" a -ep \""+Name+" \" \""+pPatchDB+pNameDB+".dtb\"";


 SetCurrentDir(ExtractFilePath(Application->ExeName));

 bool A=CreateProcess(NULL, N.c_str(),
                     NULL, NULL, false,
                     CREATE_NEW_CONSOLE |
                     IDLE_PRIORITY_CLASS,
                     NULL, NULL, &StartInfo,                     &ProcInfo);
 if (!A )
 {
SetCurrentDir(CurDir);
this->Connected=Connected;
Ret=false;
//Application->MessageBoxA("������������� ����������.\n����������� ������ Resurs.dat\n\n��������� ���������� ������", "������������� ���� ������", MB_ICONERROR|MB_OK|MB_TOPMOST);

 }
 else
 {
//Label2->Visible=true;
int N=0;

while(WaitForSingleObject(ProcInfo.hProcess, 1000)== WAIT_TIMEOUT)
{
N++;

if(FArchTimer) OnArchTimer(this, N);
}

}
this->Connected=Connected;
SetCurrentDir(CurDir);
Ret=true;
//return Ret;
}
else
{
Ret=false;
}
}
}
else
{
Ret=false;
}
if(FileExists(Name))
{
Ret=true;
}
else
{
Ret=false;
}


return Ret;
}
//----------------------------------------------
void MDBConnector::ConnectDB()
{
this->Connected=true;
}
//----------------------------------------------
void MDBConnector::DisconnectDB()
{
this->Connected=false;
}
//----------------------------------------------
bool MDBConnector::PackDB()
{
bool Conn=this->Connected;
bool Ok=false;
DisconnectDB();
//String Path=ExtractFilePath(Application->ExeName);
String CurDir=GetCurrentDir();
SetCurrentDir(pPatchDB);
AnsiString BP=pPatchDB+pNameDB+".ldb";

if(FileExists(BP)==false)
{

Variant JE = Variant::CreateObject ("JRO.JetEngine");

WideString WSCurrConn;
WSCurrConn=this->ConnectionString;
try
 {

   String N1=pPatchDB+"temp.mdb";
    DeleteFile(N1);
  JE.OleFunction ("CompactDatabase",  WSCurrConn,
   "Provider=Microsoft.Jet.OLEDB.4.0;Password="";Mode=ReadWrite;Data Source=temp.mdb");
  if(FileExists(N1))
   {
   String N=pPatchDB+pNameDB+".dtb";
    CopyFile(N1.c_str(), N.c_str(), false);
    DeleteFile(N1);
    Ok=true;
   }
 }

catch ( ... )
{
SetCurrentDir(CurDir);
}



}

this->Connected=Conn;
SetCurrentDir(CurDir);
return Ok;
}
//----------------------------------------------
String MDBConnector::CreateName()
{
String Name="";
String Patch;

Word Day, Month, Year;
DecodeDate(Now(), Year, Month, Day);
int D, M, Y;
D=(int)Day;
M=(int)Month;
Y=(int)Year;
int i=1;
do
{
Name=IntToStr(Y)+"_"+IntToStr(M)+"_"+IntToStr(D)+" (� "+IntToStr(i)+").rar";
Patch=pPathBackup+"\\"+pNameDB+" �� "+Name;
//Patch="D:\\"+Name;

i++;

}
while(FileExists(Patch)==true);

return Patch;
}
//----------------------------------------------
String MDBConnector::CreateName2()
{
String Name="";
String Patch;

Word Day, Month, Year;
DecodeDate(Now(), Year, Month, Day);
int D, M, Y;
D=(int)Day;
M=(int)Month;
Y=(int)Year;
/*
int i=1;
do
{
Name=IntToStr(Y)+"_"+IntToStr(M)+"_"+IntToStr(D)+" (� "+IntToStr(i)+").rar";
Patch=pPathBackup+"\\"+pNameDB+" �� "+Name;
//Patch="D:\\"+Name;

i++;

}
while(FileExists(Patch)==true);
*/
Name=IntToStr(Y)+"_"+IntToStr(M)+"_"+IntToStr(D)+".rar";
Patch=pPathBackup+"\\"+pNameDB+" �� "+Name;

return Patch;
}
//----------------------------------------------------
void MDBConnector::ClearArchive(int Age)
{
try
{
if(Age>0)
{
TFileListBox* FLB=new TFileListBox(pOwner);

FLB->Mask="*.rar";
FLB->Visible=false;
FLB->Parent=(TWinControl*)pOwner;
FLB->Directory=ExtractFilePath(Application->ExeName)+pPathBackup+"\\";
for(int i=0;i<FLB->Count;i++)
{
FLB->ItemIndex=i;
String Name=FLB->FileName;
int FAge=FileAge(Name);

if(FAge>=Age)
{
DeleteFile(Name);
}
}
delete FLB;
}
}
catch(...)
{

}
}
//----------------------------------------------------
int MDBConnector::FileAge(String File)
{
FILETIME lpCreationTime, lpLastAccessTime, lpLastWriteTime;
HANDLE H = CreateFile(File.c_str(), 0,
                      0, NULL, OPEN_EXISTING, 0, NULL);
if(H == INVALID_HANDLE_VALUE)
{
 //ShowMessage(SysErrorMessage(GetLastError()));
 String S=SysErrorMessage(GetLastError());
 String S1="��������� ������ ��� ������� ������: "+S;
Application->MessageBoxA(S1.c_str(), "������� ������", MB_ICONEXCLAMATION|MB_OK|MB_TOPMOST);

 return -1;
}
GetFileTime(H, &lpCreationTime, &lpLastAccessTime,
            &lpLastWriteTime);
CloseHandle(H);
SYSTEMTIME ST;
AnsiString S;
//FileTimeToSystemTime(&lpCreationTime, &ST);
FileTimeToSystemTime(&lpLastWriteTime, &ST);

TDateTime DT=SystemTimeToDateTime(ST);
int Age=Now()-DT;
return Age;
}
//-------------------------------------------------------
String MDBConnector::GetDBPatch()
{
return pPatchDB;
}
//-------------------------------------------------------
String MDBConnector::GetDBName()
{
return pNameDB;
}
//-------------------------------------------------------
#pragma package(smart_init)
 