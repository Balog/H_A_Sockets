//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//---------------------------------------------------------------------------
USEFORM("Zastavka.cpp", Zast);
USEFORM("Docs.cpp", Documents);
USEFORM("Password.cpp", Pass);
USEFORM("Progress.cpp", FProg);
USEFORM("FMoveAsp.cpp", MAsp);
USEFORM("InpDocs.cpp", InputDocs);
USEFORM("MainForm.cpp", Form1);
USEFORM("InputFiltr.cpp", Filter);
USEFORM("Metod.cpp", Metodika);
USEFORM("F_Vvedenie.cpp", Vvedenie);
USEFORM("Rep1.cpp", Report1);
USEFORM("Svod.cpp", FSvod);
USEFORM("About.cpp", FAbout);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {

   CreateMutex(NULL, true, "NetHazardsMutex");

  int Res = GetLastError();

  if(Res == ERROR_ALREADY_EXISTS)

  {

   ShowMessage("Уже имеется выполняемое приложение");

   Application->Terminate();

  }

  if(Res == ERROR_INVALID_HANDLE)

  {

   ShowMessage("Ошибка в имени mutex object");

   Application->Terminate();

  }
                 if(Res==0)
                 {

                 Application->Initialize();
                 Application->CreateForm(__classid(TZast), &Zast);
                 Application->CreateForm(__classid(TDocuments), &Documents);
                 Application->CreateForm(__classid(TPass), &Pass);
                 Application->CreateForm(__classid(TFProg), &FProg);
                 Application->CreateForm(__classid(TMAsp), &MAsp);
                 Application->CreateForm(__classid(TInputDocs), &InputDocs);
                 Application->CreateForm(__classid(TForm1), &Form1);
                 Application->CreateForm(__classid(TFilter), &Filter);
                 Application->CreateForm(__classid(TMetodika), &Metodika);
                 Application->CreateForm(__classid(TVvedenie), &Vvedenie);
                 Application->CreateForm(__classid(TReport1), &Report1);
                 Application->CreateForm(__classid(TFSvod), &FSvod);
                 Application->CreateForm(__classid(TFAbout), &FAbout);
                 Application->Run();
                }
        }
        catch (Exception &exception)
        {
                 Application->ShowException(&exception);
        }
        catch (...)
        {
                 try
                 {
                         throw Exception("");
                 }
                 catch (Exception &exception)
                 {
                         Application->ShowException(&exception);
                 }
        }
        return 0;
}
//---------------------------------------------------------------------------
