//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//---------------------------------------------------------------------------
USEFORM("Main.cpp", Documents);
USEFORM("PassForm.cpp", Pass);
USEFORM("Progress.cpp", Prog);
USEFORM("Zastavka.cpp", Zast);
USEFORM("MainForm.cpp", Form1);
USEFORM("Rep1.cpp", Report1);
USEFORM("Svod.cpp", FSvod);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {
                 Application->Initialize();
                 Application->CreateForm(__classid(TZast), &Zast);
                 Application->CreateForm(__classid(TDocuments), &Documents);
                 Application->CreateForm(__classid(TPass), &Pass);
                 Application->CreateForm(__classid(TProg), &Prog);
                 Application->CreateForm(__classid(TForm1), &Form1);
                 Application->CreateForm(__classid(TReport1), &Report1);
                 Application->CreateForm(__classid(TFSvod), &FSvod);
                 Application->Run();
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
