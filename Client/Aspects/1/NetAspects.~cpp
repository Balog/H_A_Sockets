//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//---------------------------------------------------------------------------
USEFORM("Main.cpp", Form1);
USEFORM("PassForm.cpp", Pass);
USEFORM("Progress.cpp", Prog);
USEFORM("Zastavka.cpp", Zast);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {
                 Application->Initialize();
                 Application->CreateForm(__classid(TZast), &Zast);
                 Application->CreateForm(__classid(TForm1), &Form1);
                 Application->CreateForm(__classid(TPass), &Pass);
                 Application->CreateForm(__classid(TProg), &Prog);
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
