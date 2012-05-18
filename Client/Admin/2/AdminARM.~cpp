//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//---------------------------------------------------------------------------
USEFORM("Zastavka.cpp", Zast);
USEFORM("Main.cpp", Form1);
USEFORM("PassForm.cpp", Pass);
USEFORM("EditLogin.cpp", EditLogins);
USEFORM("Diary.cpp", FDiary);
USEFORM("Progress.cpp", Prog);
USEFORM("About.cpp", FAbout);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {
                 Application->Initialize();
                 Application->CreateForm(__classid(TZast), &Zast);
                 Application->CreateForm(__classid(TForm1), &Form1);
                 Application->CreateForm(__classid(TPass), &Pass);
                 Application->CreateForm(__classid(TEditLogins), &EditLogins);
                 Application->CreateForm(__classid(TFDiary), &FDiary);
                 Application->CreateForm(__classid(TProg), &Prog);
                 Application->CreateForm(__classid(TFAbout), &FAbout);
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
