//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//---------------------------------------------------------------------------
USEFORM("AdmMain.cpp", Main);
USEFORM("EditLogin.cpp", EditLogins);
USEFORM("Password.cpp", Pass);
USEFORM("Diary.cpp", FDiary);
USEFORM("About.cpp", FAbout);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {
          CreateMutex(NULL, true, "AdminMutex");

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
                 Application->CreateForm(__classid(TMain), &Main);
                 Application->CreateForm(__classid(TEditLogins), &EditLogins);
                 Application->CreateForm(__classid(TPass), &Pass);
                 Application->CreateForm(__classid(TFDiary), &FDiary);
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
