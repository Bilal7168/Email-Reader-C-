#include "stdafx.h"
#include <Windows.h>
#include <string.h>
#include <fstream>
#include <iostream>
#include <tchar.h>

#include "EAGetMailobj.tlh"
using namespace EAGetMailObjLib;
using namespace std;

class Emailreader{
private:
    DWORD  _getCurrentPath(LPTSTR lpPath, DWORD nSize)
    {
        DWORD dwSize = ::GetModuleFileName(NULL, lpPath, nSize);
        if (dwSize == 0 || dwSize == nSize)
        {
            return 0;
        }

        // Change file name to current full path
        LPCTSTR psz = _tcsrchr(lpPath, _T('\\'));
        if (psz != NULL)
        {
            lpPath[psz - lpPath] = _T('\0');
            return _tcslen(lpPath);
        }

        return 0;
    }
    void fileoutforstandardstring(wofstream& _fo, wstring ws) {
        for (int i = 0; ws[i] != '\0'; i++) {
            if (ws[i] >= 0 && ws[i] <= 127) {
                _fo << ws[i];
            }
        }
        _fo << endl;
    }
    BSTR convertstrbstr(string _toread) {
        int index = 0;
        for (int i = 0; _toread[i] != '\0'; i++) {
            index++;
        }
        char* reader = new char[index + 1];
        reader[index] = '\0';
        for (int i = 0; _toread[i] != '\0'; i++) {
            reader[i] = _toread[i];
        }
        return _com_util::ConvertStringToBSTR(reader);
    }
    wstring** subbody;
    int totalemails;
public:
    Emailreader() {
        totalemails = 0;
        subbody = NULL;
    }
    void outputemailstofile(string outputfile, string serverside, string useremail, string userpassword, int _ssport) {
        const int MailServerPop3 = 0;
        const int MailServerImap4 = 1;
        const int MailServerEWS = 2;
        const int MailServerDAV = 3;
        int indexforsubbody = 0;

        // Initialize COM environment
        ::CoInitialize(NULL);

        TCHAR szPath[MAX_PATH + 1];
        _getCurrentPath(szPath, MAX_PATH);

        TCHAR szMailBox[MAX_PATH + 1];
        wsprintf(szMailBox, _T("%s\\inbox"), szPath);

        // Create a folder to store emails
        ::CreateDirectory(szMailBox, NULL);
        int variable = 0;
        wofstream filed(outputfile);

        try
        {
            IMailServerPtr oServer = NULL;
            oServer.CreateInstance(__uuidof(EAGetMailObjLib::MailServer));
            oServer->Server = convertstrbstr(serverside);
            oServer->User = convertstrbstr(useremail);
            oServer->Password = convertstrbstr(userpassword);
            oServer->Protocol = MailServerPop3;

            // Enable SSL/TLS connection, most modern email servers require SSL/TLS by default
            oServer->SSLConnection = VARIANT_TRUE;
            oServer->Port = _ssport;        //for Gmail = 995

            // If your POP3 doesn't deploy SSL connection
            // Please use
            // oServer->SSLConnection = VARIANT_FALSE;
            // oServer->Port = 110;

            IMailClientPtr oClient = NULL;
            oClient.CreateInstance(__uuidof(EAGetMailObjLib::MailClient));
            oClient->LicenseCode = _T("TryIt");

            oClient->Connect(oServer);
            _tprintf(_T("Connected\r\n"));

            IMailInfoCollectionPtr infos = oClient->GetMailInfoList();
            _tprintf(_T("Total %d emails\r\n"), infos->Count);
            totalemails = infos->Count;
            subbody = new wstring * [totalemails];
            //initialising the 2d string store
            for (int i = 0; i < totalemails; i++) {
                subbody[i] = new wstring[2]; //the first has the subject, the second the body
            }
            //the 2d string store end

            for (long i = 0; i < infos->Count; i++)
            {
                IMailInfoPtr pInfo = infos->GetItem(i);

                TCHAR szFile[MAX_PATH + 1];
                // Generate a random file name by current local datetime,
                // You can use your method to generate the filename if you do not like it
                SYSTEMTIME curtm;
                ::GetLocalTime(&curtm);
                ::wsprintf(szFile, _T("%s\\%04d%02d%02d%02d%02d%02d%03d%d.eml"),
                    szMailBox,
                    curtm.wYear,
                    curtm.wMonth,
                    curtm.wDay,
                    curtm.wHour,
                    curtm.wMinute,
                    curtm.wSecond,
                    curtm.wMilliseconds,
                    i);

                // Receive email from POP3 server
                IMailPtr oMail = oClient->GetMail(pInfo);
                BSTR* point = new BSTR[1];
                oMail->get_Subject(point);
                SysFreeString(*point);
                std::wstring str(*point, SysStringLen(*point));
                fileoutforstandardstring(filed, str); //outputs the Subject
                subbody[indexforsubbody][0] = str;
                filed << "-----------------------------------------------------------" << endl;
                BSTR* pointtwo = new BSTR[1];
                oMail->get_TextBody(pointtwo);
                SysFreeString(*pointtwo);
                std::wstring str2(*pointtwo, SysStringLen(*pointtwo));
                fileoutforstandardstring(filed, str2);
                subbody[indexforsubbody][1] = str;
                indexforsubbody++;
                filed << "___________________________________________________________" << endl;
                filed << "___________________________________________________________" << endl;

                // Mark email as deleted from POP3 server.
               // oClient->Delete(pInfo);
            }

            // Delete method just mark the email as deleted,
            // Quit method expunge the emails from server exactly.
            oClient->Quit();
        }
        catch (_com_error& ep)
        {
            _tprintf(_T("Error: %s"), (const TCHAR*)ep.Description());
        }
    }

};

int main() {

    Emailreader er;
    er.outputemailstofile("D:\\filedefault.txt", "pop.gmail.com", "emailname", "emailpassword", 995);


    return 0;
}