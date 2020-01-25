// a2shtcut.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "a2shtcut.h"


class CoInit
{
public:
	CoInit()
	{
		::CoInitialize(NULL);
	}
	~CoInit()
	{
		::CoUninitialize();
	}
};

// forward
bool CreateLink(LPCWSTR szTarget, LPCWSTR szPath, LPCWSTR szName, LPCWSTR szWDir, int nIcon, LPCWSTR szIconPath, LPCWSTR szDescr, LPCWSTR szArgs);
void RemoveLink(LPCWSTR szTarget, LPCWSTR szName);

int APIENTRY _tWinMain(HINSTANCE hInstance,
	HINSTANCE hPrevInstance,
	LPTSTR    lpCmdLine,
	int       nCmdShow)
{
	CoInit ci;
	enum CMD {
		_cmd_none = 0,
		_cmd_add = 1,
		_cmd_remove = 2,
	};
	CMD cmd = _cmd_none;

	CString str(lpCmdLine);

	CString path;   // link path
	CString target; // save to
	CString name;   // link name
	CString wdir;   // working dir
	int     icon = -1;
	CString idir; // icon directory
	CString descr; // description
	CString args;  // path arguments

	int pos = 0;
	CString token = str.Tokenize(L"/", pos);
	while (pos != -1) {
		token.Trim();
		int len = token.GetLength();
		if (token == L"add") {
			cmd = _cmd_add;
		}
		else if (token == L"remove") {
			cmd = _cmd_remove;
		}
		else if (token.Find(L"target=") == 0) {
			target = token.Right(len - 7);
		}
		else if (token.Find(L"path=") == 0) {
			path = token.Right(len - 5);
		}
		else if (token.Find(L"name=") == 0) {
			name = token.Right(len - 5);
		}
		else if (token.Find(L"wdir=") == 0) {
			wdir = token.Right(len - 5);
		}
		else if (token.Find(L"idir=") == 0) {
			idir = token.Right(len - 5);
		}
		else if (token.Find(L"icon=") == 0) {
			icon = _ttoi(token.Right(len - 5));
		}
		else if (token.Find(L"descr=") == 0) {
			descr = token.Right(len - 6);
		}
		else if (token.Find(L"args=") == 0) {
			args = token.Right(len - 5);
		}
		token = str.Tokenize(L"/", pos);
	}

	if (cmd == _cmd_add) {
		if (!CreateLink(target, path, name, wdir, icon, idir, descr, args)) {
			CString str;
			str.Format(L"Error creating shortcut.\ntarget='%s'\npath='%s'\nname='%s'", (LPCWSTR)target, (LPCWSTR)path, (LPCWSTR)name);
			::MessageBox(NULL, str, L"A2v10", MB_OK | MB_ICONEXCLAMATION);
		}
	}
	else if (cmd == _cmd_remove) {
		RemoveLink(target, name);
	}
	else {
		::MessageBox(nullptr, L"usage a2shtcut -[add!remove] [params]", L"A2v10", MB_OK | MB_ICONEXCLAMATION);
	}

	return 0;
}

bool CreateLink(LPCWSTR szTarget, LPCWSTR szPath, LPCWSTR szName, LPCWSTR szWDir, int nIcon, LPCWSTR szIconPath, LPCWSTR szDescr, LPCWSTR szArgs)
{
	CComPtr<IShellLinkW> pShellLink;
	HRESULT hr = S_OK;
	hr = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
		IID_IShellLinkW,
		reinterpret_cast<void **>(&pShellLink));
	if (FAILED(hr))
		return false;
	CComQIPtr<IPersistFile> pFile = pShellLink;
	if (pFile == NULL)
		return false;

	pShellLink->SetPath(szPath);
	if (szArgs && *szArgs) {
		pShellLink->SetArguments(szArgs);
	}
	if (szIconPath)
		pShellLink->SetIconLocation(szIconPath, 0);
	else if (nIcon != -1)
		pShellLink->SetIconLocation(szPath, nIcon);

	pShellLink->SetWorkingDirectory(szWDir);
	pShellLink->SetDescription(szDescr);

	wchar_t linkName[MAX_PATH + 1] = { 0 };
	wcscpy_s(linkName, MAX_PATH, szName);
	wcscat_s(linkName, MAX_PATH, L".lnk");

	wchar_t path[MAX_PATH + 1];
	wcscpy_s(path, MAX_PATH, szTarget);
	::PathAppend(path, linkName);
	hr = pFile->Save(path, TRUE);
	if (FAILED(hr))
		return false;
	return true;
}

void RemoveLink(LPCWSTR szTarget, LPCWSTR szName)
{

	wchar_t linkName[MAX_PATH + 1] = { 0 };
	wcscpy_s(linkName, MAX_PATH, szName);
	wcscat_s(linkName, MAX_PATH, L".lnk");

	wchar_t path[MAX_PATH + 1];
	wcscpy_s(path, MAX_PATH, szTarget);


	::PathAppend(path, linkName);
	::DeleteFile(path);
}
