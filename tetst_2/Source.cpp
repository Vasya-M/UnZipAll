#include<Windows.h>
#include<iostream>
#include<vector>
#include<string>
#include<Shldisp.h>
#include<atlbase.h>
#include<shlobj.h>
#include<sstream>


using namespace std;

std::string GetProgramPath() {
	using namespace std;
	string pt(_MAX_PATH + 1, '\0');
	GetModuleFileName(NULL, (LPSTR)pt.c_str(), _MAX_PATH);
	return pt;
}
std::string& CutToFolder(std::string& path)
{
	char c;
	int len = path.length();
	while (len && (c = path[len--] != '\\'));
	path[len + 1] = '\0';
	path.resize(len + 1);
	return path;
}

std::string AddFolder(std::string& path, string& fold, bool add_folder = true)
{
	if (add_folder)
		return std::string(path + '\\' + fold);
	else
		return std::string(path + fold);
}
std::string AddFolder(std::string& path, int i)
{
	return AddFolder(path, std::to_string(i));
}

std::string CreateFolder(std::string& path)
{
	if (CreateDirectory(path.c_str(), NULL) ||
		ERROR_ALREADY_EXISTS == GetLastError())
	{
		return path;
	}
	return "\0";
}

void CreateFolderRange(std::string& path, int count)
{
	while (count)
		CreateFolder(AddFolder(path, count--));
}

bool Unzip2Folder(BSTR lpZipFile, BSTR lpFolder)
{
	IShellDispatch *pISD;

	Folder  *pZippedFile = 0L;
	Folder  *pDestination = 0L;

	long FilesCount = 0;
	IDispatch* pItem = 0L;
	FolderItems *pFilesInside = 0L;

	VARIANT Options, OutFolder, InZipFile, Item;
	CoInitialize(NULL);
	__try {
		if (CoCreateInstance(CLSID_Shell, NULL, CLSCTX_INPROC_SERVER, IID_IShellDispatch, (void **)&pISD) != S_OK)
			return 1;

		InZipFile.vt = VT_BSTR;
		InZipFile.bstrVal = lpZipFile;
		pISD->NameSpace(InZipFile, &pZippedFile);
		if (!pZippedFile)
		{
			pISD->Release();
			return 1;
		}

		OutFolder.vt = VT_BSTR;
		OutFolder.bstrVal = lpFolder;
		pISD->NameSpace(OutFolder, &pDestination);
		if (!pDestination)
		{
			pZippedFile->Release();
			pISD->Release();
			return 1;
		}

		pZippedFile->Items(&pFilesInside);
		if (!pFilesInside)
		{
			pDestination->Release();
			pZippedFile->Release();
			pISD->Release();
			return 1;
		}

		pFilesInside->get_Count(&FilesCount);
		if (FilesCount < 1)
		{
			pFilesInside->Release();
			pDestination->Release();
			pZippedFile->Release();
			pISD->Release();
			return 0;
		}

		pFilesInside->QueryInterface(IID_IDispatch, (void**)&pItem);

		Item.vt = VT_DISPATCH;
		Item.pdispVal = pItem;

		Options.vt = VT_I4;
		Options.lVal = 1024 | 512 | 16 | 4;//http://msdn.microsoft.com/en-us/library/bb787866(VS.85).aspx

		bool retval = pDestination->CopyHere(Item, Options) == S_OK;

		pItem->Release(); pItem = 0L;
		pFilesInside->Release(); pFilesInside = 0L;
		pDestination->Release(); pDestination = 0L;
		pZippedFile->Release(); pZippedFile = 0L;
		pISD->Release(); pISD = 0L;

		return retval;

	}
	__finally
	{
		CoUninitialize();
	}
}


vector<string>  GetAllFiles(string folder)
{
	vector<string> names;
	string search_path = folder + "/*.zip";
	WIN32_FIND_DATA fd;
	HANDLE hFind = ::FindFirstFile(search_path.c_str(), &fd);
	if (hFind != INVALID_HANDLE_VALUE) {
		do {
			// read all (real) files in current folder
			// , delete '!' read other 2 default folder . and ..
			if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
				names.push_back(fd.cFileName);
			}
		} while (::FindNextFile(hFind, &fd));
		::FindClose(hFind);
	}
	return names;
}
static int CALLBACK BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData)
{

	if (uMsg == BFFM_INITIALIZED)
	{
		std::string tmp = (const char *)lpData;
		std::cout << "path: " << tmp << std::endl;
		SendMessage(hwnd, BFFM_SETSELECTION, TRUE, lpData);
	}

	return 0;
}

std::string BrowseFolder(std::string saved_path)
{
	TCHAR path[MAX_PATH];

	const char * path_param = saved_path.c_str();

	BROWSEINFO bi = { 0 };
	bi.lpszTitle = ("Browse for folder...");
	bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_BROWSEINCLUDEURLS | BIF_EDITBOX | BIF_STATUSTEXT | BIF_USENEWUI;
	bi.lpfn = BrowseCallbackProc;
	bi.lParam = (LPARAM)path_param;

	LPITEMIDLIST pidl = SHBrowseForFolder(&bi);

	if (pidl != 0)
	{
		//get the name of the folder and put it in path
		SHGetPathFromIDList(pidl, path);

		//free memory used
		IMalloc * imalloc = 0;
		if (SUCCEEDED(SHGetMalloc(&imalloc)))
		{
			imalloc->Free(pidl);
			imalloc->Release();
		}

		return path;
	}

	return "";
}

std::string CutToName(std::string file)
{
	char c;
	int len = file.length();
	while (len && (c = file[len--] != '.'));
	file[len + 1] = '\0';
	file.resize(len + 1);
	return file;
}

int main(int argc, char* argv[]) {


	string path;
	if (argc > 1)
	{
		path = argv[1];
		goto doing;
	}
	path = (GetProgramPath());

	int result = MessageBox(NULL, "Use current folder?\n(No) for selecting another one", "question", MB_YESNO);
	switch (result)
	{
	case IDYES:
		CutToFolder(path);
		break;
	case IDNO:
		int result2 = MessageBox(NULL, "Use console input?\n(No) for using dialog window", "question", MB_YESNO);
		switch (result2)
		{
		case IDYES:
			std::getline(cin, path);
			break;
		case IDNO:
			path = BrowseFolder(path);
			break;
		}
		break;
	}
doing:
	cout << path << endl;
	vector<string> files = GetAllFiles(path);
	for (int i = 0; i < files.size(); i++)
	{
		std::string& _folder = AddFolder(path, CutToName(files[i]));
		std::string& _file = AddFolder(path, files[i]);
		CComBSTR folder(CreateFolder(_folder).c_str());
		CComBSTR file(_file.c_str());
		cout << "\nfile - " << _file << endl;
		cout << "folder - " << _folder << endl;
		cout << "result - " << ((Unzip2Folder(file, folder)) ? "successful" : "fail") << "\n";
	}
	cout << "end\n";
	system("pause");
}
