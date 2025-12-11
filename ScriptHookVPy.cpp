#include <windows.h>
#include <wininet.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <Python.h>
#include <pybind11/pybind11.h>
#include <pybind11/embed.h>
#include <filesystem>
#include <vector>
#include <fstream>
#include "../inc/ScriptHookV.h"
#include "../inc/natives.h"

#pragma comment(lib, "wininet.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")

namespace py = pybind11;
namespace fs = std::filesystem;

std::vector<py::module_> loadedScripts;
std::ofstream logFile;
HMODULE pythonDll = nullptr;
HMODULE g_hModule = nullptr;
bool pythonInitialized = false;

const wchar_t* PYTHON_URL = L"https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip";
const wchar_t* PYTHON_ZIP_NAME = L"python-3.11.9-embed-amd64.zip";

const wchar_t* REQUIRED_FILES[] = {
    L"python311.dll",
    L"python311.zip",
    L"python3.dll",
    L"vcruntime140.dll",
    L"vcruntime140_1.dll",
    L"libffi-8.dll"
};
const int REQUIRED_FILES_COUNT = 6;

void Log(const std::string& msg)
{
    if (logFile.is_open())
    {
        logFile << msg << std::endl;
        logFile.flush();
    }
}

std::wstring GetModuleDirectory(HMODULE hModule)
{
    wchar_t path[MAX_PATH];
    GetModuleFileNameW(hModule, path, MAX_PATH);
    std::wstring dir(path);
    size_t pos = dir.find_last_of(L"\\/");
    if (pos != std::wstring::npos)
        dir = dir.substr(0, pos);
    return dir;
}

std::string WideToNarrow(const std::wstring& wide)
{
    char buffer[MAX_PATH * 2];
    WideCharToMultiByte(CP_UTF8, 0, wide.c_str(), -1, buffer, sizeof(buffer), nullptr, nullptr);
    return std::string(buffer);
}

bool FileExists(const std::wstring& path)
{
    return GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES;
}

bool CheckRequiredFiles(const std::wstring& baseDir)
{
    for (int i = 0; i < REQUIRED_FILES_COUNT; i++)
    {
        std::wstring filePath = baseDir + L"\\" + REQUIRED_FILES[i];
        if (!FileExists(filePath))
        {
            Log("Missing: " + WideToNarrow(REQUIRED_FILES[i]));
            return false;
        }
    }
    return true;
}

bool DownloadFile(const wchar_t* url, const std::wstring& destPath)
{
    Log("Downloading: " + WideToNarrow(destPath));

    HINTERNET hInternet = InternetOpenW(L"ScriptHookVPy", INTERNET_OPEN_TYPE_PRECONFIG, NULL, NULL, 0);
    if (!hInternet)
    {
        Log("InternetOpen failed: " + std::to_string(GetLastError()));
        return false;
    }

    HINTERNET hUrl = InternetOpenUrlW(hInternet, url, NULL, 0,
        INTERNET_FLAG_RELOAD | INTERNET_FLAG_SECURE | INTERNET_FLAG_NO_CACHE_WRITE, 0);
    if (!hUrl)
    {
        Log("InternetOpenUrl failed: " + std::to_string(GetLastError()));
        InternetCloseHandle(hInternet);
        return false;
    }

    HANDLE hFile = CreateFileW(destPath.c_str(), GENERIC_WRITE, 0, NULL,
        CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE)
    {
        Log("CreateFile failed: " + std::to_string(GetLastError()));
        InternetCloseHandle(hUrl);
        InternetCloseHandle(hInternet);
        return false;
    }

    char buffer[8192];
    DWORD bytesRead;
    DWORD totalBytes = 0;

    while (InternetReadFile(hUrl, buffer, sizeof(buffer), &bytesRead) && bytesRead > 0)
    {
        DWORD bytesWritten;
        WriteFile(hFile, buffer, bytesRead, &bytesWritten, NULL);
        totalBytes += bytesRead;
    }

    CloseHandle(hFile);
    InternetCloseHandle(hUrl);
    InternetCloseHandle(hInternet);

    Log("Downloaded " + std::to_string(totalBytes) + " bytes");
    return totalBytes > 0;
}

bool ExtractZip(const std::wstring& zipPath, const std::wstring& destDir)
{
    Log("Extracting: " + WideToNarrow(zipPath));

    CoInitialize(NULL);

    IShellDispatch* pShell = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_Shell, NULL, CLSCTX_INPROC_SERVER,
        IID_IShellDispatch, (void**)&pShell);

    if (FAILED(hr) || !pShell)
    {
        Log("CoCreateInstance failed");
        CoUninitialize();
        return false;
    }

    VARIANT vZip, vDest;
    VariantInit(&vZip);
    VariantInit(&vDest);

    vZip.vt = VT_BSTR;
    vZip.bstrVal = SysAllocString(zipPath.c_str());

    vDest.vt = VT_BSTR;
    vDest.bstrVal = SysAllocString(destDir.c_str());

    Folder* pZipFolder = nullptr;
    Folder* pDestFolder = nullptr;

    hr = pShell->NameSpace(vZip, &pZipFolder);
    if (FAILED(hr) || !pZipFolder)
    {
        Log("Failed to open zip");
        pShell->Release();
        VariantClear(&vZip);
        VariantClear(&vDest);
        CoUninitialize();
        return false;
    }

    hr = pShell->NameSpace(vDest, &pDestFolder);
    if (FAILED(hr) || !pDestFolder)
    {
        Log("Failed to open dest folder");
        pZipFolder->Release();
        pShell->Release();
        VariantClear(&vZip);
        VariantClear(&vDest);
        CoUninitialize();
        return false;
    }

    FolderItems* pItems = nullptr;
    pZipFolder->Items(&pItems);

    if (pItems)
    {
        VARIANT vItems;
        VariantInit(&vItems);
        vItems.vt = VT_DISPATCH;
        vItems.pdispVal = pItems;

        VARIANT vOptions;
        VariantInit(&vOptions);
        vOptions.vt = VT_I4;
        vOptions.lVal = 0x614;

        hr = pDestFolder->CopyHere(vItems, vOptions);

        Sleep(5000);

        pItems->Release();
    }

    pZipFolder->Release();
    pDestFolder->Release();
    pShell->Release();
    VariantClear(&vZip);
    VariantClear(&vDest);
    CoUninitialize();

    Log("Extraction complete");
    return SUCCEEDED(hr);
}

bool EnsurePythonExists(const std::wstring& baseDir)
{
    if (CheckRequiredFiles(baseDir))
    {
        Log("All Python files present");
        return true;
    }

    std::wstring zipPath = baseDir + L"\\" + PYTHON_ZIP_NAME;

    if (!FileExists(zipPath))
    {
        Log("Python zip not found, downloading...");

        if (!DownloadFile(PYTHON_URL, zipPath))
        {
            Log("Download failed");
            return false;
        }
    }
    else
    {
        Log("Python zip already downloaded");
    }

    if (!ExtractZip(zipPath, baseDir))
    {
        Log("Extraction failed");
        return false;
    }

    for (int i = 0; i < 20; i++)
    {
        if (CheckRequiredFiles(baseDir))
        {
            Log("Python installed successfully");
            return true;
        }
        Sleep(500);
    }

    Log("Some Python files still missing after extraction");
    return false;
}

bool LoadPythonDll(const std::wstring& baseDir)
{
    SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS);
    AddDllDirectory(baseDir.c_str());
    SetDllDirectoryW(baseDir.c_str());

    std::wstring vcruntime = baseDir + L"\\vcruntime140.dll";
    if (FileExists(vcruntime))
    {
        LoadLibraryW(vcruntime.c_str());
    }

    std::wstring vcruntime1 = baseDir + L"\\vcruntime140_1.dll";
    if (FileExists(vcruntime1))
    {
        LoadLibraryW(vcruntime1.c_str());
    }

    std::wstring fullPath = baseDir + L"\\python311.dll";

    pythonDll = LoadLibraryExW(fullPath.c_str(), NULL, LOAD_WITH_ALTERED_SEARCH_PATH);

    if (!pythonDll)
    {
        Log("LoadLibrary failed: " + std::to_string(GetLastError()));
        return false;
    }

    Log("python311.dll loaded");
    return true;
}

bool SetupPythonEnvironment(const std::wstring& baseDir)
{
    PyConfig config;
    PyConfig_InitIsolatedConfig(&config);

    static std::wstring pythonHome = baseDir;
    static std::wstring pythonPath = baseDir + L"\\python311.zip;" +
        baseDir + L";" +
        baseDir + L"\\pythons";

    PyStatus status;

    status = PyConfig_SetString(&config, &config.home, pythonHome.c_str());
    if (PyStatus_Exception(status))
    {
        Log("PyConfig_SetString home failed");
        PyConfig_Clear(&config);
        return false;
    }

    status = PyConfig_SetString(&config, &config.pythonpath_env, pythonPath.c_str());
    if (PyStatus_Exception(status))
    {
        Log("PyConfig_SetString pythonpath failed");
        PyConfig_Clear(&config);
        return false;
    }

    config.site_import = 0;

    status = Py_InitializeFromConfig(&config);
    PyConfig_Clear(&config);

    if (PyStatus_Exception(status))
    {
        Log("Py_InitializeFromConfig failed");
        return false;
    }

    if (!Py_IsInitialized())
    {
        Log("Python not initialized");
        return false;
    }

    pythonInitialized = true;
    Log("Python initialized: " + std::string(Py_GetVersion()));

    return true;
}

void LoadPythonScripts(const std::string& scriptsPath)
{
    fs::path scriptsDir(scriptsPath);

    if (!fs::exists(scriptsDir))
    {
        fs::create_directory(scriptsDir);
        Log("Created pythons directory");
        return;
    }

    for (const auto& entry : fs::directory_iterator(scriptsDir))
    {
        if (entry.path().extension() == ".py")
        {
            std::string fileName = entry.path().filename().string();
            std::string moduleName = entry.path().stem().string();

            try
            {
                py::module_ script = py::module_::import(moduleName.c_str());
                loadedScripts.push_back(script);
                Log(fileName + " loaded");
            }
            catch (const py::error_already_set& e)
            {
                Log(fileName + " error: " + std::string(e.what()));
            }
        }
    }

    Log("Scripts loaded: " + std::to_string(loadedScripts.size()));
}

void RunScripts()
{
    for (auto& script : loadedScripts)
    {
        try
        {
            if (py::hasattr(script, "update"))
            {
                script.attr("update")();
            }
        }
        catch (const py::error_already_set& e)
        {
            Log("Runtime error: " + std::string(e.what()));
        }
    }
}

void Cleanup()
{
    loadedScripts.clear();

    if (pythonInitialized)
    {
        Py_Finalize();
        pythonInitialized = false;
    }
}

void ScriptMain()
{
    std::wstring baseDir = GetModuleDirectory(g_hModule);
    std::string baseDirNarrow = WideToNarrow(baseDir);

    std::string logPath = baseDirNarrow + "\\ScriptHookVPy.log";
    logFile.open(logPath);
    Log("=== ScriptHookVPy Starting ===");
    Log("Base directory: " + baseDirNarrow);

    if (!EnsurePythonExists(baseDir))
    {
        Log("FATAL: Cannot get Python");
        return;
    }

    if (!LoadPythonDll(baseDir))
    {
        Log("FATAL: Cannot load python311.dll");
        return;
    }

    if (!SetupPythonEnvironment(baseDir))
    {
        Log("FATAL: Cannot initialize Python");
        return;
    }

    try
    {
        {
            py::gil_scoped_acquire acquire;

            py::module_ sys = py::module_::import("sys");
            std::string scriptsPath = baseDirNarrow + "\\pythons";
            sys.attr("path").attr("insert")(0, scriptsPath);

            LoadPythonScripts(scriptsPath);
        }

        while (true)
        {
            {
                py::gil_scoped_acquire acquire;
                RunScripts();
            }
            WAIT(0);
        }
    }
    catch (const py::error_already_set& e)
    {
        Log("Python exception: " + std::string(e.what()));
    }
    catch (const std::exception& e)
    {
        Log("C++ exception: " + std::string(e.what()));
    }
    catch (...)
    {
        Log("Unknown exception");
    }

    Cleanup();
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD reason, LPVOID lpReserved)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_hModule = hModule;
        DisableThreadLibraryCalls(hModule);
        scriptRegister(hModule, ScriptMain);
    }
    else if (reason == DLL_PROCESS_DETACH)
    {
        scriptUnregister(hModule);
        Cleanup();
        if (logFile.is_open()) logFile.close();
        if (pythonDll) FreeLibrary(pythonDll);
    }
    return TRUE;
}