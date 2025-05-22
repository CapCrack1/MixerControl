#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#include <mmdeviceapi.h>
#include <audioclient.h>
#include <audiopolicy.h>
#include <endpointvolume.h>
#include <functiondiscoverykeys_devpkey.h>
#include <string>
#include <vector>
#include <sstream>
#include <iostream>
#include <processthreadsapi.h> // Для QueryFullProcessImageNameW

#pragma comment(lib, "ws2_32.lib")
#pragma comment(lib, "ole32.lib")

// COM initialization helper
struct ComInit {
    ComInit() { CoInitialize(nullptr); }
    ~ComInit() { CoUninitialize(); }
};

// Structure to hold audio session info
struct AudioSession {
    std::wstring name;
    float volume;
    DWORD processId;
};

// Функция для получения имени процесса по PID
std::wstring GetProcessName(DWORD pid) {
    std::wstring processName = L"Unknown";
    HANDLE hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
    if (hProcess) {
        WCHAR buffer[MAX_PATH];
        DWORD size = MAX_PATH;
        if (QueryFullProcessImageNameW(hProcess, 0, buffer, &size)) {
            std::wstring fullPath(buffer);
            size_t lastSlash = fullPath.find_last_of(L"\\");
            if (lastSlash != std::wstring::npos) {
                processName = fullPath.substr(lastSlash + 1);
            } else {
                processName = fullPath;
            }
        } else {
            std::wcerr << L"Failed to get process name for PID " << pid << L", error: " << GetLastError() << std::endl;
        }
        CloseHandle(hProcess);
    } else {
        std::wcerr << L"Failed to open process with PID " << pid << L", error: " << GetLastError() << std::endl;
    }
    return processName;
}

// Get audio sessions with names and volumes
std::vector<AudioSession> GetAudioSessions() {
    std::vector<AudioSession> sessions;
    HRESULT hr;

    IMMDeviceEnumerator* pEnumerator = nullptr;
    IMMDevice* pDevice = nullptr;
    IAudioSessionManager2* pSessionManager = nullptr;
    IAudioSessionEnumerator* pSessionEnumerator = nullptr;

    hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
                          __uuidof(IMMDeviceEnumerator), (void**)&pEnumerator);
    if (FAILED(hr)) return sessions;

    hr = pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
    if (FAILED(hr)) { pEnumerator->Release(); return sessions; }

    hr = pDevice->Activate(__uuidof(IAudioSessionManager2), CLSCTX_ALL, nullptr, (void**)&pSessionManager);
    if (FAILED(hr)) { pDevice->Release(); pEnumerator->Release(); return sessions; }

    hr = pSessionManager->GetSessionEnumerator(&pSessionEnumerator);
    if (FAILED(hr)) { pSessionManager->Release(); pDevice->Release(); pEnumerator->Release(); return sessions; }

    int sessionCount;
    pSessionEnumerator->GetCount(&sessionCount);

    for (int i = 0; i < sessionCount; i++) {
        IAudioSessionControl* pSessionControl = nullptr;
        IAudioSessionControl2* pSessionControl2 = nullptr;
        hr = pSessionEnumerator->GetSession(i, &pSessionControl);
        if (FAILED(hr)) continue;

        hr = pSessionControl->QueryInterface(__uuidof(IAudioSessionControl2), (void**)&pSessionControl2);
        if (FAILED(hr)) { pSessionControl->Release(); continue; }

        LPWSTR wszName = nullptr;
        pSessionControl2->GetDisplayName(&wszName);
        std::wstring name = (wszName && wcslen(wszName) > 0) ? wszName : L"";
        if (wszName) CoTaskMemFree(wszName);

        DWORD pid;
        pSessionControl2->GetProcessId(&pid);

        // Если имя пустое или отсутствует, получаем имя процесса по PID
        if (name.empty() || name == L"Unknown") {
            name = GetProcessName(pid);
        }

        ISimpleAudioVolume* pSimpleVolume = nullptr;
        hr = pSessionControl->QueryInterface(__uuidof(ISimpleAudioVolume), (void**)&pSimpleVolume);
        if (SUCCEEDED(hr)) {
            float volume;
            pSimpleVolume->GetMasterVolume(&volume);
            sessions.push_back({name, volume, pid});
            pSimpleVolume->Release();
        }

        pSessionControl2->Release();
        pSessionControl->Release();
    }

    pSessionEnumerator->Release();
    pSessionManager->Release();
    pDevice->Release();
    pEnumerator->Release();
    return sessions;
}

// Set volume for a specific session
bool SetSessionVolume(DWORD pid, float volume) {
    HRESULT hr;
    IMMDeviceEnumerator* pEnumerator = nullptr;
    IMMDevice* pDevice = nullptr;
    IAudioSessionManager2* pSessionManager = nullptr;
    IAudioSessionEnumerator* pSessionEnumerator = nullptr;

    hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
                          __uuidof(IMMDeviceEnumerator), (void**)&pEnumerator);
    if (FAILED(hr)) return false;

    hr = pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
    if (FAILED(hr)) { pEnumerator->Release(); return false; }

    hr = pDevice->Activate(__uuidof(IAudioSessionManager2), CLSCTX_ALL, nullptr, (void**)&pSessionManager);
    if (FAILED(hr)) { pDevice->Release(); pEnumerator->Release(); return false; }

    hr = pSessionManager->GetSessionEnumerator(&pSessionEnumerator);
    if (FAILED(hr)) { pSessionManager->Release(); pDevice->Release(); pEnumerator->Release(); return false; }

    int sessionCount;
    pSessionEnumerator->GetCount(&sessionCount);

    for (int i = 0; i < sessionCount; i++) {
        IAudioSessionControl* pSessionControl = nullptr;
        IAudioSessionControl2* pSessionControl2 = nullptr;
        hr = pSessionEnumerator->GetSession(i, &pSessionControl);
        if (FAILED(hr)) continue;

        hr = pSessionControl->QueryInterface(__uuidof(IAudioSessionControl2), (void**)&pSessionControl2);
        if (FAILED(hr)) { pSessionControl->Release(); continue; }

        DWORD sessionPid;
        pSessionControl2->GetProcessId(&sessionPid);
        if (sessionPid == pid) {
            ISimpleAudioVolume* pSimpleVolume = nullptr;
            hr = pSessionControl->QueryInterface(__uuidof(ISimpleAudioVolume), (void**)&pSimpleVolume);
            if (SUCCEEDED(hr)) {
                hr = pSimpleVolume->SetMasterVolume(volume, nullptr);
                pSimpleVolume->Release();
            }
            pSessionControl2->Release();
            pSessionControl->Release();
            pSessionEnumerator->Release();
            pSessionManager->Release();
            pDevice->Release();
            pEnumerator->Release();
            return SUCCEEDED(hr);
        }
        pSessionControl2->Release();
        pSessionControl->Release();
    }

    pSessionEnumerator->Release();
    pSessionManager->Release();
    pDevice->Release();
    pEnumerator->Release();
    return false;
}


// Convert wstring to string (for HTTP response)
std::string WstringToString(const std::wstring& wstr) {
    std::string str;
    int size = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, nullptr, 0, nullptr, nullptr);
    str.resize(size);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &str[0], size, nullptr, nullptr);
    str.pop_back(); // Remove null terminator
    return str;
}

// Simple HTTP server
void HandleClient(SOCKET clientSocket) {
    char buffer[4096];
    int bytesReceived = recv(clientSocket, buffer, sizeof(buffer) - 1, 0);
    if (bytesReceived <= 0) {
        closesocket(clientSocket);
        return;
    }
    buffer[bytesReceived] = '\0';

    std::string request(buffer);
    if (request.find("GET / ") != std::string::npos || request.find("GET /index.html") != std::string::npos) {
        // Serve HTML interface
        std::string html = R"(
        <!DOCTYPE html>
        <html>
        <head>
            <title>Mixer Control</title>
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                .session { margin: 10px 0; }
                .slider { width: 200px; }
            </style>
        </head>
        <body>
            <h1>Windows Mixer Control</h1>
            <div id="sessions"></div>
            <script>
                async function fetchSessions() {
                    const response = await fetch('/sessions');
                    const text = await response.text();
                    const sessions = text.split('\n').map(line => {
                        const [pid, volume, name] = line.split('|');
                        return { pid: parseInt(pid), volume: parseFloat(volume), name };
                    }).filter(s => s.pid);
                    const sessionsDiv = document.getElementById('sessions');
                    sessionsDiv.innerHTML = '';
                    sessions.forEach(session => {
                        const div = document.createElement('div');
                        div.className = 'session';
                        div.innerHTML = 
                            "<label>" + session.name + " (PID: " + session.pid + ")</label><br>" +
                            "<input type='range' class='slider' min='0' max='1' step='0.01' value='" + session.volume + "'" +
                            " oninput='updateVolume(" + session.pid + ", this.value)'>" +
                            "<span>Volume: <span id='vol-" + session.pid + "'>" + (session.volume * 100).toFixed(0) + "%</span></span>";
                        sessionsDiv.appendChild(div);
                    });
                }

                async function updateVolume(pid, volume) {
                    await fetch("/set_volume?pid=" + pid + "&volume=" + volume, { method: "GET" });
                    document.getElementById("vol-" + pid).innerText = (volume * 100).toFixed(0) + "%";
                }

                fetchSessions();
                setInterval(fetchSessions, 5000);
            </script>
        </body>
        </html>
        )";
        std::string response = "HTTP/1.1 200 OK\r\nContent-Type: text/html\r\nContent-Length: " +
                               std::to_string(html.size()) + "\r\n\r\n" + html;
        send(clientSocket, response.c_str(), response.size(), 0);
    }
    else if (request.find("GET /sessions") != std::string::npos) {
        // Serve session data as plain text
        auto sessions = GetAudioSessions();
        std::stringstream ss;
        for (const auto& session : sessions) {
            ss << session.processId << "|" << session.volume << "|" << WstringToString(session.name) << "\n";
        }
        std::string data = ss.str();
        std::string response = "HTTP/1.1 200 OK\r\nContent-Type: text/plain\r\nContent-Length: " +
                               std::to_string(data.size()) + "\r\n\r\n" + data;
        send(clientSocket, response.c_str(), response.size(), 0);
    }
    else if (request.find("GET /set_volume") != std::string::npos) {
        // Handle volume setting
        size_t pidPos = request.find("pid=");
        size_t volPos = request.find("volume=");
        if (pidPos != std::string::npos && volPos != std::string::npos) {
            std::string pidStr = request.substr(pidPos + 4, request.find('&', pidPos) - pidPos - 4);
            std::string volStr = request.substr(volPos + 7, request.find(' ', volPos) - volPos - 7);
            DWORD pid = std::stoul(pidStr);
            float volume = std::stof(volStr);
            bool success = SetSessionVolume(pid, volume);
            std::string response = "HTTP/1.1 200 OK\r\nContent-Type: text/plain\r\nContent-Length: 2\r\n\r\nOK";
            if (!success) response = "HTTP/1.1 500 Internal Server Error\r\nContent-Type: text/plain\r\nContent-Length: 6\r\n\r\nFailed";
            send(clientSocket, response.c_str(), response.size(), 0);
        }
    }
    else {
        std::string response = "HTTP/1.1 404 Not Found\r\nContent-Length: 0\r\n\r\n";
        send(clientSocket, response.c_str(), response.size(), 0);
    }
    closesocket(clientSocket);
}

int main() {
    ComInit comInit;

    // Initialize Winsock
    WSADATA wsaData;
    if (WSAStartup(MAKEWORD(2, 2), &wsaData) != 0) {
        std::cerr << "WSAStartup failed\n";
        return 1;
    }

    // Create socket
    SOCKET serverSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
    if (serverSocket == INVALID_SOCKET) {
        std::cerr << "Socket creation failed\n";
        WSACleanup();
        return 1;
    }

    // Bind socket
    sockaddr_in serverAddr;
    serverAddr.sin_family = AF_INET;
    serverAddr.sin_addr.s_addr = INADDR_ANY;
    serverAddr.sin_port = htons(8080);
    if (bind(serverSocket, (sockaddr*)&serverAddr, sizeof(serverAddr)) == SOCKET_ERROR) {
        std::cerr << "Bind failed\n";
        closesocket(serverSocket);
        WSACleanup();
        return 1;
    }

    // Listen
    if (listen(serverSocket, SOMAXCONN) == SOCKET_ERROR) {
        std::cerr << "Listen failed\n";
        closesocket(serverSocket);
        WSACleanup();
        return 1;
    }

    std::cout << "Server running at http://0.0.0.0:8080\n";

    // Accept clients
    while (true) {
        SOCKET clientSocket = accept(serverSocket, nullptr, nullptr);
        if (clientSocket == INVALID_SOCKET) continue;
        HandleClient(clientSocket);
    }

    closesocket(serverSocket);
    WSACleanup();
    return 0;
}