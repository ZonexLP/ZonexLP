#include <windows.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <audiopolicy.h>
#include <psapi.h>
#include <iostream>
#include <thread>
#include <vector>
#include <string>
#include <mutex>
#include <queue>
#include <condition_variable>
#include <algorithm>
#include <chrono>

using namespace std;

// Globale Variablen
IAudioEndpointVolume* pMasterVolume = NULL; // Mastervolume-Steuerung
HANDLE hSerial;                             // Serielle Schnittstelle
vector<wstring> audioPrograms;              // Programme mit Audioausgabe
vector<ISimpleAudioVolume*> audioVolumes;   // LautstÑrke-Schnittstellen der Programme
int currentProgramIndex = -1;               // Index des aktuell ausgewÑhlten Programms (-1 fÅr Mastervolume)
bool stopProcessing = false;                // Steuerung fÅr Threads
float lastVolume = -1.0f;                   // Letzter LautstÑrkewert

// Synchronisation
mutex programMutex, consoleMutex, volumeMutex;

// Funktionsprototypen
void setupSerial(const wchar_t* portName);
void initializeAudioEndpoint();
void setVolume(float volume);
void sendInitialVolumeToArduino();
void processSerialData();
void fetchAudioPrograms();
void switchProgram();
void handleMediaControl(const string& command);
void debugAudioSessions();
string wstringToString(const wstring& wstr);

// Makros fÅr synchronisierte Konsolenausgabe
#define LOCKED_COUT(x) { lock_guard<mutex> lock(consoleMutex); cout << x << endl; }
#define LOCKED_WCOUT(x) { lock_guard<mutex> lock(consoleMutex); wcout << x << endl; }

// Funktion: Unicode-String zu ASCII-String konvertieren
string wstringToString(const wstring& wstr) {
    int len;
    int wlen = static_cast<int>(wstr.size());
    len = WideCharToMultiByte(CP_ACP, 0, wstr.c_str(), wlen, NULL, 0, NULL, NULL);
    string str(len, '\0');
    WideCharToMultiByte(CP_ACP, 0, wstr.c_str(), wlen, &str[0], len, NULL, NULL);
    return str;
}

// Funktion: Serielle Schnittstelle einrichten
void setupSerial(const wchar_t* portName) {
    hSerial = CreateFileW(portName, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hSerial == INVALID_HANDLE_VALUE) {
        LOCKED_COUT("ERROR: Could not open serial port: " << GetLastError());
        exit(1);
    }

    DCB dcbSerialParams = { 0 };
    dcbSerialParams.DCBlength = sizeof(dcbSerialParams);
    if (!GetCommState(hSerial, &dcbSerialParams)) {
        LOCKED_COUT("ERROR: Could not get serial port state.");
        CloseHandle(hSerial);
        exit(1);
    }

    dcbSerialParams.BaudRate = CBR_115200;
    dcbSerialParams.ByteSize = 8;
    dcbSerialParams.StopBits = ONESTOPBIT;
    dcbSerialParams.Parity = NOPARITY;

    if (!SetCommState(hSerial, &dcbSerialParams)) {
        LOCKED_COUT("ERROR: Could not set serial port state.");
        CloseHandle(hSerial);
        exit(1);
    }

    COMMTIMEOUTS timeouts = { 0 };
    timeouts.ReadIntervalTimeout = 10;
    timeouts.ReadTotalTimeoutConstant = 10;
    timeouts.ReadTotalTimeoutMultiplier = 1;

    if (!SetCommTimeouts(hSerial, &timeouts)) {
        LOCKED_COUT("ERROR: Could not set serial port timeouts.");
        CloseHandle(hSerial);
        exit(1);
    }

    LOCKED_COUT("DEBUG: Serial port successfully configured.");
}

// Funktion: Mastervolume initialisieren
void initializeAudioEndpoint() {
    sendInitialVolumeToArduino(); // Sende initiale LautstÑrke beim Start
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        LOCKED_COUT("ERROR: CoInitialize failed. HRESULT: " << hr);
        exit(1);
    }

    IMMDeviceEnumerator* pEnumerator = NULL;
    hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_INPROC_SERVER, __uuidof(IMMDeviceEnumerator), (void**)&pEnumerator);
    if (FAILED(hr)) {
        LOCKED_COUT("ERROR: CoCreateInstance for IMMDeviceEnumerator failed. HRESULT: " << hr);
        exit(1);
    }

    IMMDevice* pDevice = NULL;
    hr = pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
    if (FAILED(hr)) {
        LOCKED_COUT("ERROR: GetDefaultAudioEndpoint failed. HRESULT: " << hr);
        exit(1);
    }

    hr = pDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_INPROC_SERVER, NULL, (void**)&pMasterVolume);
    if (FAILED(hr)) {
        LOCKED_COUT("ERROR: Activate IAudioEndpointVolume failed. HRESULT: " << hr);
        exit(1);
    }

    LOCKED_COUT("DEBUG: Audio endpoint successfully initialized.");
    sendInitialVolumeToArduino();
}

void handleMediaControl(const string& command) {
    INPUT input = { 0 };
    input.type = INPUT_KEYBOARD;

    if (command == "PLAY_PAUSE") {
        input.ki.wVk = VK_MEDIA_PLAY_PAUSE;
    } else if (command == "NEXT_TRACK") {
        input.ki.wVk = VK_MEDIA_NEXT_TRACK;
    } else if (command == "PREV_TRACK") {
        input.ki.wVk = VK_MEDIA_PREV_TRACK;
    } else {
        return;
    }

    SendInput(1, &input, sizeof(INPUT));
    input.ki.dwFlags = KEYEVENTF_KEYUP; // Taste loslassen
    SendInput(1, &input, sizeof(INPUT));
    LOCKED_COUT("DEBUG: Media control command executed: " << command);
}


void sendInitialVolumeToArduino() {
    if (currentProgramIndex == -1) { // Mastervolume
        if (pMasterVolume) {
            float currentVolume = 0.0f;
            HRESULT hr = pMasterVolume->GetMasterVolumeLevelScalar(&currentVolume);
            if (SUCCEEDED(hr)) {
                int scaledVolume = static_cast<int>(currentVolume * 255); // Skaliert auf 0?255
                string command = "SET_VOLUME:" + to_string(scaledVolume) + "\n";
                DWORD bytesWritten;
                WriteFile(hSerial, command.c_str(), command.size(), &bytesWritten, NULL);
                LOCKED_COUT("DEBUG: Sent initial Master volume to Arduino: " << scaledVolume);
            } else {
                LOCKED_COUT("ERROR: Failed to get Master volume.");
            }
        }
    } else { // ProgrammlautstÑrke
        ISimpleAudioVolume* pVolume = audioVolumes[currentProgramIndex];
        if (pVolume) {
            float programVolume = 0.0f;
            HRESULT hr = pVolume->GetMasterVolume(&programVolume);
            if (SUCCEEDED(hr)) {
                int scaledVolume = static_cast<int>(programVolume * 255); // Skaliert auf 0?255
                string command = "SET_VOLUME:" + to_string(scaledVolume) + "\n";
                DWORD bytesWritten;
                WriteFile(hSerial, command.c_str(), command.size(), &bytesWritten, NULL);
                LOCKED_COUT("DEBUG: Sent initial program volume to Arduino: " << scaledVolume);
            } else {
                LOCKED_COUT("ERROR: Failed to get program volume.");
            }
        }
    }
}

// Funktion: LautstÑrke setzen
void setVolume(float volume) {
    volume = max(0.0f, min(1.0f, volume)); // Normalisieren auf [0,1]
    if (volume == lastVolume) return; // Keine énderung notwendig
    lastVolume = volume;

    if (currentProgramIndex == -1) { // Mastervolume
        if (pMasterVolume) {
            HRESULT hr = pMasterVolume->SetMasterVolumeLevelScalar(volume, NULL);
            if (SUCCEEDED(hr)) {
                LOCKED_COUT("DEBUG: Master volume set to " << volume);
            } else {
                LOCKED_COUT("ERROR: Failed to set master volume. HRESULT: " << hr);
            }
        } else {
            LOCKED_COUT("ERROR: Master volume control is not initialized.");
        }
    } else { // Programm-LautstÑrke
        ISimpleAudioVolume* pVolume = nullptr;
        {
            lock_guard<mutex> lock(programMutex);
            if (currentProgramIndex >= 0 && currentProgramIndex < audioVolumes.size()) {
                pVolume = audioVolumes[currentProgramIndex];
            }
        }
        if (pVolume) {
            HRESULT hr = pVolume->SetMasterVolume(volume, NULL);
            if (SUCCEEDED(hr)) {
                LOCKED_COUT("DEBUG: Program volume set to " << volume);
            } else {
                LOCKED_COUT("ERROR: Failed to set program volume. HRESULT: " << hr);
            }
        } else {
            LOCKED_COUT("ERROR: No valid program volume interface found.");
        }
    }
}

// Funktion: Serielle Daten verarbeiten
void processSerialData() {
    DWORD bytesRead;
    char buffer[256];
    while (!stopProcessing) {
        if (ReadFile(hSerial, buffer, sizeof(buffer) - 1, &bytesRead, NULL)) {
            buffer[bytesRead] = '\0'; // Null-terminiert
            string data(buffer);

            // Entfernen von Leerzeichen, \n oder \r
            data.erase(remove(data.begin(), data.end(), '\r'), data.end());
            data.erase(remove(data.begin(), data.end(), '\n'), data.end());

            LOCKED_COUT("DEBUG: Serial data received: " << data);

            if (data.rfind("Potentiometer-Wert:", 0) == 0) { // PrÅfen, ob die Nachricht so beginnt
                int rawValue = stoi(data.substr(19)); // Ab Position 19 lesen
                float volume = rawValue / 255.0f; // Skaliere auf [0.0, 1.0]
                setVolume(volume);
            } else if (data == "SWITCH_PROGRAM") {
                switchProgram();
            } else if (data == "PLAY_PAUSE") {
                handleMediaControl("PLAY_PAUSE");
            } else if (data == "NEXT_TRACK") {
                handleMediaControl("NEXT_TRACK");
            } else if (data == "PREV_TRACK") {
                handleMediaControl("PREV_TRACK");
            } else {
                LOCKED_COUT("DEBUG: Unknown serial command received: " << data);
            }
        }
    }
}


void switchProgram() {
    lock_guard<mutex> lock(programMutex);

    // Rotation inklusive Mastervolume (-1)
    if (currentProgramIndex == -1) {
        currentProgramIndex = 0; // Wechsle zum ersten Programm
    } else if (currentProgramIndex < static_cast<int>(audioPrograms.size()) - 1) {
        currentProgramIndex++; // NÑchstes Programm
    } else {
        currentProgramIndex = -1; // ZurÅck zum Mastervolume
    }

    // Aktionen basierend auf dem aktuellen Index
    if (currentProgramIndex == -1) {
        LOCKED_COUT("DEBUG: Switched to Master Volume.");
        sendInitialVolumeToArduino(); // Mastervolume an Arduino senden
    } else if (currentProgramIndex >= 0 && currentProgramIndex < audioPrograms.size()) {
        LOCKED_WCOUT(L"DEBUG: Switched to program: " << audioPrograms[currentProgramIndex]);

        // LautstÑrke des Programms lesen und senden
        ISimpleAudioVolume* pVolume = audioVolumes[currentProgramIndex];
        if (pVolume) {
            float programVolume = 0.0f;
            HRESULT hr = pVolume->GetMasterVolume(&programVolume);
            if (SUCCEEDED(hr)) {
                int scaledVolume = static_cast<int>(programVolume * 255); // Skaliert auf 0?255
                string command = "SET_VOLUME:" + to_string(scaledVolume) + "\n";
                DWORD bytesWritten;
                WriteFile(hSerial, command.c_str(), command.size(), &bytesWritten, NULL);
                LOCKED_COUT("DEBUG: Sent program volume to Arduino: " << scaledVolume);
            } else {
                LOCKED_COUT("ERROR: Failed to get program volume.");
            }
        }
    }
}



void fetchAudioPrograms() {
    vector<wstring> newPrograms;
    vector<ISimpleAudioVolume*> newVolumes;

    CoInitialize(NULL);

    IMMDeviceEnumerator* pEnumerator = NULL;
    IMMDevice* pDevice = NULL;
    IAudioSessionManager2* pSessionManager = NULL;
    IAudioSessionEnumerator* pSessionList = NULL;

    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_INPROC_SERVER, __uuidof(IMMDeviceEnumerator), (void**)&pEnumerator);
    if (FAILED(hr)) {
        LOCKED_COUT("ERROR: Failed to create MMDeviceEnumerator. HRESULT: " << hr);
        CoUninitialize();
        return;
    }

    hr = pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
    if (FAILED(hr)) {
        LOCKED_COUT("ERROR: Failed to get default audio endpoint. HRESULT: " << hr);
        pEnumerator->Release();
        CoUninitialize();
        return;
    }

    hr = pDevice->Activate(__uuidof(IAudioSessionManager2), CLSCTX_ALL, NULL, (void**)&pSessionManager);
    if (FAILED(hr)) {
        LOCKED_COUT("ERROR: Failed to activate audio session manager. HRESULT: " << hr);
        pDevice->Release();
        pEnumerator->Release();
        CoUninitialize();
        return;
    }

    hr = pSessionManager->GetSessionEnumerator(&pSessionList);
    if (FAILED(hr)) {
        LOCKED_COUT("ERROR: Failed to get audio session enumerator. HRESULT: " << hr);
        pSessionManager->Release();
        pDevice->Release();
        pEnumerator->Release();
        CoUninitialize();
        return;
    }

    int sessionCount = 0;
    pSessionList->GetCount(&sessionCount);

    for (int i = 0; i < sessionCount; i++) {
        IAudioSessionControl* pSessionControl = NULL;
        hr = pSessionList->GetSession(i, &pSessionControl);
        if (FAILED(hr)) continue;

        IAudioSessionControl2* pSessionControl2 = NULL;
        hr = pSessionControl->QueryInterface(__uuidof(IAudioSessionControl2), (void**)&pSessionControl2);
        if (SUCCEEDED(hr) && pSessionControl2) {
            DWORD processId = 0;
            pSessionControl2->GetProcessId(&processId);
            HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processId);
            if (hProcess) {
                wchar_t processName[MAX_PATH];
                if (GetModuleFileNameExW(hProcess, NULL, processName, MAX_PATH)) {
                    ISimpleAudioVolume* pVolume = NULL;
                    hr = pSessionControl->QueryInterface(__uuidof(ISimpleAudioVolume), (void**)&pVolume);
                    if (SUCCEEDED(hr)) {
                        newPrograms.push_back(wstring(processName));
                        newVolumes.push_back(pVolume);
                    }
                }
                CloseHandle(hProcess);
            }
            pSessionControl2->Release();
        }
        pSessionControl->Release();
    }

    {
        lock_guard<mutex> lock(programMutex);
        audioPrograms = newPrograms;
        audioVolumes = newVolumes;
    }

    pSessionList->Release();
    pSessionManager->Release();
    pDevice->Release();
    pEnumerator->Release();
    CoUninitialize();
}

void debugAudioSessions() {
    lock_guard<mutex> lock(programMutex);
    LOCKED_COUT("DEBUG: Current audio sessions:");

    for (size_t i = 0; i < audioPrograms.size(); ++i) {
        LOCKED_WCOUT(L"  [" << i << L"]: " << audioPrograms[i]);
    }

    if (currentProgramIndex == -1) {
        LOCKED_COUT("DEBUG: Currently controlling Master Volume.");
    } else if (currentProgramIndex >= 0 && currentProgramIndex < audioPrograms.size()) {
        LOCKED_WCOUT(L"DEBUG: Currently controlling program: " << audioPrograms[currentProgramIndex]);
    }
}

// Hauptprogramm
int main() {
    setupSerial(L"COM5");
    initializeAudioEndpoint();

    thread serialThread(processSerialData);
    thread fetchThread([] {
        while (!stopProcessing) {
            fetchAudioPrograms();
            this_thread::sleep_for(chrono::seconds(5));
        }
    });

    while (!stopProcessing) {
        debugAudioSessions();
        this_thread::sleep_for(chrono::milliseconds(500));
    }

    serialThread.join();
    fetchThread.join();
    return 0;
}
