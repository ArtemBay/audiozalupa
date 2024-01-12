import tkinter
import tkinter.messagebox
import customtkinter
import sys
import serial
import os
import json
import win32api  # For various API functions
import win32con  # For Windows constants
import win32gui  # For GUI-related functions
import win32process  # For process-related functions
import win32security  # For security-related functions
import pythoncom

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"
comstate = "normal"


def change_volume(app: str, volume: int):
    PyIUnknown = CoCreateInstance(clsid, unkOuter, context, iid)
    CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator)
    IID_IMMDeviceEnumerator = __uuidof(IMMDeviceEnumerator)
    mmde = pythoncom.CoCreateInstance(CLSID_MMDeviceEnumerator, None, CLSCTX_ALL, IID_IMMDeviceEnumerator)
    mmd = mmde.GetDefaultAudioEndpoint(eRender, eMultimedia)
    mgr = mmd.Activate(IID_IAudioSessionManager)
    sav = mgr.GetSimpleAudioVolume(None, True)
    sav.SetMasterVolume(0.5)


def read_serial(port: str):
    ser = serial.Serial(port, baudrate=115200, timeout=1)
    i = 0
    while i < 5:
        data = ser.readline().decode("ISO-8859-1").strip()
        i += 1
    return data


def write_serial(port: str, pos1: str, pos2: str, pos3: str):
    ser = serial.Serial(port, baudrate=115200, timeout=1)
    ser.write(f"{pos1},{pos2},{pos3}".encode('ascii'))
    i = 0
    while i < 5:
        data = ser.readline().decode("ISO-8859-1").strip()
        i += 1
    return data


def serial_ports():
    if sys.platform.startswith('win'):
        ports = ['COM%s' % (i + 1) for i in range(256)]
    else:
        raise EnvironmentError('Unsupported platform')
    result = []
    for port in ports:
        if port != "COM1":
            try:
                s = serial.Serial(port)
                s.close()
                result.append(port)
            except (OSError, serial.SerialException):
                pass
    if not result:
        tkinter.messagebox.showerror(title="No connected devices found",
                                     message="No connected devices found\nTry to check device connection with PC")
        global comstate
        comstate = "disabled"
    return result


def change_appearance_mode_event(new_appearance_mode: str):
    customtkinter.set_appearance_mode(new_appearance_mode)


def change_com_port(port: str):
    print("New com port: " + port)


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("Audiozalupa - By ArtemBay")
        self.geometry(f"{1100}x{580}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="AudioZalupa",
                                                 font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.comselector_label = customtkinter.CTkLabel(self.sidebar_frame, text="Select device COM Port:", anchor="w")
        self.comselector_optionmenu = customtkinter.CTkOptionMenu(self.sidebar_frame,
                                                                  values=serial_ports(),
                                                                  command=change_com_port)
        self.comselector_optionmenu.grid(row=4, column=0, padx=20, pady=(10, 360))
        self.comselector_label.grid(row=3, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame,
                                                                       values=["Light", "Dark", "System"],
                                                                       command=change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))

        self.appearance_mode_optionemenu.set("Dark")
        if port != "None":
            self.comselector_optionmenu.set(port)
        else:
            self.comselector_optionmenu.set("Change me")
        self.comselector_optionmenu.configure(state=comstate)


if __name__ == "__main__":
    dev = True
    if not dev:
        if "config.json" not in os.listdir('./'):
            with open('config.json', 'w') as f:
                f.write('{"port":"None"}')
        else:
            with open('config.json', 'r', encoding="utf-8") as f:
                config_json = json.loads(str(f.read()))
                port = config_json['port']
        app = App()
        app.mainloop()
    else:
        res = read_serial("COM3")
        # res = write_serial("COM3", "discord", "speaker", "game") #DONT WORK
        print(res)
        #change_volume("discord", 100)
