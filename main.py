from os import listdir
from sys import platform
from asyncio.tasks import run_coroutine_threadsafe
from asyncio.events import AbstractEventLoop, get_event_loop, new_event_loop

import ctypes
from typing import Any

import orjson
import serial
from thinker import messagebox
from customthinker import (
    CTk,
    CTkFrame,
    CTkLabel,
    CTkFont,
    CTkOptionMenu,
    set_appearance_mode,
    set_default_color_theme,
)
from pycaw.pycaw import AudioUtilities

set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"


class COM:
    state = "normal"


async def change_volume(app: str, new_volume: int):
    if app in await all_apps():
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session.SimpleAudioVolume
            if session.Process and session.Process.name() == f"{app}.exe":
                volume.SetMasterVolume(float(new_volume) / 100, None)


async def all_apps() -> list[str]:
    sessions = AudioUtilities.GetAllSessions()
    apps: list[str] = []
    for session in sessions:
        try:
            if not session.Process.name() == "LEDKeeper2.exe":
                apps.append(session.Process.name()[:-4])
        except (Exception,):
            pass
    return apps


async def change_mic_volume(volume: int):
    dll = ctypes.WinDLL('winmm')
    dll.waveOutSetVolume(volume, 0x3333)


async def read_serial(__port: str) -> list[int]:
    __serial = serial.Serial(__port, baudrate=115200, timeout=1)
    __index = 0
    __data: list[str] | None = None
    while __index < 2:
        __data = list(
            __serial.readline().decode("ISO-8859-1").strip()
        )
        __index += 1
    if __data is None:
        raise ValueError("Could not read serial port")
    for __index in range(len(__data)):
        if __data[__index] == ",":
            __data[__index] = ""
    return "".join(__data).split()


async def write_serial(
        __port: str,
        pos1: str,
        pos2: str,
        pos3: str
) -> list[int]:
    __serial = serial.Serial(__port, baudrate=115200, timeout=1)
    __serial.write(f"{pos1},{pos2},{pos3}".encode('ascii'))
    __index = 0
    __data: list[int] | None = None
    while __index < 2:
        __data = __serial.readline().decode("ISO-8859-1").strip()
        __index += 1
    if not data:
        raise ValueError("Could not read serial port")
    return __data


async def serial_ports() -> list[str]:
    if platform.startswith('win'):
        ports: list[str] = ['COM%s' % (i + 1) for i in range(256)]
    else:
        raise EnvironmentError('Unsupported platform')
    __result: list[str] = []
    for __port in ports:
        if __port != "COM1":
            try:
                __serial = serial.Serial(__port)
                __serial.close()
                __result.append(__port)
            except (OSError, serial.SerialException):
                pass
    if not __result:
        messagebox.showerror(
            title="No connected devices found",
            message="No connected devices found\nTry to check device connection with PC"
        )
        COM.state = "disabled"
    return __result


def change_appearance_mode_event(new_appearance_mode: str) -> None:
    set_appearance_mode(new_appearance_mode)


def change_com_port(__port: str) -> None:
    print("New com port: " + __port)


class App(CTk):
    def __init__(self) -> None:
        super().__init__()

        # configure window
        self.title("Audiozalupa - By ArtemBay")
        self.geometry(f"{1100}x{580}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        self.logo_label = CTkLabel(
            self.sidebar_frame,
            text="AudioZalupa",
            font=CTkFont(size=20, weight="bold")
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.com_select_label = CTkLabel(
            self.sidebar_frame,
            text="Select device COM Port:",
            anchor="w"
        )
        self.com_select_option_menu = CTkOptionMenu(
            self.sidebar_frame,
            values=serial_ports(),
            command=change_com_port
        )
        self.com_select_option_menu.grid(row=4, column=0, padx=20, pady=(10, 360))
        self.com_select_label.grid(row=3, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_label = CTkLabel(
            self.sidebar_frame, text="Appearance Mode:", anchor="w"
        )
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_option_menu = CTkOptionMenu(
            self.sidebar_frame,
            values=["Light", "Dark", "System"],
            command=change_appearance_mode_event
        )
        self.appearance_mode_option_menu.grid(row=6, column=0, padx=20, pady=(10, 10))
        self.appearance_mode_option_menu.set("Dark")

        self.com_select_option_menu.set("Change me" if port == "None" else port)
        self.com_select_option_menu.configure(state=COM.state)


if __name__ == "__main__":
    loop: AbstractEventLoop = get_event_loop() or new_event_loop()
    developer_mode: bool = False

    if not developer_mode:
        if "config.json" not in listdir('./'):
            with open('config.json', 'w') as stream:
                stream.write('{"port":"None"}')
        else:
            with open('config.json', 'rb') as stream:
                data: dict[str, Any] = orjson.loads(stream.read())
                port: str = data['port']

        serial_data: list[int] = loop.run_until_complete(read_serial("COM3"))
        run_coroutine_threadsafe(change_volume("Telegram", serial_data[0]), loop)
        app = App()
        app.mainloop()
    else:
        # change_mic_volume(100)
        exit(0)
