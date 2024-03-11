import tkinter as tk
from tkinter import ttk
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume,ISimpleAudioVolume
import PIL.ImageGrab as ImageGrab
import time
import pyautogui
import shutil
import os


def write_variables_backup(coordinates,variable_value):  #Writes pixel coordinates to the var.txt file.
  with open("var.txt", "w") as f:
    for key, value in coordinates.items():
      f.write(f"{key}={value[0]},{value[1]}\n")
    for key, value in variable_value.items():
      f.write(f"{key}={value}\n")

def read_variables():
    try:
        with open("var.txt", "r") as f:
            lines = f.readlines()
            variables = {}

            # Iterate through lines once, checking for comments and parsing key-value pairs
            for line in lines:
                if line.strip().startswith("#"):  # Skip comment lines
                    continue
                if line.strip().startswith("Pixel"):  # Skip Pixel lines
                    continue
                if "=" in line:  # Check if line contains an equals sign
                    key, value = line.strip().split("=")
                    variables[key] = value
            return variables
    except FileNotFoundError:
        with open("var.txt", "w") as f:
            f.write("Pixel_Blue_Mouse_Spectate=137,858\n")
            f.write("Pixel_White_Mouse_Spectate=141,869\n")
            f.write("Pixel_Two_Dots_Timer=950,57\n")
            f.write("Pixel_Spike=950,65\n")
            f.write("Pixel_Timer_Red=960,58\n")
            f.write("Pixel_Buy_Phase=975,201\n")
            f.write("Default_Music_App=spotify\n")
            f.write("Default_Volume_Valorant_Alive=0.85\n")
            f.write("Default_Volume_Valorant_Dead=0.4\n")
            f.write("Default_Volume_Music_Alive=0.05\n")
            f.write("Default_Volume_Music_Dead=0.55\n")

def read_pixel_coordinates():  #  Reads pixel coordinates from the var.txt file.
    try:
        with open("var.txt", "r") as f:
            lines = f.readlines()
            coordinates = {}

            # Iterate through lines once, checking for comments and parsing key-value pairs
            for line in lines:
                if line.strip().startswith("#"):  # Skip comment lines
                    continue
                if line.strip().startswith("Pixel"): # Only consider coordinates lines
                    key, value = line.strip().split("=")
                    x, y = value.split(",")
                    coordinates[key] = (int(x), int(y))
            return coordinates
    except FileNotFoundError:
        print("var.txt Not Found")
  
# Call read_pixel_coordinates to get the dictionary
coordinates = read_pixel_coordinates()
variable_value = read_variables()

# Get variables from var.txt
Pixel_Blue_Mouse_Spectate = coordinates.get("Pixel_Blue_Mouse_Spectate")  # Cyan color check (Spectator mode indicator)
Pixel_White_Mouse_Spectate = coordinates.get("Pixel_White_Mouse_Spectate")  # White color check (Confirmation of spectator mode)
Pixel_Two_Dots_Timer = coordinates.get("Pixel_Two_Dots_Timer")  # Timer check (Round timer)
Pixel_Spike = coordinates.get("Pixel_Spike")  # Spike check (Spike icon)
Pixel_Timer_Red = coordinates.get("Pixel_Timer_Red")  # Timer (Red) check (Round timer)
Pixel_Buy_Phase = coordinates.get("Pixel_Buy_Phase")  # White color check (Buy phase indicator)

Default_Music_App=variable_value.get("Default_Music_App")
Default_Volume_Valorant_Alive=variable_value.get("Default_Volume_Valorant_Alive")
Default_Volume_Valorant_Dead=variable_value.get("Default_Volume_Valorant_Dead")
Default_Volume_Music_Alive=variable_value.get("Default_Volume_Music_Alive")
Default_Volume_Music_Dead=variable_value.get("Default_Volume_Music_Dead")



def on_music_dead_volume_changed(value):
    global music_dead_volume_label
    music_dead_volume_label.config(text="Volume when Dead: "+ str(round(float(value)*100, 2)) +"%")

def on_music_game_volume_changed(value):
    global music_game_volume_label
    music_game_volume_label.config(text="Volume when Alive : "+ str(round(float(value)*100, 2)) +"%")

def on_valo_dead_volume_changed(value):
    global valo_dead_volume_label
    valo_dead_volume_label.config(text="Volume when Dead: "+ str(round(float(value)*100, 2)) +"%")

def on_valo_game_volume_changed(value):
    global valo_game_volume_label
    valo_game_volume_label.config(text="Volume when Alive : "+ str(round(float(value)*100, 2)) +"%")

def get_running_applications():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)


    sessions = AudioUtilities.GetAllSessions()
    apps = [session.Process.name() for session in sessions if session.Process]
    return apps

dead = True

def check_white(c):
     return c[0] >= 235 and c[1] >= 235 and c[2] >= 235
def check_cyan(c):
     return c[0] >= 165 and c[0]<= 175 and c[1] >= 230 and c[0]<= 245 and c[2] >= 220 and c[2]<= 230
def check_timer(c):
    return (c[0] >= 250 and c[1] >= 250 and c[2] >= 250) 
def check_spike(c):
    return (c[0]>= 120  and c[1] <= 5  and c[2] <= 5)
def check_timer_red(c):
    return (c[0]>= 250 and c[1] <= 5 and c[2] <= 5)


def detect_dead():
  global dead
  px = ImageGrab.grab().load()
  if dead: 
    if not(check_cyan(px[Pixel_Blue_Mouse_Spectate]) and check_white(px[Pixel_White_Mouse_Spectate])) and (check_timer(px[Pixel_Two_Dots_Timer]) or check_spike(px[Pixel_Spike]) or check_timer_red(px[Pixel_Timer_Red])) and not check_white(px[Pixel_Buy_Phase]):
      dead = False
      update_volume()
  else:
    if (check_cyan(px[Pixel_Blue_Mouse_Spectate]) and check_white(px[Pixel_White_Mouse_Spectate])) or (not check_timer(px[Pixel_Two_Dots_Timer]) and not check_timer_red(px[Pixel_Timer_Red]) and not check_spike(px[Pixel_Spike])) or check_white(px[Pixel_Buy_Phase]):
      dead = True
      update_volume()
  root.after(2000, detect_dead)


def update_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    sessions = AudioUtilities.GetAllSessions()
    global music_dead_volume_slider, music_game_volume_slider, valo_dead_volume_slider, valo_game_volume_slider, dead, music_selected_application, valo_selected_application
    music = None
    valorant = None
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process and session.Process.name() == music_selected_application.get():
            music = volume
        if session.Process and session.Process.name() == valo_selected_application.get():
            valorant = volume
        if music is not None and valorant is not None:
            break
    if valorant is not None and music is not None:
        if dead:
            for i in range(20):
                valorant.SetMasterVolume((valo_game_volume_slider.get() + i * (valo_dead_volume_slider.get() - valo_game_volume_slider.get()) / 20),None)
                music.SetMasterVolume((music_game_volume_slider.get() + i * (music_dead_volume_slider.get() - music_game_volume_slider.get()) / 20),None)
                time.sleep(0.1)
        else:
            for i in range(20):
                valorant.SetMasterVolume((valo_dead_volume_slider.get() + i * (valo_game_volume_slider.get() - valo_dead_volume_slider.get()) / 20),None)
                music.SetMasterVolume((music_dead_volume_slider.get() + i * (music_game_volume_slider.get() - music_dead_volume_slider.get()) / 20),None)
                time.sleep(0.1)


def on_closing():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    sessions = AudioUtilities.GetAllSessions()
    global music_dead_volume_slider, music_game_volume_slider, valo_dead_volume_slider, valo_game_volume_slider, dead, music_selected_application, valo_selected_application
    music = None
    valorant = None
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process and session.Process.name() == music_selected_application.get():
            music = volume
        if session.Process and session.Process.name() == valo_selected_application.get():
            valorant = volume
        if music is not None and valorant is not None:
            break
    if valorant is not None and music is not None:
        valorant.SetMasterVolume(0.5, None)
        music.SetMasterVolume(0.5, None)
    root.destroy()


def set_default_dropdown(name,dropdown):
    applications = get_running_applications()
    for app in applications:
        if name in app.lower():
            dropdown.set(app)
            break


def update_dropdown():
    global music_application_dropdown, valo_application_dropdown
    applications = get_running_applications()
    
    # Clear previous options
    music_application_dropdown['values'] = []
    valo_application_dropdown['values'] = []

    # Update options
    music_application_dropdown['values'] = applications
    valo_application_dropdown['values'] = applications

    root.after(10000, update_dropdown)  # Schedule the next update after 10 seconds





def update_coordinates(event=None):
  selected_pixel_name = selected_pixel.get()
  coordinates = read_pixel_coordinates()
  if selected_pixel_name in coordinates:
    x, y = coordinates[selected_pixel_name]  # Unpack X and Y coordinates
    x_coord_var.set(f"{x}")  # Set X coordinate with label
    y_coord_var.set(f"{y}")  # Set Y coordinate with label
  else:
    x_coord_var.set("Coordinates not found")
    y_coord_var.set("Coordinates not found")



def update_coordinates_with_mouse():
    selected_pixel_name = selected_pixel.get()
    if selected_pixel_name:
        x, y = pyautogui.position()
        x_coord_var.set(str(x))
        y_coord_var.set(str(y))
        save_coordinates_to_file()


def save_coordinates_to_file(event=None):
    selected_pixel_name = selected_pixel.get()
    if selected_pixel_name:
        x = x_coord_var.get()
        y = y_coord_var.get()
        if x.isdigit() and y.isdigit():  # Check if coordinates are numbers
            # Create a backup of the current var.txt file
            try:
                shutil.copyfile("var.txt", "var_backup.txt")
            except FileNotFoundError:
                pass  # Ignore if the file doesn't exist yet

            # Update the coordinates dictionary
            coordinates[selected_pixel_name] = (int(x), int(y))

            # Write the updated coordinates to the var.txt file
            write_variables_backup(coordinates,variable_value)
            
            print("Coordinates saved")
        else:
            print("Invalid coordinates: Please enter only numbers for X and Y.")





root = tk.Tk()
root.title("Valorant Music Controller")
root.geometry("650x500")

# Styling
style = ttk.Style()
style.configure("TLabel", font=("Calibri", 12))
style.configure("TScale", padding=10)

main_frame = ttk.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

music_frame = ttk.LabelFrame(main_frame, text="Music")
music_frame.pack(side="left", padx=10, pady=10, fill=tk.BOTH, expand=True)

valorant_frame = ttk.LabelFrame(main_frame, text="Valorant")
valorant_frame.pack(side="right", padx=10, pady=10, fill=tk.BOTH, expand=True)

# Music Frame Widgets
music_settings_frame = ttk.Frame(music_frame)
music_settings_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

music_dead_volume_label = ttk.Label(music_settings_frame, text="Volume when Dead: 0%")
music_dead_volume_label.pack(padx=10, pady=5)

music_dead_volume_slider = ttk.Scale(music_settings_frame, from_=0, to=1, orient="horizontal", command=on_music_dead_volume_changed)
music_dead_volume_slider.pack(fill="x", padx=10, pady=5)
music_dead_volume_slider.set(Default_Volume_Music_Dead)

music_game_volume_label = ttk.Label(music_settings_frame, text="Volume when Alive: 0%")
music_game_volume_label.pack(padx=10, pady=5)

music_game_volume_slider = ttk.Scale(music_settings_frame, from_=0, to=1, orient="horizontal", command=on_music_game_volume_changed)
music_game_volume_slider.pack(fill="x", padx=10, pady=5)
music_game_volume_slider.set(Default_Volume_Music_Alive)

music_application_label = ttk.Label(music_settings_frame, text="Select Application")
music_application_label.pack(padx=10, pady=5)

music_selected_application = tk.StringVar()
music_application_dropdown = ttk.Combobox(music_settings_frame, textvariable=music_selected_application, state="readonly")
music_application_dropdown.pack(fill="x", padx=10, pady=5)
set_default_dropdown(Default_Music_App, music_selected_application)


# Valorant Frame Widgets
valo_settings_frame = ttk.Frame(valorant_frame)
valo_settings_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

valo_dead_volume_label = ttk.Label(valo_settings_frame, text="Volume when Dead: 0%")
valo_dead_volume_label.pack(padx=10, pady=5)

valo_dead_volume_slider = ttk.Scale(valo_settings_frame, from_=0, to=1, orient="horizontal", command=on_valo_dead_volume_changed)
valo_dead_volume_slider.pack(fill="x", padx=10, pady=5)
valo_dead_volume_slider.set(Default_Volume_Valorant_Dead)

valo_game_volume_label = ttk.Label(valo_settings_frame, text="Volume when Alive: 0%")
valo_game_volume_label.pack(padx=10, pady=5)

valo_game_volume_slider = ttk.Scale(valo_settings_frame, from_=0, to=1, orient="horizontal", command=on_valo_game_volume_changed)
valo_game_volume_slider.pack(fill="x", padx=10, pady=5)
valo_game_volume_slider.set(Default_Volume_Valorant_Alive)

valo_application_label = ttk.Label(valo_settings_frame, text="Select Application")
valo_application_label.pack(padx=10, pady=5)

valo_selected_application = tk.StringVar()
valo_application_dropdown = ttk.Combobox(valo_settings_frame, textvariable=valo_selected_application, state="readonly")
valo_application_dropdown.pack(fill="x", padx=10, pady=5)
set_default_dropdown("valorant", valo_selected_application)

# Pixel Selector Frame Widgets
pixel_selector_frame = ttk.LabelFrame(root, text="Pixel Selector", padding=(10, 5))
pixel_selector_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

pixel_selector_subtitle = ttk.Label(pixel_selector_frame, text="Press F9 to capture mouse's coordinates")
pixel_selector_subtitle.pack(side="top", pady=(0, 5))

pixel_options = ["Pixel_Blue_Mouse_Spectate", "Pixel_White_Mouse_Spectate", "Pixel_Two_Dots_Timer", "Pixel_Spike", "Pixel_Timer_Red", "Pixel_Buy_Phase"]
selected_pixel = tk.StringVar()
pixel_dropdown = ttk.Combobox(pixel_selector_frame, textvariable=selected_pixel, values=pixel_options, state="readonly")
pixel_dropdown.pack(side="top", padx=10, pady=5, fill="x")
pixel_dropdown.bind("<<ComboboxSelected>>", update_coordinates)

coord_frame = ttk.Frame(pixel_selector_frame)
coord_frame.pack(padx=10, pady=5, fill=tk.BOTH)

x_coord_label = ttk.Label(coord_frame, text="X:")
x_coord_label.grid(row=0, column=0, padx=(10, 5), pady=5)

x_coord_var = tk.StringVar()
x_coord_entry = ttk.Entry(coord_frame, textvariable=x_coord_var)
x_coord_entry.grid(row=0, column=1, padx=(0, 10), pady=5)

y_coord_label = ttk.Label(coord_frame, text="Y:")
y_coord_label.grid(row=0, column=2, padx=(10, 5), pady=5)

y_coord_var = tk.StringVar()
y_coord_entry = ttk.Entry(coord_frame, textvariable=y_coord_var)
y_coord_entry.grid(row=0, column=3, padx=(0, 10), pady=5)

coord_frame.columnconfigure((0, 1, 2, 3), weight=1)
coord_frame.rowconfigure(0, weight=1)

x_coord_entry.bind("<KeyRelease>", save_coordinates_to_file)
y_coord_entry.bind("<KeyRelease>", save_coordinates_to_file)



root.bind("<F9>", lambda event: update_coordinates_with_mouse())
update_dropdown()  # Start the initial update of the dropdown
detect_dead()
root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()