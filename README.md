# VolMixerEXE

VolMixerEXE is a tool to change the volume and mute state of the currently focused application. Designed to be used with something like AutoHotkey to change volume with a key press. 

## Usage

There are two commands, one to change the apps volume by X amount out of 100, and one to toggle the mute of the app. These are done like so:

```
VolMixerEXE.exe appvolume 5 (increases focused applications volume by 5 out of 100)
VolMixerEXE.exe appvolume -5 (lowers focused applications volume by 5 out of 100)
VolMixerEXE.exe appmutetoggle (Toggles the current applications mute state)
```

This can be hooked up with something like AutoHotkey like so: 

```
F18::Run "C:\path\to\VolMixerEXE.exe" appvolume 5
F19::Run "C:\path\to\VolMixerEXE.exe" appvolume -5
F20::Run "C:\path\to\VolMixerEXE.exe" appmutetoggle
```

If the exe is on your PATH then it can simply be:

```
F18::Run "VolMixerEXE.exe" appvolume 5
F19::Run "VolMixerEXE.exe" appvolume -5
F20::Run "VolMixerEXE.exe" appmutetoggle
```

## Why

This tool was originally created so that I could adjust the current applications volume through a rotary encoder on a custom keyboard. Turning the rotary encoder sends F18 and F19 for clockwise and anti-clockwise respectively. Pressing the encoder sends F20. AHK then calls this tool to change the volume without it tabbing me out of whatever it is I'm doing or needing to open the volume mixer myself.
