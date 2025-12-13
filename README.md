# ScriptHookVPy 🐍

<img width="2048" height="2048" alt="ScriptHookVPy" src="https://github.com/user-attachments/assets/8bbe93f2-7d60-4c62-9ef8-08626948f49c" />


**Python scripting for Grand Theft Auto V**

Inspired by ScriptHookVDotNet, ScriptHookVPy allows you to write GTA V mods in Python with full access to all native functions.

![Version](https://img.shields.io/badge/version-1.0.1-blue)
![Python](https://img.shields.io/badge/python-3.11-yellow)
![License](https://img.shields.io/badge/license-MIT-green)

---

## Features

- 🐍 **Write mods in Python** — no C++ or C# needed
- 📦 **All natives included** — PLAYER, VEHICLE, PED, WEAPON, ENTITY, HUD and 50+ namespaces
- 🔄 **Auto-load scripts** — drop .py files into `pythons/` folder
- 📝 **Logging system** — debug with `ScriptHookVPy.log`
- ⚡ **Easy syntax** — `natives.PLAYER.PLAYER_PED_ID()`

---

## Installation

### Requirements
- [ScriptHookV](http://www.dev-c.com/gtav/scripthookv/)
- Python 3.11

### Steps
1. Download latest release
2. Extract `ScriptHookVPy.asi` and `python311.dll` to GTA V folder
3. Create `pythons/` folder in GTA V directory
4. Add your `.py` scripts to `pythons/`
5. Launch the game!

---

## How It Works

ScriptHookVPy loads all Python scripts from the `pythons/` folder and executes them in a loop.

## GTA5-MODS.COM

[ScriptHookVPy](https://www.gta5-mods.com/tools/scripthookvpy-python-scripting-for-gta-v)
