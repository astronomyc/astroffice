import tkinter as tk
from tkinter import ttk
import xml.etree.ElementTree as ET
import shutil
import os
import subprocess
import sys

# Definiciones
def restoreFromBackup():
    shutil.copyfile('configbackup', 'config.xml')

def defaultSettings():
    
    restoreFromBackup()

    section_install()

    boxProduct.set(product[0])
    boxArch.set(arch[0])
    boxLang.set(lang[0])

    cbWord.set(True)
    cbExcel.set(True)
    cbPowerpoint.set(True)
    cbAccess.set(True)
    cbPublisher.set(True)
    cbOnenote.set(True)
    cbOutlook.set(True)
    cbDrive.set(True)

def run_cmd(command):
    os.system(command)

def section_install():
    navInstall.config(fg="#840DDD")
    navUninstall.config(fg="#FFF")
    navActivate.config(fg="#FFF")
    frameInstall.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    frameUninstall.pack_forget()
    frameActivate.pack_forget()

def section_uninstall():
    navInstall.config(fg="#FFF")
    navUninstall.config(fg="#840DDD")
    navActivate.config(fg="#FFF")
    frameInstall.pack_forget()
    frameUninstall.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    frameActivate.pack_forget()

def section_activate():
    navInstall.config(fg="#FFF")
    navUninstall.config(fg="#FFF")
    navActivate.config(fg="#840DDD")
    frameInstall.pack_forget()
    frameUninstall.pack_forget()
    frameActivate.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

def changeConfig(filename, varProduct, varArch, varLang, cbWord, cbExcel, cbPowerpoint, cbAccess, cbPublisher, cbOnenote, cbOutlook, cbDrive):
    tree = ET.parse(filename)
    root = tree.getroot()

    appsProducts = {
        "Word": cbWord,
        "Excel": cbExcel,
        "PowerPoint": cbPowerpoint,
        "Access": cbAccess,
        "Publisher": cbPublisher,
        "OneNote": cbOnenote,
        "Outlook": cbOutlook,
        "OneDrive": cbDrive
    }

    for product in root.iter('Product'):
        product.set('ID', varProduct)

    for add in root.iter('Add'):
        add.set('OfficeClientEdition', varArch)

    for language in root.iter('Language'):
        language.set('ID', varLang)

    for app, var in appsProducts.items():
        if var:
            for product in root.iter('Product'):
                for exclude_app in product.findall(f".//ExcludeApp[@ID='{app}']"):
                    product.remove(exclude_app)

    tree.write(filename)

def productFn(event):
    global varProduct
    varProduct = boxProduct.get()

def archFn(event):
    global varArch
    varArch = boxArch.get()

def langFn(event):
    global varLang
    varLang = boxLang.get()

def installFn():
    changeConfig('config.xml', varProduct, varArch, varLang, cbWord.get(), cbExcel.get(), cbPowerpoint.get(), cbAccess.get(), cbPublisher.get(), cbOnenote.get(), cbOutlook.get(), cbDrive.get())
    run_cmd("setup /configure config.xml")
    defaultSettings()

def uninstallFn():
    run_cmd("setup /configure configuninstall.xml")

def activateFn():
        subprocess.Popen(["powershell", "Start-Process powershell -Verb RunAs -ArgumentList 'irm https://massgrave.dev/get |iex'"], creationflags=subprocess.CREATE_NEW_CONSOLE)

# Ventana
root = tk.Tk()
root.title("Astroffice")
root.iconbitmap('icon.ico')
root.geometry("800x600")
root.resizable(width=False, height=False)
root.config(bg='#333')

# NavBar
navBar = tk.Frame(root, bg="#222")

navInstall = tk.Button(navBar, text="Install", bd="0", bg="#222222", fg="#FFFFFF", command=section_install, cursor="hand2", activebackground="#2A2A2A")
navUninstall = tk.Button(navBar, text="Uninstall", bd="0", bg="#222222", fg="#FFFFFF", command=section_uninstall, cursor="hand2", activebackground="#2A2A2A")
navActivate = tk.Button(navBar, text="Activate", bd="0", bg="#222222", fg="#FFFFFF", command=section_activate, cursor="hand2", activebackground="#2A2A2A")

## Posicionar NavBar
navBar.pack(side=tk.TOP, fill=tk.X)
navInstall.pack(side=tk.LEFT)
navUninstall.pack(side=tk.LEFT)
navActivate.pack(side=tk.LEFT)

# Contenedores
frameInstall = tk.Frame(root, bg="#333")
frameInstall.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

frameUninstall = tk.Frame(root, bg="#333")
frameUninstall.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

frameActivate = tk.Frame(root, bg="#333")
frameActivate.pack(side=tk.TOP, fill=tk.BOTH, expand=True)


frameInstallTop = tk.Frame(frameInstall, bg="#333")
frameInstallTop.pack(side=tk.TOP, fill=tk.X)

frameInstallMiddle = tk.Frame(frameInstall, bg="#333")
frameInstallMiddle.pack(side=tk.TOP, fill=tk.X)

# Contenedor Superior
product = ['ProPlus2021Volume', 'ProPlus2019Volume']
labelProduct = ttk.Label(frameInstallTop, text="Product:", background="#333", foreground="#FFF", font=(12))

boxProduct = ttk.Combobox(frameInstallTop, values=product, state="readonly")
boxProduct.bind("<<ComboboxSelected>>", productFn)


arch = ['64', '32']
labelArch = ttk.Label(frameInstallTop, text="Architecture", background="#333", foreground="#FFF", font=(12))

boxArch = ttk.Combobox(frameInstallTop, values=arch, state="readonly")
boxArch.bind("<<ComboboxSelected>>", archFn)


lang = ['en-US', 'es-ES']
labelLang = ttk.Label(frameInstallTop, text="Lang:", background="#333", foreground="#FFF", font=(12))

boxLang = ttk.Combobox(frameInstallTop, values=lang, state="readonly")
boxLang.bind("<<ComboboxSelected>>", langFn)

## Posisionar Items del Contenedor Superior
labelProduct.pack(side=tk.LEFT, anchor="w", padx=(30, 5), pady=20)
boxProduct.pack(side=tk.LEFT, anchor="w", padx=5, pady=20)

labelArch.pack(side=tk.LEFT, anchor="w", padx=(30, 5), pady=20)
boxArch.pack(side=tk.LEFT, anchor="w", padx=5, pady=20)

labelLang.pack(side=tk.LEFT, anchor="w", padx=(30, 5), pady=20)
boxLang.pack(side=tk.LEFT, anchor="w", padx=5, pady=20)

# Contenedor Central
cbWord = tk.BooleanVar()
cbWo = ttk.Checkbutton(frameInstallMiddle, text="Word", variable=cbWord, style='Custom.TCheckbutton')

cbExcel = tk.BooleanVar()
cbEx = ttk.Checkbutton(frameInstallMiddle, text="Excel", variable=cbExcel, style='Custom.TCheckbutton')

cbPowerpoint = tk.BooleanVar()
cbPo = ttk.Checkbutton(frameInstallMiddle, text="Powerpoint", variable=cbPowerpoint, style='Custom.TCheckbutton')

cbAccess = tk.BooleanVar()
cbAc = ttk.Checkbutton(frameInstallMiddle, text="Access", variable=cbAccess, style='Custom.TCheckbutton')

cbPublisher = tk.BooleanVar()
cbPu = ttk.Checkbutton(frameInstallMiddle, text="Publisher", variable=cbPublisher, style='Custom.TCheckbutton')

cbOnenote = tk.BooleanVar()
cbOn = ttk.Checkbutton(frameInstallMiddle, text="Onenote", variable=cbOnenote, style='Custom.TCheckbutton')

cbOutlook = tk.BooleanVar()
cbOu = ttk.Checkbutton(frameInstallMiddle, text="Outlook", variable=cbOutlook, style='Custom.TCheckbutton')

cbDrive = tk.BooleanVar()
cbDr = ttk.Checkbutton(frameInstallMiddle, text="OneDrive", variable=cbDrive, style='Custom.TCheckbutton')


style = ttk.Style()
style.configure('Custom.TCheckbutton', background='#333', font=(12), foreground='#FFF')

installBtn = ttk.Button(frameInstallMiddle, text="Install", command=installFn)

## Posisionar Items del Contenedor Central
cbWo.pack()
cbEx.pack()
cbPo.pack()
cbAc.pack()
cbPu.pack()
cbOn.pack()
cbOu.pack()
cbDr.pack()

### Posisionar Boton de Instalar
installBtn.pack()

varProduct = product[0]
varArch = arch[0]
varLang = lang[0]

###Boton de Desinstalar
uninstallBtn = ttk.Button(frameUninstall, text="Uninstall", command=uninstallFn)
uninstallBtn.pack()

###Boton de Activar
activateBtn = ttk.Button(frameActivate, text="Activate", command=activateFn)
activateBtn.pack()

defaultSettings()
root.mainloop()