import os
import pathlib
import urllib.request
import subprocess

os.system("chcp 65001")

print("Install fbx sdk for visual studio? y/n")
ok = input() == "y"

if ok:
    if not pathlib.Path("fbxsdk.exe").exists():
        print("Downloading...")
        urllib.request.urlretrieve("https://damassets.autodesk.net/content/dam/autodesk/www/adn/fbx/2020-0-1/fbx202001_fbxsdk_vs2017_win.exe", "fbxsdk.exe")

    process = subprocess.Popen("fbxsdk.exe", shell=True)
    process.wait()

    print("Delete the fbxsdk installer? y/n")
    ok = input() == "y"
    os.remove("fbxsdk.exe")
