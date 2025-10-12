import json
import shutil
import os
import time
import zipfile
import pandas as pd
import re


def getJsonData(path):
    with open(path, "r") as read_file:
        return json.load(read_file)


def moveFile(Src, Dest):
    shutil.move(Src, Dest)


def removeFile(Src):
    if os.path.exists(Src):
        while True:
            try:
                os.remove(Src)
                break
            except:
                time.sleep(2)


def removeFolder(Src):
    if os.path.exists(Src):
        while True:
            try:
                shutil.rmtree(Src)
                break
            except:
                time.sleep(2)


def createFolder(Src):
    if not os.path.exists(Src):
        os.mkdir(Src)


def _zipFile(Src, Dest):
    name = os.path.splitext(os.path.basename(Src))[0]
    name = name[0][name[0].rfind("\\") + 1:]

    with zipfile.ZipFile(os.path.join(Dest, "{}.zip".format(name)), "w") as zipOject:
        for folderName, subfolders, filenames in os.walk(Src):
            for filename in filenames:
                filePath = os.path.join(folderName, filename)
                zipOject.write(filePath, os.path.basename(filename))


def _zipFolder(Src, Dest, parent):
    name = os.path.splitext(Src)
    name = name[0][name[0].rfind("\\") + 1:]

    shutil.make_archive("{}".format(name), "zip", Src)
    shutil.move(os.path.join(parent, "{}.zip".format(name)), Dest)


def removeNestings(arr):
    temp = []
    for x in arr:
        # print(x)
        if len(x.keys()) == 1:
            temp.append(x)
        else:
            for y, z in x.times():
                temp.append({y:z})

    return temp


def convertToString(list, delimiter):

    return "{0}".format(delimiter.join(list))


def unzipSingleFile(Src, Dest):
    with zipfile.ZipFile(Src, "r") as zipOject:
        # Get a list of all archived file names from the zip
        listOfFileNames = zipOject.namelist()
        print(listOfFileNames)

        # Iterate iver the file names
        for fileName in listOfFileNames:
            file = os.path.splitext(fileName)
            print(file)
            zipOject.extract(fileName, Dest)


def createFolder(path):
    if not os.path.exists(path):
        os.mkdir(path)


def unzipFileWithoutDirectoryStruture(Src, Dest, ext, action):
    print("Unzipping {}...".format(Src))
    print("Extracting {}...".format(Dest))

    dest = os.path.join(Dest)

    if not os.path.exists(dest):
        os.mkdir(dest)

    time.sleep(2)

    try:
        with zipfile.ZipFile(Src) as zip_file:
            for member in zip_file.namelist():
                filename = os.path.basename(member)
                # skip directories
                if not filename:
                    continue
                # copy file (taken from zipfile's extract)
                source = zip_file.open(member)
                sourceFileName = source.name
                sourceFileName = sourceFileName[sourceFileName.find("/") + 1:sourceFileName.find("-")-1]

                midPath = os.path.join(dest, sourceFileName)
                if not os.path.exists(midPath):
                    os.mkdir(midPath)

                destPath = os.path.join(dest, sourceFileName, action)
                if not os.path.exists(destPath):
                    os.mkdir(destPath)

                finalDest = os.path.join(destPath)
                if not os.path.exists(finalDest):
                    os.mkdir(finalDest)

                target = open(os.path.join(finalDest, filename), "wb")
                with source, target:
                    shutil.copyfileobj(source, target)

    except Exception as e:
        print(e)
        with open("error.text", "w") as writer:
            writer.write(str(e))


def winapi_path(dos_path, encoding=None):
    path = os.path.abspath(dos_path)

    if path.startswith("\\\\"):
        path = "\\\\?\\UNC\\" + path[2:]
    else:
        path = path + "\\\\?\\"

    return path


def convertTxt2Excel(Src, file, sheet):
    df = pd.read_excel(Src)
    df.to_excel(file, sheet, index=False)


def checkfileExist(Src, timeout):
    start_time = time.time()
    while not os.path.exists(Src):
        time.sleep(1)  # 每秒检查一次
        if time.time() - start_time > timeout:
            raise TimeoutError(f"{Src} 下载超时！")
    time.sleep(2)


class zipFileLongPaths(zipfile.ZipFile):
        def _extract_member(self, member, targetpath, pwd):
            targetpath = winapi_path(targetpath)
            return zipfile.ZipInfo.extract_member(self, member, targetpath, pwd)


