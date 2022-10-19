# ----------------------------------------------------------------
# Author: wayneferdon wayneferdon@hotmail.com
# Date: 2022-02-12 06:25:55
# LastEditors: wayneferdon wayneferdon@hotmail.com
# LastEditTime: 2022-10-20 01:03:12
# FilePath: \Wox.Plugin.ShellWT\main.py
# ----------------------------------------------------------------
# Copyright (c) 2022 by Wayne Ferdon Studio. All rights reserved.
# Licensed to the .NET Foundation under one or more agreements.
# The .NET Foundation licenses this file to you under the MIT license.
# See the LICENSE file in the project root for more information.
# ----------------------------------------------------------------

# -*- coding: utf-8 -*-

import os
from WoxQuery import *
import subprocess   
import json

PackageDir = os.path.join(os.environ['localAppData'.upper()],'Packages')
HistoryFilePath = './History.json'
ROOT_PATH = os.getenv("SystemRoot")+"\system32"
TERMINAL_ICON = os.path.join(os.path.abspath('./'), './Images/terminal.ico')

class ShellWT(WoxQuery):
    def query(self, queryString:str):
        runPowerShell = WoxResult('Run Windows PowerShell', 'Press Enter to Run', TERMINAL_ICON, '> PowerShell', self.Run.__name__, True, "", False).toDict()
        runCMD = WoxResult('Run CMD', 'Press Enter to Run', TERMINAL_ICON, '> CMD', self.RunCMD.__name__, True, "", False).toDict()
        runResult =  [runPowerShell,runCMD]
        if queryString == '':
            return runResult
        
        history = dict()
        if not os.path.isfile(HistoryFilePath):
            with open(HistoryFilePath,mode='w') as f:
                f.write(str(dict()))
        with open(HistoryFilePath, mode='r') as f:
            history = json.load(f)
        result = list()
        foundInHistory = False
        for data in history.keys():
            if queryString not in data:
                continue
            subTitle = 'Has been run ' + json.dumps(history[data]) + ' times before.'
            result.append(WoxResult(data, subTitle, TERMINAL_ICON, data, self.Run.__name__, True, data, False).toDict())
            if queryString == data:
                foundInHistory = True

        if not foundInHistory:
            result.append(WoxResult(queryString, 'Press Enter to Run', TERMINAL_ICON, queryString, self.Run.__name__, True, queryString, False).toDict())

        for each in runResult:
            if queryString.lower() in each['Title'].lower():
                result.append(each)
        return result

    def context_menu(self, queryString:str):
        subTitle = 'Press Enter to Run as Administrator'
        if '>' not in queryString:
            title = queryString
            method = self.Run.__name__
            return [WoxResult(title, subTitle, TERMINAL_ICON, None, method, True, queryString, True).toDict()]
        if queryString == '> PowerShell':
            title = 'Run Windows PowerShell as Administrator'
            method = self.Run.__name__
        elif queryString == '> CMD':
            title = 'Run CMD PowerShell as Administrator'
            method = self.RunCMD.__name__
        return [WoxResult(title, subTitle, TERMINAL_ICON, None, method, True, "", True).toDict()]

    def Run(self, cmd:str, isRunAsAdministrator:bool):
        cmd = 'wt powershell -NoExit ' + cmd
        if isRunAsAdministrator:
            cmd = 'sudo ' + cmd
        
        subprocess.run(cmd, cwd=ROOT_PATH, shell=True)
        if not os.path.isfile(HistoryFilePath):
            with open(HistoryFilePath,mode='w') as f:
                f.write(str(dict()))
        history = dict()
        with open(HistoryFilePath, mode='r') as f:
            history = json.load(f)

            if cmd not in history.keys():
                history[cmd] = 0
            history[cmd] += 1
        with open(HistoryFilePath, mode='w+') as f:
            f.write(json.dumps(history))

    def RunCMD(self, cmd:None, isRunAsAdministrator:bool):
        if(isRunAsAdministrator):
            subprocess.run("sudo wt cmd", cwd=ROOT_PATH, shell = True)
        else:
            subprocess.run('wt cmd', cwd=ROOT_PATH)

if __name__ == '__main__':
    ShellWT()
