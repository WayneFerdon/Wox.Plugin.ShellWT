# ----------------------------------------------------------------
# Author: wayneferdon wayneferdon@hotmail.com
# Date: 2022-02-12 06:25:55
# LastEditors: wayneferdon wayneferdon@hotmail.com
# LastEditTime: 2022-10-05 18:31:10
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

class ShellWT(WoxQuery):
    def query(self, queryString):
        icon = './Images/terminal.ico'
        history = dict()
        if not os.path.isfile(HistoryFilePath):
            with open(HistoryFilePath,mode='w') as f:
                f.write(str(dict()))
        with open(HistoryFilePath, mode='r') as f:
            history = json.load(f)
        result = list()
        foundInHistory = False
        for data in history.keys():
            if queryString in data:
                subTitle = 'Has been run ' + json.dumps(history[data]) + ' times before.'
                result.append(WoxResult(data, subTitle, icon, data, self.Run.__name__, True, data, False).toDict())
                if queryString == data:
                    foundInHistory = True

        if not foundInHistory:
            result.append(WoxResult(queryString, 'Press Enter to Run', icon, queryString, self.Run.__name__, True, queryString, False).toDict())

        runPowerShell = WoxResult('Run Windows PowerShell', 'Press Enter to Run', icon, '> PowerShell', self.RunPowerShell.__name__, True, False).toDict()
        runCMD = WoxResult('Run CMD', 'Press Enter to Run', icon, '> CMD', self.RunCMD.__name__, True, False).toDict()

        runResult =  [runPowerShell,runCMD]
        if queryString == '':
            return runResult
        for each in runResult:
            if queryString.lower() in each['Title'].lower():
                result.append(each)
        return result

    def context_menu(self, queryString):
        iconPath = './Images/terminal.ico'
        iconPath = os.path.join(os.path.abspath('./'),iconPath)
        subTitle = 'Press Enter to Run as Administrator'
        if '>' not in queryString:
            title = queryString
            method = self.Run.__name__
            return [WoxResult(title, subTitle, iconPath, None, method, True, queryString, True).toDict()]
        if queryString == '> PowerShell':
            title = 'Run Windows PowerShell as Administrator'
            method = self.RunPowerShell.__name__
        elif queryString == '> CMD':
            title = 'Run CMD PowerShell as Administrator'
            method = self.RunCMD.__name__
        return [WoxResult(title, subTitle, iconPath, None, method, True, True).toDict()]

    @classmethod
    def Run(cls, cmd, isRunAsAdministrator):
        cmd = 'wt powershell ' + cmd + '\npause'
        if isRunAsAdministrator:
            cmd = 'sudo -n ' +cmd
        
        subprocess.Popen(cmd, cwd=os.getenv("SystemRoot")+"\system32")
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

    @classmethod
    def RunPowerShell(cls, isRunAsAdministrator):
        if(isRunAsAdministrator):
            subprocess.Popen("sudo -n wt powershell -NoExit ", cwd=os.getenv("SystemRoot")+"\system32")
        else:
            subprocess.Popen('wt powershell', cwd=os.getenv("SystemRoot")+"\system32")
    
    @classmethod
    def RunCMD(cls, isRunAsAdministrator):
        if(isRunAsAdministrator):
            subprocess.Popen("sudo -n wt cmd -NoExit ", cwd=os.getenv("SystemRoot")+"\system32")
        else:
            subprocess.Popen('wt cmd', cwd=os.getenv("SystemRoot")+"\system32")

if __name__ == '__main__':
     ShellWT()
