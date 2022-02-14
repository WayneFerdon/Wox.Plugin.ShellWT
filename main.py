# -*- coding: utf-8 -*-
from math import fabs
import os
from wox import Wox, WoxAPI
import subprocess   
import json

PackageDir = os.path.join(os.environ['localAppData'.upper()],'Packages')

HistoryFilePath = './History.json'

class ShellWT(Wox):
    @classmethod
    def query(cls, queryString):
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
                result.append({
                    'Title': data,
                    'SubTitle': 'Has been run ' + json.dumps(history[data]) + ' times before.',
                    'IcoPath': icon,
                    'ContextData': data,
                    'JsonRPCAction': {
                        'method': 'Run',
                        'parameters': [data,False],
                        "doNotHideAfterAction".replace('oNo', 'on'): False
                    }
                })
                if queryString == data:
                    foundInHistory = True

        if not foundInHistory:
             result.append({
                'Title': queryString,
                'SubTitle': 'Press Enter to Run',
                'IcoPath': icon,
                'ContextData': queryString,
                'JsonRPCAction': {
                    'method': 'Run',
                    'parameters': [queryString,False],
                    "doNotHideAfterAction".replace('oNo', 'on'): False
                }
            })

        runPowerShell = {
            'Title': 'Run Windows PowerShell' ,
            'SubTitle': 'Press Enter to Run',
            'IcoPath': icon,
            'ContextData': '>>PowerShell',
            'JsonRPCAction': {
                'method': 'RunPowerShell',
                'parameters': [False],
                "doNotHideAfterAction".replace('oNo', 'on'): False
            }
        }
        runCMD = {
            'Title': 'Run CMD' ,
            'SubTitle': 'Press Enter to Run',
            'IcoPath': icon,
            'ContextData': '>>CMD',
            'JsonRPCAction': {
                'method': 'RunCMD',
                'parameters': [False],
                "doNotHideAfterAction".replace('oNo', 'on'): False
            }
        }
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
        if queryString == '>>PowerShell':
            return [{
                'Title': 'Run Windows PowerShell as Administrator' ,
                'SubTitle': 'Press Enter to Run as Administrator',
                'IcoPath': iconPath,
                'JsonRPCAction': {
                    'method': 'RunPowerShell',
                    'parameters': [True],
                    "doNotHideAfterAction".replace('oNo', 'on'): False
                }
            }]
        elif queryString == '>>CMD':
            return  [{
                'Title': 'Run CMD as Administrator' ,
                'SubTitle': 'Press Enter to Run as Administrator',
                'IcoPath': iconPath,
                'JsonRPCAction': {
                    'method': 'RunCMD',
                    'parameters': [True],
                    "doNotHideAfterAction".replace('oNo', 'on'): False
                }
            }]
        return  [{
            'Title': queryString,
            'SubTitle': 'Press Enter to Run as Administrator',
            'IcoPath': iconPath,
            'JsonRPCAction': {
                'method': 'Run',
                'parameters': [queryString,True],
                "doNotHideAfterAction".replace('oNo', 'on'): False
            }
        }]

    @classmethod
    def Run(cls, cmd, isRunAsAdministrator):
        if isRunAsAdministrator:
            subprocess.Popen("sudo -n wt powershell -NoExit " + cmd + "", cwd=os.getenv("SystemRoot")+"\system32")
        else:
            subprocess.Popen('wt powershell -NoExit ' + cmd, cwd=os.getenv("SystemRoot")+"\system32")
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
