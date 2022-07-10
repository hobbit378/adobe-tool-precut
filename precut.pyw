#!/usr/bin/env python

#
#############################################################################
#
# precuct.py - GUI for precut
# Copyright (C) 2015, testcams.com
#
# This module is licensed under GPL v3: http://www.gnu.org/licenses/gpl-3.0.html
#
#
#############################################################################
#

#
# note001 - The Tkinter implementation on OSX has a behavior/bug where it doesn't vertically
# center the image specified for a button if the button also has text. (it horizontally centers
# but not vertically). This makes the button look ugly for ipady because all the padding
# goes to the bottom instead of being evenly distributed. I work arond this by massaging
# the padding when we're running on iOSX
#

from __future__ import print_function
from __future__ import division
#
# six.py's remapping works as intended under Python 2.x but PyInstaller doesn't know how
# handle its remapping yet, so when running on Python 2.x I  always use the standard
# Python 2.x imports, just in case I'm performing a PyInstaller build, which I happen
# to doing under Python 2.x.
#
import six
if six.PY2:
    from Tkinter import *
    from ScrolledText import *
    from tkFont import *
    import ttk
    import tkFileDialog
    import tkMessageBox
else:
    from six.moves.tkinter import *
    from six.moves.tkinter_font import *
    from six.moves.tkinter_scrolledtext import *
    from six.moves import tkinter_ttk as ttk
    from six.moves import tkinter_tkfiledialog as tkFileDialog
    from six.moves import tkinter_messagebox as tkMessageBox
import time
import subprocess
import os
import errno
import platform
import json
import datetime
import signal
import sys
from collections import *

#
# App name and version
#
APP_NAME = "precut"
APP_NAME_GROUP = "precut"   # used only for establishing app data folder name
APP_VERSION = "0.91"


class Colors(object): # UI colors
    mainBgColor = "#E0E0E0"
    toolbarColor = "#B0B0B0"


class CmdArgs(object):  # user-readable options -> precutcmd command line dictionaries
    LoggingLevelChoicesDict = OrderedDict([
        ('quiet',   '--logginglevel quiet'),
        ('minimal', '--logginglevel minimal'),
        ('normal',  '--logginglevel normal'),
        ('warning', '--logginglevel warning'),
        ('verbose', '--logginglevel verbose'),
        ('debug',   '--logginglevel debug' )
    ])
    ClipCombineChoicesDict = OrderedDict([
        ('By Source File',                  '--combine sourcemedia'),
        ('By Track',                        '--combine track'),
        ('By Sequence',                     '--combine sequence'),
        ('None - Store Individual Clips',   '--combine none')
    ])
    FileExistsChoicesDict = OrderedDict([
        ('generate unique filename',    '--ifexists uniquename'),
        ('overwrite file',              '--ifexists overwrite'),
        ('skip file',                   '--ifexists skip'),
        ('prompt for each file',        '--ifexists prompt'),
        ('exit app',                    '--ifexists exit')
    ])
    FfmpegLoggingChoicesDict = OrderedDict([
        ('none',                  '--ffmpegreportloglevel none'),
        ('quiet (-8)',            '--ffmpegreportloglevel -8'),
        ('panic (0)',             '--ffmpegreportloglevel 0'),
        ('fatal (8)',             '--ffmpegreportloglevel 8'),
        ('error (16)',            '--ffmpegreportloglevel 16'),
        ('warning (24)',          '--ffmpegreportloglevel 24'),
        ('info (32)',             '--ffmpegreportloglevel 32'),
        ('verbose (40)',          '--ffmpegreportloglevel 40'),
        ('debug (48)',            '--ffmpegreportloglevel 48'),
        ('trace (56)',            '--ffmpegreportloglevel 56')
    ])


class GlobalVarsStruct(object):
    def __init__(self):
        self.isWin32 = None         # True if we're running on a Windows platform
        self.isOSX = None           # True if we're runnong on an OSX platform
        self.appDir = None          # directory where script is located. this path is used to store all metadata files, in case script is run in different working directory
        self.appDataDir = None      # directory where we keep app metadata
        self.appResourceDir = None  # directory where we keep app resources (read-only files needed by app, self.appDir + "resouce")
        self.app = None             # reference to main Application class
        self.appConfig = None       # instance of ApplicationConfig
        self.quickTip = None        # instance of QuickTip


#
# global variables
#
root = None                         # Tk root
g = GlobalVarsStruct()


#
# determines if the specified string contains non-whitespace chars.
# Returns FALSE if the string is None or zero-length or contains
# all whitespace, otherwise returns true
#
def isStrValidWithNonWhitespaceChars(str):
    return str != None and len(str) > 0 and not str.isspace()


#
# creates a list of arguments from an argument string, honoring quoted args as a single argument
#
def createArgListFromArgStr(argStr):
    return [x.strip('"') for x in re.split('( |".*?")', argStr) if x.strip()]


#
# generates a list composed of the keys from a dictonary
#
def dictionaryKeysToList(dict):
    # generate list of keys by using list comprehension to iterate keys of dictionary
    return [a for a in dict.keys()]


#
# returns the path to the current user's directory. if that directory can't be determined
# then we return the current working directory instead
#
def getDefaultUserDir():

    dir = os.getcwd() # default to last resort of current working directory, which is usually our application

    if g.isWin32:
        if os.environ['HOMEDRIVE'] and os.environ['HOMEPATH']:
            # user files path available
            dir = os.path.join(os.environ['HOMEDRIVE'], os.environ['HOMEPATH'])
    else:
        if os.environ['HOME']:
            # user files path available
            dir = os.environ['HOME']

    return os.path.abspath(dir)


#
# verifies user is running version a modern-enough version of python for this app
#
def verifyPythonVersion():
    if sys.version_info.major == 2:
        if sys.version_info.minor < 7:
            print("Warning: You are running a Python 2.x version older than app was tested with.")
            print("Version running is {:d}.{:d}.{:d}, app was tested on 2.7.x".format(sys.version_info.major, sys.version_info.minor, sys.version_info.micro))
    elif sys.version_info.major == 3:
        if sys.version_info.minor < 4:
            print("Warning: You are running a Python 3.x version older than app was tested with.")
            print("Version running is {:d}.{:d}.{:d}, app was tested on 3.4.x".format(sys.version_info.major, sys.version_info.minor, sys.version_info.micro))


#
# sets app-level globals related to the platform we're running under and
# creates path to app directories, creating them if necessary
#
def establishAppEnvironment():

    g.isWin32 = (platform.system() == 'Windows')
    g.isOSX = (platform.system() == 'Darwin')

    #
    # determine the directory our script resides in, in case the
    # user is executing from a different working directory.
    #
    g.appDir = os.path.dirname(os.path.realpath(sys.argv[0]))
    g.appResourceDir = os.path.join(g.appDir, "appresource")

    #
    # determine directory to store app data to, such as log files
    #
    g.appDataDir = None
    appDataDirFromEnv = os.getenv(APP_NAME + '_appdatadir')
    if appDataDirFromEnv:
        # user specified log file directory in environmental variable - use it
        g.appDataDir = appDataDirFromEnv
    else:
        # user didn't specify log file directory - select one based on platform
        if g.isWin32:
            if os.getenv('LOCALAPPDATA'):
                g.appDataDir = os.path.join(os.getenv('LOCALAPPDATA'), APP_NAME_GROUP) # typically C:\Users\<username>\AppData\Local\<appname>
        elif g.isOSX: # for OSX we always try to store our app data under Application Support
            userHomeDir = os.getenv('HOME')
            if userHomeDir:
                applicationSupportDir = os.path.join(userHomeDir, 'Library/Application Support')
                if os.path.exists(applicationSupportDir): # probably not necessary to check existence since every system should have this directory
                    g.appDataDir = os.path.join(applicationSupportDir, APP_NAME_GROUP) # typically /Users/<username>/Library/Application Support/<appname>
    if not g.appDataDir:
        # none of runtime-specific cases above selected an app data directory - use same directory as app file (and hope that it's writeable)
        g.appDataDir = g.appDir
    # create our app data directory if necessary
    if not os.path.exists(g.appDataDir):
        os.makedirs(g.appDataDir)


#
# brings our app window(s) to the front of the OS's window z-order
#
def bringAppToFront():
    # fix for bringing our app to front on OS X
    root.lift()
    root.call('wm', 'attributes', '.', '-topmost', True)
    root.after_idle(root.call, 'wm', 'attributes', '.', '-topmost', False)


#
# sets the icon for the given window/frame
#
def setFrameIcon(frame):
    if g.isWin32:
        frame.wm_iconbitmap(bitmap = os.path.join(g.appResourceDir, 'precut.ico'))
    else:
        frame.wm_iconbitmap(bitmap = '@' + os.path.join(g.appResourceDir, 'precut.xbm'))


#
# spawns precutcmd with specified arguments. returns errno from precutcmd
#
def launchPrecut(argStr):

    #
    # launch precutcmd and wait for its comletion
    #
    print("Launching precutcmd with args: ", argStr)
    try:

        root.withdraw() # make our main GUI window invisible while running precutcmd

        process = None
        argList = createArgListFromArgStr(argStr)
        process = subprocess.Popen(['python', os.path.join(g.appDir, 'precutcmd.py')] + argList)

        if process:
            _errno = process.wait()

    except KeyboardInterrupt as e:
        print("SIGINT received while waiting for precutcmd to complete")
        if process:
            #
            # both precut and precutcmd receive the SIGINT. make sure precutcmd
            # has completed processing its SIGTERM by waiting for it to exit
            #
            if process.poll() == None:
                print("waiting for precutcmd to finish handling its SIGINT")
                process.wait()
        _errno = errno.EINTR
    finally:
        root.deiconify() # bring our main GUI window back

    #
    # display precutcmd's log
    #
    displayPrecutLog()
    return _errno


#
# displays last output from precutcmd inside a new top-level window
#
def displayPrecutLog():

    #
    # load report
    #
    lastLogFilename = os.path.join(g.appDataDir, "precutcmd-log-last.txt")
    if not os.path.exists(lastLogFilename):
        tkMessageBox.showwarning("precutcmd log", "No precutcmd operations have been performed yet")
        return
    fileReport = None
    try:
        fileReport = open(lastLogFilename, "r")
        reportContents = fileReport.read()
    except IOError as e:
        tkMessageBox.showerror("precutcmd log", "Error reading precutcmd report file \"{:s}\". {:s}".format(lastLogFilename, str(e)))
        return
    finally:
        if fileReport:
            fileReport.close()

    #
    # detect an unclean shutdown of precutcmd by checking if the logfile is either
    # empty (no data flushed before termation) or is missing the final "session over"
    # message (last set of message(s) not flushed).
    #
    if reportContents:
        # file not empty
        posSessionOverMessage = reportContents.find(">>>> precutcmd session over")
    else:
        posSessionOverMessage = -1
    if posSessionOverMessage == -1:
        g.quickTip.show('precutcmd_unclean_shutdown', 1,
            "It appears precutcmd was uncleanly terminated. In the future please press <ctrl-c> if "\
            "you'd like to terminate precutcmd instead of closing its terminal window. This will "\
            "allow precutcmd to perform any necessary cleanup prior to exiting.")
        if not reportContents:
            tkMessageBox.showwarning("precutcmd log", "Log is empty")
            return
    reportContents = reportContents[:posSessionOverMessage] # trim off ">>>> precut session over ...." message

    #
    # create top-level window
    #
    topLevelFrame = Toplevel(root)
    topLevelFrame.geometry('900x420')
    topLevelFrame.title("precutcmd log")
    setFrameIcon(topLevelFrame)

    #
    # fill window with text control to hold report and some buttons
    #
    scrolledText = ScrolledText(topLevelFrame, bg='yellow', width=80)
    # text widgets can't be set read-only without disabling them, so achieve the same by disabling all keypresses except ctrl-c (to allow copy)
    scrolledText.bind("<Control-c>", lambda e : "")     # on ctrl-c, invoke function that returns empty string, allowing the default ahndler to process the ctrl-c copy operations
    scrolledText.bind('<Key>', lambda e: "break")       # on all other keys, invoke function that returns "break", preventing the keypress from being handled

    buttonFrame = Frame(topLevelFrame)
    button = Button(buttonFrame, text="Ok", command=lambda : topLevelFrame.destroy())
    button.grid(column=0, row=0, padx=10, pady=5, ipadx=40, ipady=5)
    button.focus_set()
    button = Button(buttonFrame, text="Copy to Clipboard", command=lambda : [root.clipboard_clear(), root.clipboard_append(reportContents)])
    button.grid(column=2, row=0, padx=80, pady=5, ipady=5, ipadx=8)
    buttonFrame.pack(side=BOTTOM)
    scrolledText.pack(side=TOP, fill=BOTH, expand=1) # packed last to give other controls real estate in frame

    # insert report into text control
    scrolledText.insert(END, reportContents)
    scrolledText.see(END)      # move cursor to end of log, since that likely has the most interesting info if the operation failed

    #
    # present top-level window as modal and wait for it to be dismissed
    #
    topLevelFrame.transient(root)
    topLevelFrame.grab_set()
    bringAppToFront()
    root.wait_window(topLevelFrame)

    if g.isOSX:
        g.quickTip.show('osx_terminal_auto_close', 4,
            "precut performs its job by launching precutcmd in a terminal window. By default OSX "\
            "keeps that terminal window open even after precutcmd has completed, requiring you to "\
            "to close it manually. You can configure OSX to close the terminal window automatically. "\
            "Go to the Terminal application and select 'Preferences' in the Terminal menu. " \
            "Click 'Profiles' at the top and then the \"Shell\" tab in the upper part of the window. "\
            "Set \"When the shell exits\" option to 'Close if the shell exited cleanly' or 'Close "\
            "the window'")


#
# called after the creation of a new combobox, where 'valuesForCombo' is the history
# from appconfig and contains the list of items to populate the combo with. we allow
# the first item an appconfig's history element to be blank, signifying the user's
# choice of no-entry. however we don't allow blank entries to be stored in the combobox
# because that can be confusing. the purpose of this routine is to strip off the first
# item if it's blank - then setting the combobox's current value to empty to achieve
#
def setComboValuesFromList_RemoveBlankEntryIfNecessary(comboBox, valuesForCombo):
    if valuesForCombo:
        if isStrValidWithNonWhitespaceChars(valuesForCombo[0]):
            # top of history is a non-blank entry - use history as-is for combobox's list
            comboBox['values'] = valuesForCombo
            comboBox.current(0)
        else:
            #
            # top of history is a blank entry - don't include blank entry in combobox's
            # list - just set the current text of the combobox to blank
            #
            comboBox['values'] = valuesForCombo[1:]
            comboBox.set("")


#
# class for creating a combo box with pre-defined, non-editable choices
# and label describing the combo box
#
class ComboBoxWithLabel(object):

    def __init__(self, parentFrame, textDescription, valueList, configDictKey, defaultIfNotInDict=None, bgColor=Colors.mainBgColor):

        label = Label(parentFrame, text=textDescription + ':', bg=bgColor)
        comboBox = ttk.Combobox(parentFrame, values=valueList, state='readonly')
        if configDictKey in g.appConfig.dict:
            comboBox.current(comboBox['values'].index(g.appConfig.dict[configDictKey]))
        elif defaultIfNotInDict:
            comboBox.current(comboBox['values'].index(defaultIfNotInDict))
        else:
            comboBox.current(0)

        self.configDictKey = configDictKey
        self.label = label
        self.comboBox = comboBox


#
# class for creating a editable combo box whose contents are determined by
# the app config dictionary and label describing the combo box
#
class EditableComboBoxWithLabel(object):

    def __init__(self, parentFrame, textDescription, configDictKey, defaultValuesIfNotInDict=None):

        label = Label(parentFrame, text=textDescription + ':', bg=Colors.mainBgColor)
        comboBox = ttk.Combobox(parentFrame, state='normal')

        if configDictKey in g.appConfig.dict:
            valuesForCombo = g.appConfig.dict[configDictKey]
        elif defaultValuesIfNotInDict:
            valuesForCombo = defaultValuesIfNotInDict
        else:
            valuesForCombo = None;
        setComboValuesFromList_RemoveBlankEntryIfNecessary(comboBox, valuesForCombo)

        comboBox.bind("<FocusIn>", lambda event: self.comboBoxFocusIn(event))
        comboBox.bind("<FocusOut>", lambda event: self.comboBoxFocusOut(event))

        self.configDictKey = configDictKey
        self.defaultValuesIfNotInDict = defaultValuesIfNotInDict
        self.label = label
        self.comboBox = comboBox

    def comboBoxFocusIn(self, event):
        #
        # if there is a pseudo-value as the current value, select the entire value
        # to make it easy to type and replace. note TKinter automatically selects
        # the entire value when the user navigates to control via TAB - this logic
        # is for when the mouse is used to select the field instead
        #
        currentValue = self.comboBox.get()
        if currentValue and currentValue.find("<<") != -1:
            self.comboBox.selection_range(0, len(currentValue))

    def comboBoxFocusOut(self, event):
        self.comboBox.selection_range(0, 0)


#
# class for presenting and managing UI controls that allow the user to select a
# directory or filename. The three controls are a text label, a combobox, and a button.
# The label describes what the directory or filename is for. The combobox contains a
# list of directories/files, and the button is used to bring up a folder/file selector
# dialog to add to the list in the combobox.
#
class PickPathOrFileControls(object):

    CONTROL_FLAG_DIRECTORY      =   (0x00000001<<0)
    CONTROL_FLAG_LOAD_FILE      =   (0x00000001<<1)
    CONTROL_FLAG_SAVE_FILE      =   (0x00000001<<2)
    CONTROL_FLAG_EDITABLE       =   (0x00000001<<3)
    # internal clase use below
    CONTROL_FLAG_FILE           =   (CONTROL_FLAG_LOAD_FILE | CONTROL_FLAG_SAVE_FILE)

    __DefaultFileValue = "<< No File Chosen >>"

    def __init__(self, parentFrame, textPathOrFileDescription, controlFlags, configDictKey, nonPathDefaultValues=None):

        label = Label(parentFrame, text=textPathOrFileDescription + ':', bg=Colors.mainBgColor)
        comboBox = ttk.Combobox(parentFrame, state='normal' if controlFlags & PickPathOrFileControls.CONTROL_FLAG_EDITABLE else 'readonly')

        if configDictKey in g.appConfig.dict:
            valuesForCombo = g.appConfig.dict[configDictKey]
        elif nonPathDefaultValues:
            valuesForCombo = nonPathDefaultValues
        elif controlFlags & PickPathOrFileControls.CONTROL_FLAG_DIRECTORY:
            valuesForCombo = [ getDefaultUserDir() ]
        else:
            valuesForCombo = None
        setComboValuesFromList_RemoveBlankEntryIfNecessary(comboBox, valuesForCombo)

        comboBox.bind("<FocusIn>", lambda event: self.comboBoxFocusIn(event))
        comboBox.bind("<FocusOut>", lambda event: self.comboBoxFocusOut(event))

        button = Button(parentFrame, text="Select Directory" if controlFlags & PickPathOrFileControls.CONTROL_FLAG_DIRECTORY else "Select File", command=lambda : self.buttonClick())

        self.nonPathDefaultValues = nonPathDefaultValues
        self.configDictKey = configDictKey
        self.controlFlags = controlFlags
        self.textPathOrFileDescription = textPathOrFileDescription
        self.label = label
        self.comboBox = comboBox
        self.button = button

    def comboBoxFocusIn(self, event):
        #
        # if there is a pseudo-value as the current value, select the entire value
        # to make it easy to type and replace. note TKinter automatically selects
        # the entire value when the user navigates to control via TAB - this logic
        # is for when the mouse is used to select the field instead
        #
        currentValue = self.comboBox.get()
        if currentValue and currentValue.find("<<") != -1:
            self.comboBox.selection_range(0, len(currentValue))

    def comboBoxFocusOut(self, event):
        self.comboBox.selection_range(0, 0)

    def buttonClick(self):

        fComboOnlyHasPlaceholderValue = False

        currentComboValue = self.comboBox.get()
        fCurrentComboValueIsPseudoValue = (currentComboValue[0:2].find("<<") != -1)

        # open dialog to get path or filename
        if self.controlFlags & PickPathOrFileControls.CONTROL_FLAG_DIRECTORY:

            if not fCurrentComboValueIsPseudoValue:
                initialDir = currentComboValue
            else:   # current selection is a pseudo option - pick user dir as default
                initialDir = getDefaultUserDir()

            dir_opt = { 'initialdir'    : initialDir,
                        'mustexist'     : True,
                        'title'         : self.textPathOrFileDescription,
            }
            newValue = tkFileDialog.askdirectory(**dir_opt)
        else:
            if fCurrentComboValueIsPseudoValue or not currentComboValue:
                initialDir = getDefaultUserDir()
                initialFile = ""
            else:
                if os.path.isdir(currentComboValue):
                    initialDir = currentComboValue
                else:
                    (initialDir, initialFile) = os.path.split(currentComboValue)

            file_opt = { 'initialdir'       : initialDir,
                         'initialfile'      : initialFile,
                         'title'            : self.textPathOrFileDescription,
                         'defaultextension' : '.xml',
                         'filetypes'        : [('xml files', '.xml')]
            }
            if self.controlFlags & PickPathOrFileControls.CONTROL_FLAG_LOAD_FILE:
                newValue = tkFileDialog.askopenfilename(**file_opt)
            else:
                newValue = tkFileDialog.asksaveasfilename(**file_opt)

        # process selection
        if newValue:
            if g.isWin32:
                # ask***() converts paths to unix style. while these work on Windows,
                # they're confusing to see so convert it back
                newValue = newValue.replace('/', '\\')

            comboBoxValuesList = list(self.comboBox['values'])       # combobox maintains list as tuples - convert to list for easier manipulation
            if newValue not in comboBoxValuesList:
                #
                # user selected a value that's not already in the combo list
                # add it as the first element of the list, and set that as
                # selected element
                #
                comboBoxValuesList.insert(0, newValue)
                self.comboBox['values'] = comboBoxValuesList    # combobox automatically converts list to tuples on assignment
                self.comboBox.current(0)
            else:
                # user selected a value that's already in the combo list. make that the current selection
                self.comboBox.current(comboBoxValuesList.index(newValue))


#
# class for managing an Application's configuration, including loading/storing
# the configuration from the file
#
class AppConfig(object):

    def __init__(self, appConfigFilename):
        self.dict = None
        self.filename = appConfigFilename
        self.loadAppConfig()

    def loadAppConfig(self):
        try:
            if os.path.exists(self.filename):
                f = open(self.filename, "r")
                self.dict  = json.loads(f.read())
                f.close()
        except IOError as e:
            tkMessageBox.showwarning("Loading App Config", "Could not read app config data at \"{:s}\". {:s}. Defaults will be used.".format(self.filename, str(e)))
        except ValueError as e:
            tkMessageBox.showwarning("Loading App Config", "Could not decode app config data at \"{:s}\". {:s}. Defaults will be used.".format(self.filename, str(e)))
        if not self.dict:
            self.dict = {}

    def saveAppConfig(self):
        try:
            f = open(self.filename, "w")
            f.write(json.dumps(self.dict))
            f.close()
        except IOError as e:
            tkMessageBox.showwarning("Saving App Config", "Could not write app config data. {:s}.".format(str(e)))
        except ValueError as e:
            tkMessageBox.showwarning("Saving App Config", "Could not encode  app config data. {:s}.")


#
# class for managing "quick tips", which are tips presented to the user over the lifetime
# of the application
#
class QuickTip(object):

    def __init__(self, appConfigInst):
        self.appConfig = appConfigInst

    def show(self, tipDictRef, numOccurrencesBeforePresentingTip, tipStr):

        #
        # this routine is used to present the user a tip/warning about using the app.
        # it is called from various places within the module based on where the app
        # determines the tip would be useful to know. 'tipDictRef' is only tag used
        # to track this particular tip. 'tipStr' is the contents of the tip itself.
        # 'numOccurrencesBeforePresentingTip' establishes how many times the tip
        # should be evaluated before presenting it to the user; for example, if the tip
        # is to inform the user about a faster way to accomplish an action, we might
        # wait until the 5th time the user performs the action before presenting the
        # tip. That way the user gets to use the program/action for a bit before being
        # bombarded with a bunch of tips. once the occurence threshold has been reached
        # the tip will be presented to the user and the record of the presentation will
        # be saved in self.appConfig.dict so that the tip is never presented again
        #

        if 'quick_tips' in self.appConfig.dict:
            # quick_tips dictionary exists - retrieve it
            quickTipsDict = self.appConfig.dict['quick_tips']
        else:
            # this is the first quick tip check we're performing - create dictionary
            quickTipsDict = {}
        if tipDictRef in quickTipsDict:
            # we've previously evaluated this tip before
            if quickTipsDict[tipDictRef] >= numOccurrencesBeforePresentingTip:
                #  tip previously reached its occurence threshold and has already been presented to user
                return
            quickTipsDict[tipDictRef] = quickTipsDict[tipDictRef] + 1 # increase evaluation count for this tip
        else:
            # this is the first time we're evaluating this tip for presentation
            quickTipsDict[tipDictRef] = 1

        # we've updated the quick-tips dictionary - save it
        self.appConfig.dict['quick_tips'] = quickTipsDict
        self.appConfig.saveAppConfig()

        # present the tip to the user if we've reached its threshold
        if quickTipsDict[tipDictRef] >= numOccurrencesBeforePresentingTip:
            tkMessageBox.showinfo("precut Quick Tip", tipStr)


#
# main application Tkinter class
#
class Application(Frame):

    __DefaultOutputDir = "<< default for combine method >>"
    __DefaultFileContainer = "<< same as source >>"
    __DefaultOutputNameSpec = "<< default >>"
    __DefaultPrecutcmdVideoOrArgs = "<< precut defaults >>"
    __DefaultFfmpegVideoOrAudioArgs = "<< none - ffmpeg defaults >>"
    __DefaultFfmpegDir = "<< assume in system path >>"

    #
    # root -> Application(Frame)    -> Menu
    #                               -> ToolBar
    #                               -> ContentFrame
    #

    def __init__(self, master):

        Frame.__init__(self, master, bg=Colors.mainBgColor)

        root.title("precut - Cuts video files from edits in Premiere Project -> FCP7 Export")
        setFrameIcon(root)

        # resources
        self.tkLoadedResourcesDict = {}

        #
        # load config from previous session(s). this contains previous user choices, which
        # will be the defaults for each repesective wizard/form element.
        #
        g.appConfig = AppConfig(os.path.join(g.appDataDir, "precut-gui-config"))
        g.quickTip = QuickTip(g.appConfig)

        # top-level menu
        self.createTopLevelMenu()

        # toolbar
        self.toolbarFrame = Frame(self, bg=Colors.toolbarColor)

        # toolbar - Last Operation Log button
        button = Button(self.toolbarFrame, text="Last Operation Log", command=lambda : self.toolbarClick('display_last_log'))
        button.pack(side=LEFT, padx=10, pady=5, ipadx=10)

        optionsFrame = Frame(self.toolbarFrame, bg=Colors.toolbarColor)

        # toolbar - precutcmd logging level
        loggingControls = ComboBoxWithLabel(optionsFrame, "Logging", dictionaryKeysToList(CmdArgs.LoggingLevelChoicesDict), 'logging_choice', defaultIfNotInDict="normal", bgColor=Colors.toolbarColor)
        loggingControls.label.grid(column=4, row=0, sticky=E, padx=2)
        loggingControls.comboBox.grid(column=5, row=0, sticky=W, padx=10)
        self.loggingControls = loggingControls

        # toolbar - ffmpeg logging level
        ffmpegLoggingControls = ComboBoxWithLabel(optionsFrame, "ffmpeg Logging", dictionaryKeysToList(CmdArgs.FfmpegLoggingChoicesDict), 'ffmpeglogging_choice', defaultIfNotInDict="warning (24)", bgColor=Colors.toolbarColor)
        ffmpegLoggingControls.label.grid(column=6, row=0, sticky=E, padx=2)
        ffmpegLoggingControls.comboBox.grid(column=7, row=0, sticky=W, padx=10)
        self.ffmpegLoggingControls = ffmpegLoggingControls

        optionsFrame.pack(side=LEFT, padx=0, pady=5)
        self.toolbarFrame.pack(side=TOP, fill=X)

        #
        # main window area
        #
        self.contentFrame = Frame(self, bg=Colors.mainBgColor)

        controlsFrame = Frame(self.contentFrame, bg=Colors.mainBgColor, padx=15, pady=15)

        # XML file source controls
        row = 0
        xmlFileControls = PickPathOrFileControls(controlsFrame, "Final Cut Pro 7 XML File", PickPathOrFileControls.CONTROL_FLAG_LOAD_FILE | PickPathOrFileControls.CONTROL_FLAG_EDITABLE, "filesource_history")
        xmlFileControls.label.grid(column=0, row=row, sticky=E, pady=5)
        xmlFileControls.comboBox.grid(column=1, row=row, sticky=W, ipadx=80)
        xmlFileControls.button.grid(column=2, row=row, sticky=W, pady=5,  padx=5)
        xmlFileControls.comboBox.focus_set()
        self.xmlFileControls = xmlFileControls
        row += 1

        # output directory
        outputDirControls = PickPathOrFileControls(controlsFrame, "Output Directory", PickPathOrFileControls.CONTROL_FLAG_DIRECTORY  | PickPathOrFileControls.CONTROL_FLAG_EDITABLE, "outputdir_history", [ Application.__DefaultOutputDir ])
        outputDirControls.label.grid(column=0, row=row, sticky=E, pady=5)
        outputDirControls.comboBox.grid(column=1, row=row, sticky=W, ipadx=80)
        outputDirControls.button.grid(column=2, row=row, sticky=W, pady=5,  padx=5)
        outputDirControls.comboBox.focus_set()
        self.outputDirControls = outputDirControls
        row += 1

        # combine edits selection
        videoCombineControls = ComboBoxWithLabel(controlsFrame, "Combine Clip Edits", dictionaryKeysToList(CmdArgs.ClipCombineChoicesDict), 'combineclips_choice')
        videoCombineControls.label.grid(column=0, row=row, sticky=E, pady=5)
        videoCombineControls.comboBox.grid(column=1, row=row, sticky=W, ipadx=80)
        self.videoCombineControls = videoCombineControls
        row += 1

        # if file exists selection
        fileExistsControls = ComboBoxWithLabel(controlsFrame, "If Output File(s) Exist", dictionaryKeysToList(CmdArgs.FileExistsChoicesDict), 'if_file_exists_choice')
        fileExistsControls.label.grid(column=0, row=row, sticky=E, pady=5)
        fileExistsControls.comboBox.grid(column=1, row=row, sticky=W, ipadx=80)
        self.fileExistsControls = fileExistsControls
        row += 1

        # container
        fileContainerControls = EditableComboBoxWithLabel(controlsFrame, "Video Container (Extension)", 'container_history', [ Application.__DefaultFileContainer ])
        labelContainerExample = Label(controlsFrame, text="(ex: MP4, MOV, etc...)", bg=Colors.mainBgColor)
        fileContainerControls.label.grid(column=0, row=row, sticky=E, pady=5)
        fileContainerControls.comboBox.grid(column=1, row=row, sticky=W, ipadx=80)
        labelContainerExample.grid(column=2, row=row, sticky=W, pady=5)
        self.fileContainerControls = fileContainerControls
        row += 1

        # filename spec
        outputNameSpecControls = EditableComboBoxWithLabel(controlsFrame, "Filename spec", 'outputnamespec_history', [ Application.__DefaultOutputNameSpec ])
        outputNameSpecControls.label.grid(column=0, row=row, sticky=E, pady=5)
        outputNameSpecControls.comboBox.grid(column=1, row=row, sticky=W, ipadx=80)
        self.outputNameSpecControls = outputNameSpecControls
        row += 1

        # video args
        videoArgsControls = EditableComboBoxWithLabel(controlsFrame, "Video Args (ffmpeg)", 'videoargs_history', [ '-c:v copy -copyinkf -avoid_negative_ts 1 -copyts', Application.__DefaultFfmpegVideoOrAudioArgs ])
        videoArgsControls.label.grid(column=0, row=row, sticky=E, pady=5)
        videoArgsControls.comboBox.grid(column=1, row=row, sticky=W, ipadx=80)
        self.videoArgsControls = videoArgsControls
        row += 1

        # audio args
        audioArgsControls = EditableComboBoxWithLabel(controlsFrame, "Audio Args (ffmpeg)", 'audioargs_history', [ '-c:a copy', Application.__DefaultFfmpegVideoOrAudioArgs ])
        audioArgsControls.label.grid(column=0, row=row, sticky=E, pady=5)
        audioArgsControls.comboBox.grid(column=1, row=row, sticky=W, ipadx=80)
        self.audioArgsControls = audioArgsControls
        row += 1

        # additional ffmpeg args
        additionalFfmpegArgsControls = EditableComboBoxWithLabel(controlsFrame, "Additional Args (ffmpeg)", 'additionalffmpegargs_history', None)
        additionalFfmpegArgsControls.label.grid(column=0, row=row, sticky=E, pady=5)
        additionalFfmpegArgsControls.comboBox.grid(column=1, row=row, sticky=W, ipadx=80)
        self.additionalFfmpegArgsControls = additionalFfmpegArgsControls
        row += 1

        # additional precut args
        additionalPrecutArgsControls = EditableComboBoxWithLabel(controlsFrame, "Additional Args (precutcmd)", 'additionalprecutargs_history', None)
        additionalPrecutArgsControls.label.grid(column=0, row=row, sticky=E, pady=5)
        additionalPrecutArgsControls.comboBox.grid(column=1, row=row, sticky=W, ipadx=80)
        self.additionalPrecutArgsControls = additionalPrecutArgsControls
        row += 1

        # ffmpeg directory
        ffmpegDirControls = PickPathOrFileControls(controlsFrame, "Directory of ffmpeg executable", PickPathOrFileControls.CONTROL_FLAG_DIRECTORY  | PickPathOrFileControls.CONTROL_FLAG_EDITABLE, "ffmpegdir_history", [ Application.__DefaultFfmpegDir ])
        ffmpegDirControls.label.grid(column=0, row=row, sticky=E, pady=5)
        ffmpegDirControls.comboBox.grid(column=1, row=row, sticky=W, ipadx=80)
        ffmpegDirControls.button.grid(column=2, row=row, sticky=W, pady=5,  padx=5)
        ffmpegDirControls.comboBox.focus_set()
        self.ffmpegDirControls = ffmpegDirControls
        row += 1

        # run button
        button = Button(controlsFrame, image=self.getResource_Image("gobutton.gif"), compound=TOP, text="Run", bg=Colors.mainBgColor, command=lambda : self.runButtonClick())
        button.grid(column=1, row=row, sticky=W, ipadx=50, ipady=5 if not g.isOSX else 0, pady=20, padx=40) # note001

        # pack the frames
        controlsFrame.pack(side=LEFT, fill=BOTH, expand=1)
        self.contentFrame.pack(side=LEFT, fill=BOTH, expand=1)
        self.pack(fill=BOTH, expand=1)

        bringAppToFront()

    def createTopLevelMenu(self):

        menubar = Menu(root)

        # file menu
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="Exit", command=self.wmDeleteWindow)
        menubar.add_cascade(label="File", menu=filemenu)

        # help menu
        helpmenu = Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About", command=lambda : tkMessageBox.showinfo("precut", \
            "precut - Version {:s}\n\nRunning under Python Version {:d}.{:d}.{:d}\n\nApplication is licensed under GPL v3\n\n"\
            "Movie icon courtesy of Mateusz Piotrowski, from https://www.iconfinder.com/icons/1167956, licensed under Creative Commons 2.5 (https://creativecommons.org/licenses/by/2.5)\n\n"\
            "Green arrow icon courtesy of ricardomaia, from https://openclipart.org/detail/122407\n\n"\
            .format(APP_VERSION, sys.version_info.major, sys.version_info.minor, sys.version_info.micro)))
        menubar.add_cascade(label="Help", menu=helpmenu)

        root.config(menu=menubar)

    def toolbarClick(self, str):
        if str == 'display_last_log':
            displayPrecutLog()

    def getResource_Image(self, filename):

        #
        # gets photo resource. returns the resource if it's previously been loaded,
        # otherwise we load it. we keep the reference in a dictionary, as required
        # because tKinter doesn't keep a reference itself and so without our
        # reference the image would be garbage collected
        #
        if filename in self.tkLoadedResourcesDict:
            return self.tkLoadedResourcesDict[filename]
        photo = PhotoImage(file = os.path.join(g.appResourceDir, filename))
        self.tkLoadedResourcesDict[filename] = photo
        return photo

    def saveOptionsToAppConfig(self):

        #
        # stores the current value of a combobox into appconfig
        #
        def storeComboBoxCurrentValueInControlGroupToAppConfig(controlsGroup):
            comboBox = controlsGroup.comboBox
            configDictKey = controlsGroup.configDictKey
            comboBoxCurrentValue = comboBox.get()
            g.appConfig.dict[configDictKey] = comboBoxCurrentValue
            return comboBoxCurrentValue


        #
        # stores all the values of a combobox into appconfig
        #
        def storeComboBoxAllValuesInControlGroupToAppConfig(controlsGroup):

            comboBox = controlsGroup.comboBox
            configDictKey = controlsGroup.configDictKey

            #
            # get list of current values from combobox, including the current
            # value the user entered which may not be in the combobox's list yet
            #
            comboBoxValuesList = list(comboBox['values'])       # combobox maintains list as tuples - convert to list for easier manipulation
            comboBoxCurrentValue = comboBox.get()
            if isStrValidWithNonWhitespaceChars(comboBoxCurrentValue):
                #
                # if value is already in list then we first remove
                # it. we insert the value as the first item of the list, to maintain an
                # MRU list of user selections
                #
                if comboBoxCurrentValue in comboBoxValuesList:
                    indexCurrentValueInList = comboBoxValuesList.index(comboBoxCurrentValue)
                    comboBoxValuesList.pop(indexCurrentValueInList)
                comboBoxValuesList.insert(0, comboBoxCurrentValue)
                del comboBoxValuesList[32:] # limit size of history to 32 elements, otherwise it can get unwieldly
                comboBox['values'] = comboBoxValuesList # combobox automatically converts list to tuples on assignment
            else:
                #
                # empty value. we don't allow empty values in the combobox list but we do
                # store them in the appconfig history so that we can recall them as the
                # default value when the appdict is loaded at session start
                #
                comboBoxValuesList.insert(0, "")  # empty value as first element of history

            #
            # store combobox values into appconfig
            #
            if comboBoxValuesList:
                # combo has a non-empty list of values. save them to appconfig dictionary
                g.appConfig.dict[configDictKey] = comboBoxValuesList
            elif configDictKey in g.appConfig.dict:
                # combo is empty and there are values for it in the appconfig dictionary. remove those values
                g.appConfig.dict.pop(configDictKey)

            return comboBoxCurrentValue

        #
        # get values from each UI control, saving to appconfig as we do
        #
        g.appConfig.dict['app_version'] = APP_VERSION
        values = {}
        values['xmlFileSource'] = storeComboBoxAllValuesInControlGroupToAppConfig(self.xmlFileControls)
        values['outputDir'] = storeComboBoxAllValuesInControlGroupToAppConfig(self.outputDirControls)
        values['videoCombine'] = storeComboBoxCurrentValueInControlGroupToAppConfig(self.videoCombineControls)
        values['fileExists'] = storeComboBoxCurrentValueInControlGroupToAppConfig(self.fileExistsControls)
        values['fileContainer'] = storeComboBoxAllValuesInControlGroupToAppConfig(self.fileContainerControls)
        values['outputNameSpec'] = storeComboBoxAllValuesInControlGroupToAppConfig(self.outputNameSpecControls)
        values['videoArgs'] = storeComboBoxAllValuesInControlGroupToAppConfig(self.videoArgsControls)
        values['audioArgs'] = storeComboBoxAllValuesInControlGroupToAppConfig(self.audioArgsControls)
        values['additionalFfmpegArgs'] = storeComboBoxAllValuesInControlGroupToAppConfig(self.additionalFfmpegArgsControls)
        values['addtionalPrecutArgs'] = storeComboBoxAllValuesInControlGroupToAppConfig(self.additionalPrecutArgsControls)
        values['loggingLevel'] = storeComboBoxCurrentValueInControlGroupToAppConfig(self.loggingControls)
        values['ffmpegLoggingLevel'] = storeComboBoxCurrentValueInControlGroupToAppConfig(self.ffmpegLoggingControls)
        values['ffmpegdir'] = storeComboBoxAllValuesInControlGroupToAppConfig(self.ffmpegDirControls)
        g.appConfig.saveAppConfig()
        return values

    def runButtonClick(self):

        #
        # get values from UI controls, saving them to appconfig as well
        #
        values = self.saveOptionsToAppConfig()

        #
        # validate file/paths, building arguments to precut as we do
        #
        if not values['xmlFileSource']:
            tkMessageBox.showerror("XML File Source", "No XML filename was specified")
            return
        if not os.path.exists(values['xmlFileSource']):
            tkMessageBox.showerror("XML File Source", "XML file \"{:s}\" does not exist".format(values['xmlFileSource']))
            return
        argStr = "\"{:s}\"".format(values['xmlFileSource'])

        if not values['outputDir']:
            tkMessageBox.showerror("Output Directory", "No Output Directory was specified")
            return
        if values['outputDir'] != Application.__DefaultOutputDir:
            if not os.path.exists(values['outputDir']):
                tkMessageBox.showerror("Output Directory", "Output Directory \"{:s}\" does not exist".format(values['outputDir']))
                return
            argStr += " --outputdir \"{:s}\"".format(values['outputDir'])

        argStr += " {:s}".format(CmdArgs.ClipCombineChoicesDict[values['videoCombine']])
        argStr += " {:s}".format(CmdArgs.FileExistsChoicesDict[values['fileExists']])

        if values['fileContainer'] and values['fileContainer'] != Application.__DefaultFileContainer:
            argStr += " --container {:s}".format(values['fileContainer'])

        if values['outputNameSpec'] and values['outputNameSpec'] != Application.__DefaultOutputNameSpec:
            argStr += " --outputnamespec \"{:s}\"".format(values['outputNameSpec'])

        #
        # for video/audio args there are a few cases:
        #
        #   __DefaultPrecutcmdVideoOrArgs - use default values built into precutcmd, which means
        #   not passing any --videoargs and/or --audioargs to precutcmd. I decided to leave
        #   this value out as a predefined/default option in the combobox selection because I found
        #   it too confusing for users when its included along with __DefaultFfmpegVideoOrAudioArgs
        #
        #   __DefaultFfmpegVideoOrAudioArgs - use default values built into ffmpeg, which means
        #   we pass a --videoargs "" and/or --audioargs "", which tells precutcmd to not pass its
        #   own internal default video and/or audio args to ffmpeg
        #
        #   Blank value - same as the __DefaultFfmpegVideoOrAudioArgs case. we pass any empty parameter string
        #
        #   Any other value - we pass as values to --videoargs and/or --audioargs
        #
        #   NOTE: To work around the issue of precutcmd's argparse interpreting the command-line dash passed
        #   to it for parameters that will then be passed to ffmpeg, we encase the arguments in qutoes and insert
        #   a space before the first user-entered argument so that argparse doesn't interpret the first dash
        #
        if values['videoArgs'] != Application.__DefaultPrecutcmdVideoOrArgs:
            videoArgs = values['videoArgs']
            if videoArgs == Application.__DefaultFfmpegVideoOrAudioArgs:
                videoArgs = ''
            argStr += " --videoargs \"{:s}\"".format(videoArgs)

        if values['audioArgs'] != Application.__DefaultPrecutcmdVideoOrArgs:
            audioArgs = values['audioArgs']
            if audioArgs == Application.__DefaultFfmpegVideoOrAudioArgs:
                audioArgs = ''
            argStr += " --audioargs \" {:s}\"".format(audioArgs)

        if values['additionalFfmpegArgs']:
            argStr += " --ffmpegargs \" {:s}\"".format(values['additionalFfmpegArgs'])

        if values['addtionalPrecutArgs']:
            argStr += "{:s}".format(values['addtionalPrecutArgs'])

        argStr += " {:s}".format(CmdArgs.LoggingLevelChoicesDict[values['loggingLevel']])
        argStr += " {:s}".format(CmdArgs.FfmpegLoggingChoicesDict[values['ffmpegLoggingLevel']])

        if values['ffmpegdir'] and values['ffmpegdir'] != Application.__DefaultFfmpegDir:
            if not os.path.exists(values['ffmpegdir']):
                tkMessageBox.showerror("ffmpeg Directory", "Directory \"{:s}\" specified for ffmpeg executable does not exist".format(values['ffmpegdir']))
                return
            argStr += " --ffmpegdir \"{:s}\"".format(values['ffmpegdir'])

        launchPrecut(argStr)
    def wmDeleteWindow(self):
        root.destroy()


#
# main app routine
#
def main():

    global root

    #
    # verify we're running under a tested version of python
    #
    verifyPythonVersion()

    #
    # establish our app environment, including our app-specific subdirectories
    #
    establishAppEnvironment()
    if not os.path.exists(g.appResourceDir):
        tkMessageBox.showwarning("Setup Error", "The appresource subdirectory is missing. This directory and its files (icons/bitmaps/etc..) are needed for proper functioning of precut :(")
        return

    #
    # initialize tkinter and run app
    #
    root = Tk()
    app = Application(master=root)
    g.app = app
    root.protocol("WM_DELETE_WINDOW", app.wmDeleteWindow)
    app.mainloop()

#
# program entry point
#
if __name__ == "__main__":
    main()
