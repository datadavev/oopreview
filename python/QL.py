'''
For OSX, make this accessible from:
  ~/Library/Application Support/OpenOffice/4/user/Scripts/python

Add that location to your trusted scripts in OO preferences, then for
convenience, assign keys like ⌘⇧w = doPreview, ⌘⇧e = doOpen, and
⌘⇧d = doFinder. Then with the cursor over a cell with a filename, press the
key combo to view / edit / locate the listed file.
'''
import uno
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_OK, DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY, DEFAULT_BUTTON_YES, DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX

import os
from subprocess import Popen, call

DEFAULT_CONFIG_FILE=os.path.join(os.path.expanduser('~'),'dv_oo_preview.ini')
DEFAULT_BASE=os.path.expanduser('~/Documents/Dropbox/')
EXTENSIONS = ['.pdf','.jpg','.txt','.md','.png','.csv']


def getConfigSettings(config_file=DEFAULT_CONFIG_FILE):
  pass


# Show a message box with the UNO based toolkit
def MessageBox(ParentWin,
               MsgText,
               MsgTitle, MsgType=MESSAGEBOX, MsgButtons=BUTTONS_OK):
  ctx = uno.getComponentContext()
  sm = ctx.ServiceManager
  sv = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
  myBox = sv.createMessageBox(ParentWin, MsgType, MsgButtons, MsgTitle, MsgText)
  return myBox.execute()


def doMessageBox(msg, title="Message"):
  doc = XSCRIPTCONTEXT.getDocument()
  parentwin = doc.CurrentController.Frame.ContainerWindow
  return MessageBox(parentwin, msg, title)


def _getFilePath(fname=None, fbase=None, check_exists=True):
  if fname is None:
    doMessageBox("No file selected.\nCheck cursor position?",title="ERROR")
    return None
  fname = fname.strip()
  if len(fname) < 1:
    doMessageBox("No file selected.\nCheck cursor position?",title="ERROR")
    return None
  if fbase is None:
    fbase = ''
  fbase = fbase.strip()
  fpath = os.path.join(DEFAULT_BASE, fbase, fname)
  if not check_exists:
    return fpath
  if os.path.exists(fpath):
    return fpath
  for ext in EXTENSIONS:
    ftest = fpath + ext
    if os.path.exists(ftest):
      return ftest
  doMessageBox("File not found: {0}\nCheck cursor position?".format(fpath),title="ERROR")
  return None


def _getSelectedColumnRow():
  oDoc = XSCRIPTCONTEXT.getDocument()
  oSheet = oDoc.CurrentController.ActiveSheet
  oSelection = oDoc.getCurrentSelection()
  oArea = oSelection.getRangeAddress()
  return oSheet, oArea.StartColumn, oArea.StartRow


def _getFilePathFromCursor():
  oSheet, fcol, frow = _getSelectedColumnRow()
  fbase = oSheet.getCellByPosition(fcol, 0).String
  fname = oSheet.getCellByPosition(fcol, frow).String
  return _getFilePath(fname=fname, fbase=fbase)


def _getFileSourceDestFromCursor():
  oSheet, fcol, frow = _getSelectedColumnRow()
  fbase = oSheet.getCellByPosition(fcol, 0).String
  fname = oSheet.getCellByPosition(fcol, frow).String
  fdest = oSheet.getCellByPosition(fcol+1, frow).String
  fsource = _getFilePath(fname=fname, fbase=fbase)
  filename, extension = os.path.splitext(fsource)
  fdest = os.path.join(DEFAULT_BASE,fbase,fdest + extension)
  return [fsource, fdest]


def CopyAndRename():
  oSheet, fsource, fdest = _getFileSourceDestFromCursor()
  if fsource is None:
    doMessageBox("File not found, cursor position?",title="ERROR")
    return
  if os.path.exists(fdest):
    doMessageBox("Destination "+fdest+" exists!",title="ERROR")
    return
  call(["cp", "-p", fsource, fdest])


def CopyToOriginals():
  oSheet, fcol, frow = _getSelectedColumnRow()
  fbase = oSheet.getCellByPosition(fcol, 0).String
  fname = oSheet.getCellByPosition(fcol, frow).String
  fsource = _getFilePath(fname=fname, fbase=fbase)
  if fsource is None:
    doMessageBox("File not found, cursor position?",title="ERROR")
    return
  filename, extension = os.path.splitext(fsource)
  filename0, extension0 = os.path.splitext(fname)
  if extension0 == extension:
    fdest = os.path.join(DEFAULT_BASE,fbase,'ORIGINALS',fname)
  else:
    fdest = os.path.join(DEFAULT_BASE,fbase,'ORIGINALS',fname + extension)
  if os.path.exists(fdest):
    doMessageBox("Destination "+fdest+" exists!",title="ERROR")
    return
  call(["cp", "-p", fsource, fdest])
  call(["chmod","a-w",fdest])
  if not os.path.exists(fdest):
    doMessageBox("File not copied!\n{0}".format(fdest),title="ERROR")
    return
  doMessageBox("File copied OK.",title="Notice")


def doPreview( ):
  fpath = _getFilePathFromCursor()
  if fpath is None:
    return
  Popen(["qlmanage","-p",fpath])
  return None


def doOpen():
  fpath = _getFilePathFromCursor()
  if fpath is None:
    return
  Popen(["open",fpath])


def doFinder():
  fpath = _getFilePathFromCursor()
  if fpath is None:
    return
  Popen(["open","-R",fpath])


g_exportedScripts = (doPreview, doOpen, doFinder)
