# Prototype Python code for GUI for tool to generate workscope exhibits
#
# NOTES: 
#   1. This script was written to run under Python 2.7.x
#   2. This script relies upon the wxPython librarybeing installed
#   3. To install wxPython under Python 2.7.x:
#       <Python_installation_folder>/python.exe -m pip install wxPython
#
# Author: Benjamin Krepp
# Date: 31 August 2018
#

import wx, wx.html
from workscope_exhibit_tool import main

# Code for the application's GUI begins here.
#
aboutText = """<p>Help text for this program is TBD.<br>
This program is running on version %(wxpy)s of <b>wxPython</b> and %(python)s of <b>Python</b>.
See <a href="http://wiki.wxpython.org">wxPython Wiki</a></p>""" 

class HtmlWindow(wx.html.HtmlWindow):
    def __init__(self, parent, id, size=(600,400)):
        wx.html.HtmlWindow.__init__(self,parent, id, size=size)
        if "gtk2" in wx.PlatformInfo:
            self.SetStandardFonts()
    # end_def __init__()

    def OnLinkClicked(self, link):
        wx.LaunchDefaultBrowser(link.GetHref())
    # end_def OnLinkClicked()
# end_class HtmlWindow

class AboutBox(wx.Dialog):
    def __init__(self):
        wx.Dialog.__init__(self, None, -1, "About the Workscope Exhibit Tool",
                           style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER|wx.TAB_TRAVERSAL)
        hwin = HtmlWindow(self, -1, size=(400,200))
        vers = {}
        vers["python"] = sys.version.split()[0]
        vers["wxpy"] = wx.VERSION_STRING
        hwin.SetPage(aboutText % vers)
        btn = hwin.FindWindowById(wx.ID_OK)
        irep = hwin.GetInternalRepresentation()
        hwin.SetSize((irep.GetWidth()+25, irep.GetHeight()+10))
        self.SetClientSize(hwin.GetSize())
        self.CentreOnParent(wx.BOTH)
        self.SetFocus()
    # end_def __init__()
# end_class AboutBox

# This is the class for the main GUI itself.
class Frame(wx.Frame):
    xlsxFileName = ''
    def __init__(self, title):
        wx.Frame.__init__(self, None, title=title, pos=(150,150), size=(600,250),
                          style=wx.SYSTEM_MENU | wx.CAPTION | wx.CLOSE_BOX)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        menuBar = wx.MenuBar()
        menu = wx.Menu()
        m_exit = menu.Append(wx.ID_EXIT, "E&xit\tAlt-X", "Close window and exit program.")
        self.Bind(wx.EVT_MENU, self.OnClose, m_exit)
        menuBar.Append(menu, "&File")
        menu = wx.Menu()
        m_about = menu.Append(wx.ID_ABOUT, "&About", "Information about this program")
        self.Bind(wx.EVT_MENU, self.OnAbout, m_about)
        menuBar.Append(menu, "&Help")
        self.SetMenuBar(menuBar)
        
        self.statusbar = self.CreateStatusBar()

        panel = wx.Panel(self)
        box = wx.BoxSizer(wx.VERTICAL)
        box.AddSpacer(20)
              
        m_select_file = wx.Button(panel, wx.ID_ANY, "Select Excel workbook")
        m_select_file.Bind(wx.EVT_BUTTON, self.OnSelectFile)
        box.Add(m_select_file, 0, wx.CENTER)
        box.AddSpacer(20)
        
        m_generate = wx.Button(panel, wx.ID_ANY, "Generate HTML for Exhibits")
        m_generate.Bind(wx.EVT_BUTTON, self.OnGenerate)
        box.Add(m_generate, 0, wx.CENTER)
 
        # Placeholder for name of selected .xlsx file; it is populated in OnSelectFile(). 
        self.m_text = wx.StaticText(panel, -1, " ")
        self.m_text.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.NORMAL))
        self.m_text.SetSize(self.m_text.GetBestSize())
        box.Add(self.m_text, 0, wx.ALL, 10)      
        
        panel.SetSizer(box)
        panel.Layout()
    # end_def __init__()
        
    def OnClose(self, event):
        dlg = wx.MessageDialog(self, 
            "Do you really want to close this application?",
            "Confirm Exit", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_OK:
            self.Destroy()
    # end_def OnClose()

    def OnSelectFile(self, event):
        frame = wx.Frame(None, -1, 'win.py')
        frame.SetSize(0,0,200,50)
        openFileDialog = wx.FileDialog(frame, "Select workscope exhibit spreadsheet", "", "", 
                                       "Excel files (*.xlsx)|*.xlsx", 
                                       wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        openFileDialog.ShowModal()
        self.xlsxFileName = openFileDialog.GetPath()
        self.m_text.SetLabel("Selected .xlsx file: " + self.xlsxFileName)
        openFileDialog.Destroy()
    # end_def OnSelectFile()
    
    def OnGenerate(self, event):
        dlg = wx.MessageDialog(self, 
            "Do you really want to run the HTML generation tool?",
            "Confirm: OK/Cancel", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_OK:
            main(self.xlsxFileName)
            message = "HTML for workscope exhibits generated."
            caption = "Work Scope Exhibit Tool"
            dlg = wx.MessageDialog(None, message, caption, wx.OK | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()
            self.Destroy()
        else:
            self.Destroy()
    # end_def OnGenerate()

    def OnAbout(self, event):
        dlg = AboutBox()
        dlg.ShowModal()
        dlg.Destroy() 
    # end_def OnAbout()
# end_class Frame

# The code for the GUI'd application itself begins here.
#
app = wx.App(redirect=True)   # Error messages go to popup window
top = Frame("Workscope Exhibit Tool")
top.Show()
app.MainLoop()
