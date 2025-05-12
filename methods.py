import win32com.client

class ComMethods:
    def __init__(self, ps):
        self.ps = ps

    #Session Management Functions

    def Login(self, username: str, password: str, server: str):
        try:
            self.ps.Login(username, password, server)
        except Exception as e:
            if username == None or password == None or server == None:
                raise RuntimeError(f"Failed to create login. This is most likely due to a passed parameter being \"None\".") from e
            else:
                raise RuntimeError(f"Failed to login using username {username}") from e
    
    def LoginEx(self, username: str, password: str):
        try:
            self.ps.LoginEx(username, password)
        except Exception as e:
            if username == None or password == None:
                raise RuntimeError(f"Failed to create login. This is most likely due to a passed parameter being \"None\".") from e
            else:
                raise RuntimeError(f"Failed to login using username {username}") from e

    def Logout(self):
        self.ps.Logout

    def LogoutEx(self):
        self.ps.LogoutEx

    def Start(self, started: bool):
        try:
            self.ps.Start(started)
        except Exception as e:
            if started == None:
                raise RuntimeError(f"Failed to start RadWhere. This is most likely due to a passed parameter being \"None\".") from e
            else:
                raise RuntimeError(f"Failed to start RadWhere") from e

    def Terminate(self):
        self.ps.Terminate

    def Stop(self):
        self.ps.Stop

    #Report Management

    def CreateNewReport(self, accession: str, bStartDictation: bool):
        try:
            self.ps.CreateNewReport(accession, bStartDictation)
        except Exception as e:
            if accession == None or bStartDictation == None:
                raise RuntimeError(f"Failed to create new report. This is most likely due to a passed parameter being \"None\".") from e
            else:
                raise RuntimeError(f"Failed to create new report for accession {accession}") from e
    
    def OpenReport(self, accessions: str, site=None):
        try:
            print(accessions, site)
            self.ps.OpenReport(None, accessions)
        except Exception as e:
            if accessions == None or site == None:
                raise RuntimeError(f"Failed to open report(s). This is most likely due to a passed parameter being \"None\".") from e
            else:
                raise RuntimeError(f"Failed to open report(s) for accession(s) {accessions}") from e
    
    def OpenReportMrn(self, accessions: str, site="", mrn=""):
        try:
            self.ps.OpenReport(site, accessions, mrn)
        except Exception as e:
            if accessions == None or site == None or mrn == None:
                raise RuntimeError(f"Failed to open report(s). This is most likely due to a passed parameter being \"None\".") from e
            else:
                raise RuntimeError(f"Failed to open report(s) for accession(s) {accessions} with MRN(s) {mrn}") from e

    def CloseReport(self, shouldSign: bool, markPrelim: bool):
        try:
            self.ps.CloseReport(shouldSign, markPrelim)
        except Exception as e:
            if shouldSign == None or markPrelim == None:
                raise RuntimeError(f"Failed to close report. This is most likely due to a passed parameter being \"None\".") from e
            else:
                raise RuntimeError(f"Failed to close report") from e
            
    def SaveReport(self, shouldClose: bool):
        try:
            self.ps.SaveReport(shouldClose)
        except Exception as e:
            if shouldClose == None:
                raise RuntimeError(f"Failed to save report. This is most likely due to a passed parameter being \"None\".") from e
            else:
                raise RuntimeError(f"Failed to save report") from e
    
    def InsertAutoText(self, autoTextName: str, strToReplace: str):
        try:
            self.ps.InsertAutoText(autoTextName, strToReplace)
        except Exception as e:
            if autoTextName == None or strToReplace == None:
                raise RuntimeError(f"Failed to insert autotext. This is most likely due to a passed parameter being \"None\".") from e
            else:
                raise RuntimeError(f"Failed to insert autotext") from e

    #Order Management
    def PreviewOrder(self, accesionNumbers: str, site=""):
        try:
            self.ps.PreviewOrders(accesionNumbers)
        except Exception as e:
            if accesionNumbers == None:
                raise RuntimeError(f"Failed to preview order(s). This is most likely due to a passed parameter being \"None\".") from e
            else:
                raise RuntimeError(f"Failed to preview order(s)") from e
    def AssociateOrders(self, accessionNumbers: str):
        try:
            self.ps.AssociateOrders(accessionNumbers)
        except Exception as e:
            if accessionNumbers == None:
                raise RuntimeError(f"Failed to associate orders. This is most likely due to a passed parameter being \"None\".") from e
            else:
                raise RuntimeError(f"Failed to associate orders") from e
            
    def AssociateOrdersWCurrent(self, currentAccession: str, newAccessions: str, site=""):
        try:
            self.ps.AssociateOrdersEx(site, currentAccession, newAccessions)
        except Exception as e:
            if currentAccession == None or newAccessions == None or site == None:
                raise RuntimeError(f"Failed to associate orders This is most likely due to a passed parameter being \"None\".") from e
            else:
                raise RuntimeError(f"Failed to associate orders") from e
    
    def DissociateOrders(self, accessionNumbers: str):
        try:
            self.ps.DissociateOrders(accessionNumbers)
        except Exception as e:
            if accessionNumbers == None:
                    raise RuntimeError(f"Failed to dissociate orders This is most likely due to a passed parameter being \"None\".") from e
            else:
                raise RuntimeError(f"Failed to dissociate orders") from e
        
    #Getter Methods
    def GetActiveAccessions(self) -> str:
        try:
            accessions = self.ps.AccessionNumbers
            return accessions
        except Exception as e:
            raise RuntimeError("Error getting AccessionNumbers") from e

    def GetAlwaysOnTop(self) -> str:
        try:
            always_on_top = self.ps.AlwaysOnTop
            return always_on_top
        except Exception as e:
            raise RuntimeError("Error getting AlwaysOnTop") from e

    def GetMinimized(self) -> str:
        try:
            minimized = self.ps.Minimized
            return minimized
        except Exception as e:
            raise RuntimeError("Error getting Minimized") from e

    def GetRestrictedSession(self) -> str:
        try:
            restricted_session = self.ps.RestrictedSession
            return restricted_session
        except Exception as e:
            raise RuntimeError("Error getting RestrictedSession") from e

    def GetRestrictedWorkflow(self) -> str:
        try:
            restricted_workflow = self.ps.RestrictedWorkflow
            return restricted_workflow
        except Exception as e:
            raise RuntimeError("Error getting RestrictedWorkflow") from e

    def GetSiteName(self) -> str:
        try:
            site_name = self.ps.SiteName
            return site_name
        except Exception as e:
            raise RuntimeError("Error getting SiteName") from e

    def GetUsername(self) -> str:
        try:
            username = self.ps.Username
            return username
        except Exception as e:
            raise RuntimeError("Error getting Username") from e

    def GetLoggedIn(self) -> str:
        try:
            logged_in = self.ps.LoggedIn
            return logged_in
        except Exception as e:
            raise RuntimeError("Error getting LoggedIn") from e
    
    def GetSite(self) -> str:
        try:
            site = self.ps.Site
            return site
        except Exception as e:
            raise RuntimeError("Error getting Site") from e

    def GetVisible(self) -> bool:
        try:
            visible = self.ps.Visible
            return visible
        except Exception as e:
            raise RuntimeError("Error getting Visible") from e

    #Setter Methods

    def SetAlwaysOnTop(self, value: bool) -> None:
        try:
            self.ps.AlwaysOnTop = value
        except Exception as e:
            raise RuntimeError("Error setting AlwaysOnTop") from e

    def SetMinimized(self, value: bool) -> None:
        try:
            self.ps.Minimized = value
        except Exception as e:
            raise RuntimeError("Error setting Minimized") from e

    def SetRestrictedSession(self, value: bool) -> None:
        try:
            self.ps.RestrictedSession = value
        except Exception as e:
            raise RuntimeError("Error setting RestrictedSession") from e

    def SetRestrictedWorkflow(self, value: bool) -> None:
        try:
            self.ps.RestrictedWorkflow = value
        except Exception as e:
            raise RuntimeError("Error setting RestrictedWorkflow") from e
    
    def SetSite(self, value: str) -> None:
        try:
            self.ps.Site = value
        except Exception as e:
            raise RuntimeError("Error setting Site") from e
        
    def SetVisible(self, value: bool) -> None:
        try:
            self.ps.Visible = value
        except Exception as e:
            raise RuntimeError("Error setting Visible") from e