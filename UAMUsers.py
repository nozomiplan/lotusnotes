from win32com.client import Dispatch
import smtplib
import pywintypes
import sys # Try..Catch..Finally
import datetime



class ManageUserAccount: #CreateUserAccount:
    
    user_dbName ='names.nsf'
    db_viewname = '($Users)'


    def __init__(self,accountInfo):
        self.accountInfo = accountInfo
    def DoCreate(self):
        try:
            notes = Dispatch('Lotus.NotesSession')
            notes.Initialize('P@ssIDAM')
            notesdb = notes.GetDatabase(self.accountInfo.getMailServer(), self.user_dbName,1)
            viewUsers = notesdb.GetView(self.db_viewname)
            currentDoc = viewUsers.GetDocumentByKey(self.accountInfo.getuserfullaccountname())

            if currentDoc != None: 
                return "DupAccount"
            else:
                reg = notes.CreateRegistration()
                reg.CreateMailDb = True
                reg.CertifierIDFile = self.accountInfo.getCertlogloc()
                reg.RegistrationServer = self.accountInfo.getMailServer()
                reg.IDType = 172
                reg.RegistrationLog = self.accountInfo.getreglogfile()
                reg.Expiration=datetime.datetime.now() + datetime.timedelta(days=1095) #Starting from date of creating this data for 1095 days
                reg.UpdateAddressBook = True
                reg.StoreIDInAddressBook = True
                reg.MailACLManager=self.accountInfo.getACL()
                reg.MailOwnerAccess=2



                reg.RegisterNewUser(self.accountInfo.getuseraccountsurname()#name
                            ,self.accountInfo.getuseridfile()#dest_id
                            ,self.accountInfo.getMailServer()#mailserver
                            ,self.accountInfo.getuseraccountgivenname() 
                            ,self.accountInfo.getuseraccountmiddlename()#middlename
                            ,self.accountInfo.getaccesspwd()#accesspwd
                            ,""
                            ,""
                            ,self.accountInfo.getusernsf()#dest_nsf
                            ,""
                            ,self.accountInfo.getuserpwd() #newUserPwd
                            ,176)
               
                return "Added"  
            reg = None
            notes = None  
        except:
            print ("Unexcepted Error is found during create account. Description:", sys.exc_info()[0])
            reg = None
            notes = None  
       
            raise
    
    def DoModify(self):
        try:
            notes = Dispatch('Lotus.NotesSession')
            notes.Initialize('P@ssIDAM')
            notesdb = notes.GetDatabase(self.accountInfo.getMailServer(), self.user_dbName,1)
            viewUsers = notesdb.GetView(self.db_viewname)
            currentDoc = viewUsers.GetDocumentByKey(self.accountInfo.getuserfullaccountname())

            if currentDoc != None:

                currentDoc.ReplaceItemValue("FirstName",self.accountInfo.getuseraccountgivenname()) 
                currentDoc.ReplaceItemValue("MiddleInitial",self.accountInfo.getuseraccountmiddlename()) 
                currentDoc.ReplaceItemValue("LastName",self.accountInfo.getuseraccountsurname()) 
                currentDoc.ReplaceItemValue("JobTitle",self.accountInfo.getuseraccountjobtitle()) 
                currentDoc.ReplaceItemValue("PhoneNumber",self.accountInfo.getuseraccounthomenumber())
                currentDoc.ReplaceItemValue("OfficePhoneNumber",self.accountInfo.getuseraccountofficenumber())
                currentDoc.ReplaceItemValue("Form",self.accountInfo.getform()) 
                #getrank (Contract System Analyst)
                currentDoc.ReplaceItemValue("Rank",self.accountInfo.getrank()) 
                #getgrade (Contract System Analyst)
                currentDoc.ReplaceItemValue("Grade",self.accountInfo.getgrade()) 
                #getlocation (Queensway Government OFfice)
                currentDoc.ReplaceItemValue("Location",self.accountInfo.getlocation()) 
                #getsection (ITOT)
                currentDoc.ReplaceItemValue("Section",self.accountInfo.getsection()) 
                
                
                result = currentDoc.save(True,True)
               
                return  "Modified" if result==True else "ModifyError"
               

            else:
                
                
                return "AccountNotFound"
            currentDoc = None
            notes = None 
        except:
            print ("Unexcepted Error is found during modify account. Description:", sys.exc_info()[0])
            return "Failed"
            raise
    def DoRemove(self):
        try:
            notes = Dispatch('Lotus.NotesSession')
            notes.Initialize('P@ssIDAM')
            notesdb = notes.GetDatabase(self.accountInfo.getMailServer(), self.user_dbName,1)
            viewUsers = notesdb.GetView(self.db_viewname)
            currentDoc = viewUsers.GetDocumentByKey(self.accountInfo.getuserfullaccountname())
            #print (currentDoc)
            if currentDoc != None:
                currentDoc.ReplaceItemValue("FullName",'Disabled_' + self.accountInfo.getuserfullaccountname()) 
                currentDoc.ReplaceItemValue("FirstName",'Disabled_' + self.accountInfo.getuseraccountgivenname()) 
                currentDoc.ReplaceItemValue("MiddleInitial",'Disabled_' + self.accountInfo.getuseraccountmiddlename()) 
                currentDoc.ReplaceItemValue("LastName",'Disabled_' + self.accountInfo.getuseraccountsurname()) 
                               
                result = currentDoc.save(True,True)
               
                return  "Removed" if result==True else "RemoveError"
               

            else:
                
                
                return "AccountNotFound"
            currentDoc = None
            notes = None
        except:
            print ("Unexcepted Error is found during modify account. Description:", sys.exc_info()[0])
            return "Failed"
            raise    
    def DoAssignToGroup(self):
        try:
            notes = Dispatch('Lotus.NotesSession')
            notes.Initialize('P@ssIDAM')
            notesdb = notes.GetDatabase(self.accountInfo.getMailServer(), self.user_dbName,1)
            viewUsers = notesdb.GetView(self.db_viewname)
            currentDoc = viewUsers.GetDocumentByKey(self.accountInfo.getuseraccountname())
            if currentDoc != None:
                return "DoAssign_1"
            else:
                return "AccountNotFound"
        except:
            print ("Unexcepted Error is found during assign account. Description:", sys.exc_info()[0])
            return "Failed"
            raise
 

