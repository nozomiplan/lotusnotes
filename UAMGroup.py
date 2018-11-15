from win32com.client import Dispatch
import smtplib
import pywintypes
import sys # Try..Catch..Finally
import datetime

class ManageGroupAccount:

    user_dbName ='names.nsf'
    db_viewname = '($Users)'
 
    
    def __init__(self,accountInfo):
  
        self.accountInfo = accountInfo
        self.memberList=[]   
    

    def DoCreate(self):
        try:
            
            groupname = self.accountInfo.groupname
            notes = Dispatch('Lotus.NotesSession')
            notes.Initialize('P@ssIDAM')
            notesdb = notes.GetDatabase(self.accountInfo.getMailServer(), self.user_dbName,1)
       
            viewGroups = notesdb.GetView(self.db_viewname)
        
            newdoc = viewGroups.GetDocumentByKey(groupname)

            adminP = notes.CreateAdministrationProcess(self.accountInfo.getMailServer())

            if newdoc != None:
                return "DupAccount"
                notes = None 
            else:
                noteID = adminP.AddGroupMembers(groupname,  usersArray)
                if noteID != "":
                    return "Added"  
                else:
                    return "Cannot be added because noteID is not empty" 
                notes = None 
        except:
            print ("Unexcepted Error is found during create new group. Description:", sys.exc_info()[0])
            notes = None  
            raise

    def DoModifyGroup(self,NewGroupName):
        try:    
            oldGroupname = self.accountInfo.groupname
            notes = Dispatch('Lotus.NotesSession')
            notes.Initialize('P@ssIDAM')
            notesdb = notes.GetDatabase(self.accountInfo.getMailServer(), self.user_dbName,1)
            viewGroups = notesdb.GetView(self.db_viewname)

            doc = viewGroups.GetDocumentByKey(oldGroupname)

            if doc != None:
                listName = doc.GetFirstItem("ListName")
                if listName != None:
                    doc.RemoveItem("ListName")
                    doc.AppendItemValue("ListName",NewGroupName)
                    result = doc.Save(True,True)
                
            return result
        except:
            print ("Unexcepted Error is found during create new group. Description:", sys.exc_info()[0])
            notes = None  
            raise             


    def DoRemoveGroup(self):
        try:
            
            groupname = self.accountInfo.groupname
            notes = Dispatch('Lotus.NotesSession')
            notes.Initialize('P@ssIDAM')
            notesdb = notes.GetDatabase(self.accountInfo.getMailServer(), self.user_dbName,1)
       
            viewGroups = notesdb.GetView(self.db_viewname)
        
            deletedoc = viewGroups.GetDocumentByKey(groupname)

            adminP = notes.CreateAdministrationProcess(self.accountInfo.getMailServer())

            if deletedoc != None:
                noteID = adminP.DeleteGroup(groupname , True) 
                if noteID != "":
                    return "Removed"
                else:
                    return "Cannot be removed because note ID is empty"

            notes = None
        except:
            print ("Unexcepted Error is found during create new group. Description:", sys.exc_info()[0])
            notes = None  
            raise
    def DoAssignUsers(self,UserArray):
        try:    
            groupname = self.accountInfo.groupname
            notes = Dispatch('Lotus.NotesSession')
            notes.Initialize('P@ssIDAM')
            notesdb = notes.GetDatabase(self.accountInfo.getMailServer(), self.user_dbName,1)
            viewGroups = notesdb.GetView(self.db_viewname)
        
            doc = viewGroups.GetDocumentByKey(groupname)

            memberList = []

            if doc != None:
                memberRaw = doc.GetFirstItem("Members")
                if memberRaw != None:
                    memberArray = memberRaw.values
                    if memberArray != None:
                        for memberItem in memberArray:
                            memberList.append(memberItem)
                    
                    if UserArray != None:
                        for userItem in UserArray:
                            memberList.append(userItem)

                    
                    doc.RemoveItem("Members")
                    doc.AppendItemValue("Members",memberList)
                    result = doc.Save(True,True)
                    if result == True:
                        return "Assigned"
                    else:
                        return "Cannot be assigned"
                else: 
                    if UserArray != None:
                        for userItem in UserArray:
                            memberList.append(userItem)

                    doc.RemoveItem("Members")
                    doc.AppendItemValue("Members",memberList)
                    result = doc.Save(True,True)

                    if result == True:
                        return "Assigned"
                    else:
                        return "Cannot be assigned"


        except:
            print ("Unexcepted Error is found during create new group. Description:", sys.exc_info()[0])
            notes = None  
            raise       
        
        #return "DoModify"
    def DoRemoveUsers(self,UserList):
        try:    
            groupname = self.accountInfo.groupname
            notes = Dispatch('Lotus.NotesSession')
            notes.Initialize('P@ssIDAM')
            notesdb = notes.GetDatabase(self.accountInfo.getMailServer(), self.user_dbName,1)
            viewGroups = notesdb.GetView(self.db_viewname)
        
            doc = viewGroups.GetDocumentByKey(groupname)
            if doc != None:
                memberRaw = doc.GetFirstItem("Members")
                if memberRaw != None:
                    memberList = memberRaw.values


                    num_beforeremove = len(memberList)
                    memberList = [item for item in memberList if item not in UserList]
                    num_afteremove = len(memberList)


                  
                    if (num_beforeremove==num_afteremove):
                        return "Item not selected correctly"
                    else:
                        doc.RemoveItem("Members")
                        doc.AppendItemValue("Members",memberList)
                        result = doc.Save(True,True)
                        
                        if result == True:
                            return "Removed"
                        else:
                            return "Cannot be removed"


        except:
            print ("Unexcepted Error is found during remove users. Description:", sys.exc_info()[0])
            notes = None  
            raise       
    def DoListMembers(self,GroupName):       # Users parameters are ARRAY
        try:    
            groupname = self.accountInfo.groupname
            notes = Dispatch('Lotus.NotesSession')
            notes.Initialize('P@ssIDAM')
            notesdb = notes.GetDatabase(self.accountInfo.getMailServer(), self.user_dbName,1)
            viewGroups = notesdb.GetView(self.db_viewname)
        
            doc = viewGroups.GetDocumentByKey(groupname)

            if doc != None:
                memberRaw = doc.GetFirstItem("Members")
                memberRaw.IsSummary=True

                if memberRaw != None:
                    memberArray = memberRaw.values
                    return memberArray
                else:
                    return None 
        except:
            print ("Unexcepted Error is found during create new group. Description:", sys.exc_info()[0])
            notes = None  
            raise       