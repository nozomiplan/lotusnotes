from nameko.rpc import rpc
from nameko.events import EventDispatcher, event_handler
import UAMUsers
import UAMGroup
import NotesSetup

from pathlib import Path
import random


import json
from nameko.web.handlers import http
from werkzeug.wrappers import Response

class NotesService:
    name='notes_service'
    db_name="certlog.nsf"
    notes_rootfolder="C:\\IBM\\Lotus\\Domino\\data"
    acl="CN=IDAM ADMIN/O=POC"
    mailserver="IDAMDEVIUSR02.poc.et"
    accesspwd= 'P@ssIDAM'
    reglogfile ='certlog.nsf'
    db_folder ="C:\\source codes\\nameko_env\\projects\\notes_and_nameko\\userid"

    dispatch = EventDispatcher()
    


    @rpc
    def createMail(self,surname,middlename,firstname,passwd,jobtitle,internetaddress,home_tel,office_tel,formName,rank,grade,location,section):

        certlogloc= self.notes_rootfolder + "\\cert.id"

        name = firstname
        surname = surname
        middlename = middlename
        fullname = name + " " + middlename + " " +  surname
        
        needinternetaccount = False
        internetemailaddress = internetaddress

        accountname= 'CN='+ fullname + '/O=POC'
        home_phone=home_tel
        office_phone=office_tel
        newUserPwd = passwd
        job_title = jobtitle

        dest_nsf='mail\\'+ fullname.replace(" ", "")  +'.nsf'
        dest_id = self.db_folder + "\\" + fullname.replace(" ", "")+".id"

        randout = str(hash(random.randint(1,10001)))

        nsf_file = Path(self.notes_rootfolder+dest_nsf)
        if nsf_file.is_file():
            dest_nsf='mail\\'+ fullname  + '_'+randout + '.nsf'


        notesInfo = NotesSetup.NotesSetup()
        notesInfo.setuseraccountname(name + " " +  middlename + " " + surname)
        notesInfo.setCertlogloc(certlogloc)
        notesInfo.setaccesspwd(self.accesspwd)
        notesInfo.setACL(self.acl)
        notesInfo.setdbfolder(self.db_folder)
        notesInfo.setDBName(self.db_name)
        notesInfo.setMailServer(self.mailserver)
        notesInfo.setinternetemailaddress(internetemailaddress)
        notesInfo.setneedinternetaccount(needinternetaccount)
        notesInfo.setreglogfile(self.reglogfile)
        notesInfo.setuseraccountgivenname(name)
        notesInfo.setuseraccountmiddlename(middlename)
        notesInfo.setuseraccountsurname(surname)
        notesInfo.setuseraccounthomenumber(home_phone)
        notesInfo.setuseraccountofficenumber(office_phone)
        notesInfo.setuseraccountjobtitle(job_title)
        notesInfo.setuseridfile(dest_id)
        notesInfo.setusernsf(dest_nsf)

        notesInfo.setuserpwd(newUserPwd)
        notesInfo.setuserfullaccountname(accountname)

        noteInfo.setform(formName)
        noteInfo.setgrade(grade)
        noteInfo.setrank(rank)
        noteInfo.setlocation(location)
        noteInfo.setsection(section)


        sUAMUser = UAMUsers.ManageUserAccount(notesInfo) #init

       
        return "User Account is found duplicated, cannot be created" if sUAMUser.DoCreate()=='DupAccount' else "User account is created" if sUAMUser.DoModify()=='Modified' else "Unexcepted Error found during create account"

    @rpc
    def modifyMail(self,surname,middlename,firstname,jobtitle,internetaddress,home_tel,office_tel,formName,rank,grade,location,section):

        name = firstname
        surname = surname
        middlename = middlename
        fullname = name + " " + middlename + " " +  surname
        accountname= 'CN='+ fullname + '/O=POC'
        notesInfo = NotesSetup.NotesSetup()
        notesInfo.setuserfullaccountname(accountname)     
        notesInfo.setuseraccountgivenname(name)
        notesInfo.setuseraccountmiddlename(middlename)
        notesInfo.setuseraccountsurname(surname)  
        notesInfo.setuseraccountjobtitle(jobtitle)
        notesInfo.setinternetemailaddress(internetaddress)
        notesInfo.setuseraccounthomenumber(home_tel)
        notesInfo.setuseraccountofficenumber(office_tel)
        notesInfo.setMailServer(self.mailserver)
        noteInfo.setform(formName)
        noteInfo.setgrade(grade)
        noteInfo.setrank(rank)
        noteInfo.setlocation(location)
        noteInfo.setsection(section)
        sUAMUser = UAMUsers.ManageUserAccount(notesInfo) #init
        return "User account is modified" if sUAMUser.DoModify()=="Modified" else "Unexcepted Error found during modify account"

    @rpc
    def removeMail(self,surname,middlename,firstname):
        name = firstname
        surname = surname
        middlename = middlename
        fullname = name + " " + middlename + " " +  surname
        accountname= 'CN='+ fullname + '/O=POC'
        notesInfo = NotesSetup.NotesSetup()
        notesInfo.setuserfullaccountname(accountname)     
        notesInfo.setuseraccountgivenname(name)
        notesInfo.setuseraccountmiddlename(middlename)
        notesInfo.setuseraccountsurname(surname)  
        notesInfo.setMailServer(self.mailserver)

        sUAMUser = UAMUsers.ManageUserAccount(notesInfo) #init
        return "User account is removed" if sUAMUser.DoRemove()=="Removed" else "Unexcepted Error during remove account"

    @rpc
    def createGroup(self,groupName):
        notesInfo = NotesSetup.GroupSetup()
        notesInfo.setgroupname(groupName)
        notesInfo.setMailServer(self.mailserver)

        sUAMGroup = UAMGroup.ManageGroupAccount(notesInfo)
        return sUAMGroup.DoCreate()
        #return "Group is found duplicated, cannot be created" if sUAMGroup.DoCreate()=='DupAccount' else "Group is created" if sUAMGroup.DoModify()=='Modified' else "Unexcepted Error found during create group"
        #return "New Group has been created"
    @rpc
    def modifyGroupName(self,groupName,newGroupName):
        notesInfo = NotesSetup.GroupSetup()
        notesInfo.setgroupname(groupName)
        notesInfo.setMailServer(self.mailserver)
        sUAMGroup = UAMGroup.ManageGroupAccount(notesInfo)
        return sUAMGroup.DoModifyGroup(newGroupName)
        
    @rpc
    def removeGroup(self,groupName):  
        notesInfo = NotesSetup.GroupSetup()
        notesInfo.setgroupname(groupName)
        notesInfo.setMailServer(self.mailserver)
        sUAMGroup = UAMGroup.ManageGroupAccount(notesInfo)
        return sUAMGroup.DoRemoveGroup()  
        
    @rpc
    def assignUserToGroup(self,userArray,groupName):   
        notesInfo = NotesSetup.GroupSetup()
        notesInfo.setgroupname(groupName)
        notesInfo.setMailServer(self.mailserver)
        sUAMGroup = UAMGroup.ManageGroupAccount(notesInfo)
        return sUAMGroup.DoAssignUsers(userArray)
    @rpc
    def removeUserFromGroup(self,userArray,groupName):   
        notesInfo = NotesSetup.GroupSetup()
        notesInfo.setgroupname(groupName)
        notesInfo.setMailServer(self.mailserver)
        sUAMGroup = UAMGroup.ManageGroupAccount(notesInfo)
        return sUAMGroup.DoRemoveUsers(userArray)
        
        

    @rpc
    def listGroupUsers(self,groupName):
        notesInfo = NotesSetup.GroupSetup()
        notesInfo.setgroupname(groupName)
        notesInfo.setMailServer(self.mailserver)
        sUAMGroup = UAMGroup.ManageGroupAccount(notesInfo)
        return sUAMGroup.DoListMembers(groupName)

    
    