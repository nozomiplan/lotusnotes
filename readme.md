

The following items should be ready

Python 3.6 (32bit)
pypiwin32
Lotus Notes DLL 
VirtualEnv
nameko
RabbitMQ (Message Queue)

Entry Point
temp_messenger\service.py


Host Call 
1. nameko run temp_messenger.service --config config.yaml

Client Call

1.  command "nameko shell"
2.  Create Mail Account
        n.rpc.notes_service.createMail("Leung","ABC33","John","P@ssword","Tester33","timhk.jhnleung@poc.et","12345678","12345555")
    Remove Mail Account    
        n.rpc.notes_service.removeMail("Leung","ABC","John")
    Create New Group
        n.rpc.notes_service.createGroup("Python Group")
    Assign Group Member into Group
        n.rpc.notes_service.assignUserToGroup(['CN=Leung CC John/O=POC','CN=Leung CC1 John/O=POC','CN=Leung CC2 John/O=POC'],"Python Group")  # User Name Array + Group Name
    Remove Group Member from Group
        n.rpc.notes_service.removeUserFromGroup(['CN=Leung CC John/O=POC','CN=Leung CC1 John/O=POC','CN=Leung CC2 John/O=POC'],"Python Group")
    Delete Group
        n.rpc.notes_service.removeGroup("Python Group")