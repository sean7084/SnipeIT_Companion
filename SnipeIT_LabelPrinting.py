import socket     
host1 = "192.168.31.110" 
port = 9100   

def printbrandlocationlabel(companyname,sitelocation):
    mysocket = socket.socket(socket.AF_INET,socket.SOCK_STREAM)
    mysocket.connect((host1, port)) #connecting to host
    ZPLCommend = """
        ^XA
        ^FX Brand + User or Store Name
        ^FO180,95^A0N,100^FB827,2,0,C^FD{}\&{}^FS
        ^XZ""".format(companyname,sitelocation)
    mysocket.send(bytes(ZPLCommend,encoding='utf8'))#using bytes
    mysocket.close () #closing connection

def printkeringassettaglabel(keringassettag):
    mysocket = socket.socket(socket.AF_INET,socket.SOCK_STREAM)
    mysocket.connect((host1, port)) #connecting to host
    ZPLCommend = """
        ^XA
        ^FX Code 128 barcode for kering asset tag
        ^FO180,95^BCn,120,y,n,y,n^FD{}^FS
        ^XZ""".format(keringassettag)
    mysocket.send(bytes(ZPLCommend,encoding='utf8'))#using bytes
    mysocket.close () #closing connection