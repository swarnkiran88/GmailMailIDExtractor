import imaplib, email, os 
import pandas as pd
 
def split_addrs(s):
    #split an address list into list of tuples of (name,address)
    if not(s): return []
    outQ = True
    cut = -1
    res = []
    for i in range(len(s)):
        if s[i]=='"': outQ = not(outQ)
        if outQ and s[i]==',':
            res.append(email.utils.parseaddr(s[cut+1:i]))
            cut=i
    res.append(email.utils.parseaddr(s[cut+1:i+1]))
    return res
 
mail=imaplib.IMAP4_SSL('imap.gmail.com')
mail.login('EMAILID','PASSWORD')
mail.select("INBOX")
result,data=mail.search(None,"ALL")
ids=data[0].split()
#print(ids)
msgs = mail.fetch(b",".join(ids),'(BODY.PEEK[HEADER])')[1][0::2]
#print(msgs)
addr=[]
for x,msg in msgs:
    msgobj = email.message_from_string(msg.decode('utf-8'))
    addr.extend(split_addrs(msgobj['to']))
    addr.extend(split_addrs(msgobj['from']))
    addr.extend(split_addrs(msgobj['cc']))
print(addr)
df1_header= ["Name","Email"]
df1=pd.DataFrame(addr,columns=df1_header)
print(type(addr))
print(df1)
print(df1.drop_duplicates(subset=["Email"], keep='first'))
df1.to_csv(os.getcwd()+"/email.csv",sep=',',index=False)
