import os

hostname = "google.com"

def ipstat(hostname)
    response = os.system("ping -c 1 " + hostname)
    if response == 0:
        eqsts = 'up'
    else:
        eqsts = 'down'
    return eqsts
