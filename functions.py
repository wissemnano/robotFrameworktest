import string
def funcname(*list):
    for target_list in list:
        print (target_list)

if __name__ == "__main__":
    myIntList   =   ['0' , '1' , '2' , '3' , '4' , '5','6', '7' , '8' , '9']
    funcname("test", 'b' ,'c')
    listAlpha   =   string.ascii_uppercase
    print (listAlpha[0])
    if '1' in myIntList:
        print ("val in list")