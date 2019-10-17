import os
path = os.path.dirname(os.path.abspath(__file__))
print(os.path.join(path,'uploads'))
print(os.path.join(path,'downloads'))
print(os.path.join(path,'temp'))