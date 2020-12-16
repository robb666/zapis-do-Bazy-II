import re


s = '90-349'
if re.search('\d{2}[-|\xad]\d{3}', s):
    print(s)