import re

test = "4€May€2020\n7077852"
re_param2 = "[\d]{1,2}€[A-Za-z]{3}€[\d]{4}\\n[\d]{7}"
search = re.findall(re_param2, test)
print(search)
