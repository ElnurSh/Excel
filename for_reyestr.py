import pandas as pd

data = pd.read_excel('~/Desktop/NEZ-Reyestr-İyun-2022-01-.xlsx', sheet_name=None, usecols = "A,D,G,J", header = 5, nrows=0)
text = []
sheetnames = list(data.keys())
for name in sheetnames:
    text += list(data[name])
print(len(text))
#print(text)
print(list(set([x for x in text if text.count(x) > 1])))
text1 = list(set([x for x in text if text.count(x) > 1]))
print(len(list(set([x for x in text if text.count(x) > 1]))))

df = pd.DataFrame(text1)
df.to_excel('~/Desktop/NEzzz.xlsx', sheet_name='welcome', index=False)
