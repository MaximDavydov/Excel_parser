a = [1,2,3]
b = [4,5,6]
c = [7,8,9,0]

for i, j, k in zip(a, b, c):
  print(i, '+', j, '+', k)

m = map(str.strip, 'zx')
print(list(m))