b0 = 3
b1 = 4
b2 = 5
b3 = 6


tm_reg_ind = [b0, b1, b0, b0, b1]
b0 = 6
print(tm_reg_ind)
tm_reg_ind = [b0, b1, b0, b0, b1]
print(tm_reg_ind)

tm_reg_ind[2] = 30

print(b0)
