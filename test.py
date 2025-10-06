import os

os.system("cls")

tags_list = ["%P30NGP101", "%_030FZL120A", "%P01XA942", "%_030XZS058A"]
new_tags_list = []
pos_list = []
for t in tags_list:
    start_idx = 1 + (t[1] == "_")
    last_part_beginning_index = -3 - (t[-1].isalpha())

    nt = t[start_idx:]
    nt = nt[:3] + "-" + nt[3:last_part_beginning_index] + "-" + nt[last_part_beginning_index:]

    new_tags_list.append(nt)

print(new_tags_list)
print(pos_list)