import sys

with open("app.py", "r") as f:
    lines = f.readlines()

deletion_start = 2768 - 1
deletion_end = 2823 - 1

block_to_move = "".join(lines[deletion_start:deletion_end+1])

print("BLOCK LENGTH:", len(block_to_move))
print("FIRST LINE:", repr(lines[deletion_start]))
print("LAST LINE:", repr(lines[deletion_end]))

# Now let's just create a modified python script safely
insertion_index = 2707 - 1

new_lines = lines[:insertion_index] + [block_to_move + "\n"] + lines[insertion_index:deletion_start] + lines[deletion_end+1:]

with open("app_modified.py", "w") as f:
    f.writelines(new_lines)
