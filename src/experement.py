 # any: Return True if bool(x) is True for any x in the iterable (dicts, lists, sets, etc.)
    # If the iterable is empty, return False.
    # dict is a collection which is unordered, changeable and indexed. In Python dictionaries are written with curly brackets, and they have keys and values.
    # list is a collection which is ordered and changeable. Allows duplicate members.
    # set is a collection which is unordered and unindexed. No duplicate members.
    # e.g ['a', 'b', 'c'] 
    # if we say any(x == 'a' for x in ['a', 'b', 'c']) it will return True

list = [1, 3, 6]

if any(x == 7 for x in list):
    print("we have 3")
else: print("nothing is found")

list = [3, 3, 3]

if all(x == 3 for x in list):
    print("all numbers are 3")
else: print('not all numbers are 3')

if any(all(x == 3 for x in list)):
    print('?')
else: print("!")