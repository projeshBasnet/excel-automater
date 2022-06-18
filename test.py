def range_char(start, stop):
    x = (chr(n) for n in range(ord(start), ord(stop) + 1))
    print(f"x: {x}")
    return x
        
# Example run
for character in range_char("a", "g"):
    print(character)