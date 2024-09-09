text = "A(-1, 2.5, 0.5, 0.0)"

textlst = text[1:].replace("(", "").replace(")", "").split(",")
# textlst = shrinked.split(",")

converted = [float(i) for i in textlst]

# print(shrinked)
# print(textlst)
print(converted)
