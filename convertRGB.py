import re, sys
rgbRegex = re.compile(r'(\d{1,3}), (\d{1,3}), (\d{1,3})')

try:
    fileName = sys.argv[1]
except IndexError:
    print("Provide a file name!")
    quit()

try:
    file = open(fileName)
except FileNotFoundError:
    print("File not found!")
    quit()

newFile = open('(new)' + fileName.split('/')[-1], "w+")
lines = file.readlines()
for line in lines:
    result = rgbRegex.search(line)
    if result == None:
        newFile.write(line)
    else:
        converted = ""
        for i in range(1, 4):
            converted += str(int(result.group(i)) / 255.0)
            if i != 3:
                converted += ', '
        newFile.write(line.replace(result.group(), converted))
