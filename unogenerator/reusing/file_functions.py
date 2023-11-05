## THIS IS FILE IS FROM https://github.com/turulomio/reusingcode/python/file_functions.py
## IF YOU NEED TO UPDATE IT PLEASE MAKE A PULL REQUEST IN THAT PROJECT AND DOWNLOAD FROM IT
## DO NOT UPDATE IT IN YOUR CODE


## @brief file_functions. Functions used to manipulate text in files

from os import remove


def replace_in_file(filename, s, r):
    data=open(filename,"r").read()
    remove(filename)
    data=data.replace(s,r)
    f=open(filename, "w")
    f.write(data)
    f.close()


## Replace a line in file that contains. Must finish with a \n or change of line
def replace_line_in_file_that_contains(filename, s, r):
    n=""
    with open(filename, "r") as f:
        for line in f.readlines():
            if s in line:
                n=n+r
            else:
                n=n+line
    remove(filename)
    with open(filename, "w") as f:
        f.write(n)



if __name__ == "__main__":
    with open("text.txt", "w") as f:
        f.write("This ia a text\n")
        f.write("This is another text\n")
        f.write("This is poetry\n")
    
    replace_in_file("text.txt", "text", "good text")
    replace_line_in_file_that_contains("text.txt", "another", "This is not another text\n")
